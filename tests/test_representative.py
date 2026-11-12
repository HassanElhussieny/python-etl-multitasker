"""Task 4: Representative - sales, returns and net per representative code (columns B, I, J, T)."""

from datetime import date
from decimal import Decimal

from openpyxl import load_workbook

from sales_etl.cli import EXIT_OK, EXIT_USAGE, main
from sales_etl.config import PipelineConfig
from sales_etl.representative import REPRESENTATIVE_COLUMNS, run_representative_totals
from tests.conftest import ERP_HEADER, R, S, alex8_rows, write_erp_csv

D = Decimal


def row(kind, net, rep, rep_name, customer="400100", doc="1", day="01-10-2026"):
    """One ERP row (21 cells): A date, B type, F/G customer, I/J representative, T صافى."""
    return [day, kind, doc, "93", "", customer, f"عميل {customer}", "", rep, rep_name, net,
            *["0"] * 8, net, ""]


ROWS = [
    row(S, "100.10", "2", "مندوب 2", customer="400100", doc="1"),
    row(S, "200.20", "2", "مندوب 2", customer="400999", doc="2"),   # other customer, same rep
    row(R, "-30.00", "2", "مندوب 2", doc="3"),
    row(S, "50.00", "7", "مندوب 7", doc="4", day="05-09-2025"),     # any date counts
    row(R, "-7.25", "7", "مندوب 7", doc="5"),
    row(R, "2.00", "7", "مندوب 7", doc="6"),                         # unusual positive return: sign kept
    row(R, "-9.99", "12", "مندوب 12", doc="7"),                      # returns only
    row("تسوية", "1000.00", "2", "مندوب 2", doc="8"),                 # unrecognised type: not counted
]


def _config(tmp_path):
    return PipelineConfig(input_dir=tmp_path / "input", output_dir=tmp_path / "output",
                          reference_date=date(2026, 10, 1))


def _rows(path):
    return [list(r) for r in load_workbook(path, data_only=True)["التقرير"].iter_rows(values_only=True)]


def test_columns_are_exactly_the_five_requested():
    assert REPRESENTATIVE_COLUMNS == ("كود المندوب", "اسم المندوب", "إجمالي المبيعات", "إجمالي المرتجعات", "الصافي")
    assert [ERP_HEADER[i] for i in (1, 8, 19)] == ["نوع المستند", "المندوب", "صافى"] and ERP_HEADER[9] == ""


def test_totals_per_representative(tmp_path):
    write_erp_csv(tmp_path / "input" / "month.csv", ROWS)
    result = run_representative_totals(_config(tmp_path))
    rows = _rows(result.report_path)
    assert rows[0] == list(REPRESENTATIVE_COLUMNS)
    assert rows[1:] == [                                                    # sorted by code: 2, 7, 12
        ["2", "مندوب 2", 300.3, -30, 270.3],                               # grouped across customers
        ["7", "مندوب 7", 50, -5.25, 44.75],                                # -7.25 + 2.00
        ["12", "مندوب 12", 0, -9.99, -9.99],
    ]
    for g in result.groups.values():
        assert g.net == g.sales + g.returns
    assert result.complete and result.report_path.parent.name.startswith("Representative_")
    assert result.report_path.name == "representative_totals.xlsx"


def test_customer_columns_are_not_used(tmp_path):
    write_erp_csv(tmp_path / "input" / "m.csv", [row(S, "5.00", "3", "مندوب 3", customer="111"),
                                                 row(S, "6.00", "3", "مندوب 3", customer="222", doc="2")])
    (only,) = run_representative_totals(_config(tmp_path)).groups.values()
    assert (only.code, only.name, only.sales) == ("3", "مندوب 3", D("11.00"))


def test_multiple_files_combined_once_per_code(tmp_path):
    write_erp_csv(tmp_path / "input" / "a.csv", ROWS[:3])
    write_erp_csv(tmp_path / "input" / "b.csv", [row(S, "1.00", "2", "مندوب 2", doc="9")])
    rows = _rows(run_representative_totals(_config(tmp_path)).report_path)
    assert [r[0] for r in rows[1:]] == ["2"] and rows[1][2:] == [301.3, -30, 271.3]


def test_same_code_with_different_names_is_one_row_and_reported(tmp_path, capsys):
    write_erp_csv(tmp_path / "in" / "a.csv", [row(S, "1.00", "5", "أحمد")])
    write_erp_csv(tmp_path / "in" / "b.csv", [row(S, "2.00", "5", "احمد على", doc="2")])
    code = main(["--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "out"),
                 "--task", "4", "--non-interactive", "--no-move"])
    assert code == EXIT_OK
    (report,) = (tmp_path / "out").glob("Representative_*/representative_totals.xlsx")
    assert _rows(report)[1:] == [["5", "أحمد", 3, 0, 3]]                  # first name found, one row
    assert "WARNING code 5 has 2 names" in capsys.readouterr().out


def test_real_alex8_content_net_equals_erp_footer(tmp_path):
    write_erp_csv(tmp_path / "input" / "alex8.csv", alex8_rows(), footer_net=D("1105646.62"), footer_count=235)
    result = run_representative_totals(_config(tmp_path))
    assert sum((g.net for g in result.groups.values()), D(0)) == D("1105646.62")


def test_menu_choice_4_and_date_arguments_rejected(tmp_path, capsys):
    write_erp_csv(tmp_path / "in" / "month.csv", ROWS)
    answers = ["4"]
    code = main(["--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "out")],
                input_fn=lambda prompt: answers.pop(0))
    assert code == EXIT_OK
    assert "4. Representative" in capsys.readouterr().out
    assert [p.name.startswith("Representative_") for p in (tmp_path / "out").iterdir()] == [True]
    assert main(["--input-dir", str(tmp_path / "in"), "--task", "4", "--start-date", "01-10-2026",
                 "--non-interactive"]) == EXIT_USAGE
