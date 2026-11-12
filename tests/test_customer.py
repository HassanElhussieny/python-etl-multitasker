"""Task 3: Customer - sales, returns and net per customer code (columns B, F, G, T)."""

from datetime import date
from decimal import Decimal

from openpyxl import load_workbook

from sales_etl.cli import EXIT_AUDIT_FAILED, EXIT_NO_INPUT, EXIT_OK, EXIT_USAGE, main
from sales_etl.config import PipelineConfig
from sales_etl.customer import CUSTOMER_COLUMNS, run_customer_totals
from tests.conftest import ERP_HEADER, R, S, Txn, alex8_rows, erp_csv_text, write_erp_csv

D = Decimal

# Txn(date, type (B), net (T), doc_no, customer code (F)); the name (G) is "عميل <code>"
ROWS = [
    Txn("01-10-2026", S, "100.10", "1", "400100"),
    Txn("15-10-2026", S, "200.20", "2", "400100"),
    Txn("20-10-2026", R, "-30.00", "3", "400100"),
    Txn("05-09-2025", S, "50.00", "4", "400200"),     # any date counts: no date filter
    Txn("31-12-2026", R, "-7.25", "5", "400200"),
    Txn("10-10-2026", R, "2.00", "6", "400200"),      # unusual positive return: sign kept
    Txn("11-10-2026", R, "-9.99", "7", "400300"),     # returns only
    Txn("12-10-2026", "تسوية", "1000.00", "8", "400100"),  # neither sales nor returns: not counted
]


def _config(tmp_path):
    return PipelineConfig(input_dir=tmp_path / "input", output_dir=tmp_path / "output",
                          reference_date=date(2026, 10, 1))


def _rows(path):
    return [list(r) for r in load_workbook(path, data_only=True)["التقرير"].iter_rows(values_only=True)]


def test_columns_are_exactly_the_five_requested():
    assert CUSTOMER_COLUMNS == ("كود العميل", "اسم العميل", "إجمالي المبيعات", "إجمالي المرتجعات", "الصافي")


def test_totals_per_customer(tmp_path):
    write_erp_csv(tmp_path / "input" / "month.csv", ROWS)
    result = run_customer_totals(_config(tmp_path))
    rows = _rows(result.report_path)
    assert rows[0] == list(CUSTOMER_COLUMNS)
    assert rows[1:] == [
        ["400100", "عميل 400100", 300.3, -30, 270.3],
        ["400200", "عميل 400200", 50, -5.25, 44.75],
        ["400300", "عميل 400300", 0, -9.99, -9.99],
    ]
    for c in result.groups.values():
        assert c.net == c.sales + c.returns
    assert result.complete


def test_columns_by_position_b_f_g_t(tmp_path):
    """Values are taken from B (type), F (code), G (name) and T (صافى) of the ERP layout."""
    assert [ERP_HEADER[i] for i in (1, 5, 19)] == ["نوع المستند", "العميل", "صافى"] and ERP_HEADER[6] == ""
    row = ["01-10-2026", S, "9", "93", "", "777", "اسم في G", "", "2", "", "1", *["0"] * 8, "12.34", ""]
    assert len(row) == 21
    write_erp_csv(tmp_path / "input" / "m.csv", [row])
    (only,) = run_customer_totals(_config(tmp_path)).groups.values()
    assert (only.code, only.name, only.sales, only.returns) == ("777", "اسم في G", D("12.34"), D("0"))


def test_each_customer_once_across_files(tmp_path):
    write_erp_csv(tmp_path / "input" / "a.csv", ROWS[:3])
    write_erp_csv(tmp_path / "input" / "b.csv", [Txn("02-10-2026", S, "1.00", "9", "400100")])
    rows = _rows(run_customer_totals(_config(tmp_path)).report_path)
    assert [r[0] for r in rows[1:]] == ["400100"] and rows[1][2:] == [301.3, -30, 271.3]


def test_real_alex8_content_net_equals_erp_footer(tmp_path):
    """All 235 real transactions: the sum of every customer's net = the file's ERP صافى footer."""
    write_erp_csv(tmp_path / "input" / "alex8.csv", alex8_rows(), footer_net=D("1105646.62"), footer_count=235)
    result = run_customer_totals(_config(tmp_path))
    assert sum((c.net for c in result.groups.values()), D(0)) == D("1105646.62")
    rows = _rows(result.report_path)
    assert len({r[0] for r in rows[1:]}) == len(rows) - 1                    # one row per code


def test_menu_choice_3_moves_file_to_processed_unchanged(tmp_path, capsys):
    path = write_erp_csv(tmp_path / "in" / "month.csv", ROWS)
    before = path.read_bytes()
    answers = ["3"]                                                          # menu: Customer, no date asked
    asked = []

    def fake_input(prompt):
        asked.append(prompt)
        return answers.pop(0)

    code = main(["--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "out")], input_fn=fake_input)
    assert code == EXIT_OK and asked == ["Select task number: "]
    assert "3. Customer" in capsys.readouterr().out
    assert list((tmp_path / "in").iterdir()) == []
    (folder,) = (tmp_path / "out").iterdir()
    assert folder.name.startswith("Customer_") and [p.name for p in folder.iterdir()] == ["customer_totals.xlsx"]
    (archive,) = (tmp_path / "processed").iterdir()
    assert archive.name == folder.name and (archive / "month.csv").read_bytes() == before


def test_no_move_leaves_input_file_unchanged(tmp_path):
    path = write_erp_csv(tmp_path / "in" / "month.csv", ROWS)
    before = path.read_bytes()
    code = main(["--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "out"),
                 "--task", "3", "--no-move", "--non-interactive"])
    assert code == EXIT_OK
    assert path.read_bytes() == before and [p.name for p in (tmp_path / "in").iterdir()] == ["month.csv"]
    for folder in ("processed", "failures", "duplicates", "state"):
        assert not (tmp_path / folder).exists()


def test_task_argument_and_errors(tmp_path):
    (tmp_path / "in").mkdir()
    args = ["--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "out"), "--non-interactive"]
    assert main([*args, "--task", "3"]) == EXIT_NO_INPUT
    assert main([*args, "--task", "3", "--ref-date", "01/10/2026"]) == EXIT_USAGE
    (tmp_path / "in" / "empty.csv").write_bytes(b"")
    write_erp_csv(tmp_path / "in" / "ok.csv", ROWS)
    assert main([*args, "--task", "3"]) == EXIT_AUDIT_FAILED                  # failed file reported


def test_utf8_file_with_comma_delimiter(tmp_path):
    (tmp_path / "input").mkdir()
    (tmp_path / "input" / "u.csv").write_text(erp_csv_text(ROWS[:3], delimiter=","), encoding="utf-8-sig")
    (c,) = run_customer_totals(_config(tmp_path)).groups.values()
    assert (c.sales, c.returns, c.net) == (D("300.30"), D("-30.00"), D("270.30"))
