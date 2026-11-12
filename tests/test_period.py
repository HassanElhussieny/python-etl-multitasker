"""Task 2: إجمالي المبيعات والمرتجعات لفترة محددة (start and end days included)."""

from datetime import date
from decimal import Decimal

from openpyxl import load_workbook

from sales_etl.cli import EXIT_AUDIT_FAILED, EXIT_NO_INPUT, EXIT_OK, EXIT_USAGE, main
from sales_etl.config import GRAND_TOTAL_LABEL, PipelineConfig
from sales_etl.period import PERIOD_COLUMNS, run_period_totals
from tests.conftest import R, S, Txn, outcome, write_erp_csv

D = Decimal

# Period 07-10-2026 .. 12-10-2026
FILE_A = [
    Txn("06-10-2026", S, "1000.00", "1"),                                       # before start: out
    Txn("07-10-2026", S, "100.10", "2"), Txn("07-10-2026", R, "-10.00", "3"),   # start day: in
    Txn("09-10-2026", S, "200.20", "4"), Txn("10-10-2026", R, "-5.05", "5"),    # inside
    Txn("12-10-2026", S, "50.00", "6"), Txn("12-10-2026", R, "-1.00", "7"),     # end day: in
    Txn("13-10-2026", S, "999.99", "8"), Txn("13-10-2026", R, "-88.00", "9"),   # after end: out
]
# sales 350.30, returns -16.05, net 334.25
FILE_B = [
    Txn("08-10-2026", S, "0.10", "1"), Txn("11-10-2026", R, "-0.20", "2"),
    Txn("11-10-2026", R, "3.00", "3"),                                          # positive return kept
    Txn("01-10-2026", R, "-50.00", "4"),                                        # before start: out
]
# sales 0.10, returns 2.80, net 2.90


def _config(tmp_path):
    return PipelineConfig(input_dir=tmp_path / "input", output_dir=tmp_path / "output",
                          reference_date=date(2026, 10, 12))


def _rows(path):
    return [list(r) for r in load_workbook(path, data_only=True)["التقرير"].iter_rows(values_only=True)]


def test_period_includes_start_and_end_days(tmp_path):
    write_erp_csv(tmp_path / "input" / "a.csv", FILE_A)
    result = run_period_totals(_config(tmp_path), date(2026, 10, 7), date(2026, 10, 12))
    (a,) = result.files
    assert (a.sales, a.returns, a.net) == (D("350.30"), D("-16.05"), D("334.25"))
    assert a.rows_in_period == 6 and result.complete


def test_values_per_file_and_grand_total_in_report(tmp_path):
    write_erp_csv(tmp_path / "input" / "a.csv", FILE_A)
    write_erp_csv(tmp_path / "input" / "b.csv", FILE_B)
    result = run_period_totals(_config(tmp_path), date(2026, 10, 7), date(2026, 10, 12))
    assert (result.sales, result.returns, result.net) == (D("350.40"), D("-13.25"), D("337.15"))
    assert result.net == result.sales + result.returns
    assert result.report_path.parent.name.startswith("Period_")
    rows = _rows(result.report_path)
    assert rows[0] == list(PERIOD_COLUMNS) == ["اسم الملف", "إجمالي المبيعات", "إجمالي المرتجعات", "الصافي"]
    assert rows[1:] == [
        ["a.csv", 350.3, -16.05, 334.25],
        ["b.csv", 0.1, 2.8, 2.9],
        [GRAND_TOTAL_LABEL, 350.4, -13.25, 337.15],
    ]


def test_single_day_period(tmp_path):
    write_erp_csv(tmp_path / "input" / "a.csv", FILE_A)
    (a,) = run_period_totals(_config(tmp_path), date(2026, 10, 12), date(2026, 10, 12)).files
    assert (a.sales, a.returns, a.net) == (D("50.00"), D("-1.00"), D("49.00"))


def test_same_parsing_as_task_1(tmp_path, run):
    """Task 2 over (-inf, ref] equals task 1's yesterday + previous days for the same file."""
    write_erp_csv(tmp_path / "input" / "a.csv", FILE_A)
    m = outcome(run("12/10/2026"), "a.csv").metrics
    (a,) = run_period_totals(_config(tmp_path), date(2000, 1, 1), date(2026, 10, 12)).files
    assert a.sales == m.sales_yesterday + m.sales_previous
    assert a.returns == m.returns_yesterday + m.returns_previous


def test_no_move_leaves_input_files_untouched(tmp_path):
    path = write_erp_csv(tmp_path / "in" / "a.csv", FILE_A)
    before = path.read_bytes()
    code = main(["--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "out"),
                 "--start-date", "07-10-2026", "--end-date", "12-10-2026", "--non-interactive", "--no-move"])
    assert code == EXIT_OK
    assert path.read_bytes() == before and [p.name for p in (tmp_path / "in").iterdir()] == ["a.csv"]
    for folder in ("processed", "failures", "duplicates", "state"):
        assert not (tmp_path / folder).exists()
    assert [p.name.startswith("Period_") for p in (tmp_path / "out").iterdir()] == [True]


def test_menu_choice_2_asks_start_then_end(tmp_path, capsys):
    write_erp_csv(tmp_path / "in" / "a.csv", FILE_A)
    answers = ["2", "7-10-2026", "05-10-2026", "12-10-2026"]   # end before start is asked again
    asked = []

    def fake_input(prompt):
        asked.append(prompt)
        return answers.pop(0)

    code = main(["--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "out")], input_fn=fake_input)
    assert code == EXIT_OK
    assert asked == ["Select task number: ", "اكتب تاريخ البداية (DD-MM-YYYY): ",
                     "اكتب تاريخ النهاية (DD-MM-YYYY): ", "اكتب تاريخ النهاية (DD-MM-YYYY): "]
    out = capsys.readouterr().out
    assert "Available Tasks:\n1. Daily\n2. Period\n3. Customer\n4. Representative\n5. Sector & Representative Sales Analysis\n" in out
    assert "قبل تاريخ البداية" in out
    assert "إجمالي المبيعات" in out and "350.30" in out and "-16.05" in out and "334.25" in out
    assert "اكتب تاريخ أمس" not in "".join(asked)          # task 1 did not run
    assert [p.name.startswith("Period_") for p in (tmp_path / "processed").iterdir()] == [True]
    assert not list((tmp_path / "out").glob("Report_*"))


def test_invalid_menu_choice_is_asked_again(tmp_path):
    write_erp_csv(tmp_path / "in" / "a.csv", FILE_A)
    answers = ["9", "", "x", "1", "12/10/2026"]
    code = main(["--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "out")],
                input_fn=lambda prompt: answers.pop(0))
    assert code == EXIT_OK and answers == []
    assert [p.name for p in (tmp_path / "processed").iterdir()][0].startswith("Run_")  # task 1 ran


def test_end_before_start_argument_is_usage_error(tmp_path):
    (tmp_path / "in").mkdir()
    assert main(["--input-dir", str(tmp_path / "in"), "--start-date", "12-10-2026",
                 "--end-date", "07-10-2026", "--non-interactive"]) == EXIT_USAGE


def test_mixing_task_arguments_is_usage_error(tmp_path):
    (tmp_path / "in").mkdir()
    assert main(["--input-dir", str(tmp_path / "in"), "--ref-date", "12/10/2026",
                 "--start-date", "07-10-2026", "--non-interactive"]) == EXIT_USAGE


def test_no_input_files(tmp_path):
    (tmp_path / "in").mkdir()
    code = main(["--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "out"),
                 "--start-date", "07-10-2026", "--end-date", "12-10-2026", "--non-interactive"])
    assert code == EXIT_NO_INPUT and not (tmp_path / "out").exists()


def test_failed_file_is_reported_and_excluded(tmp_path):
    write_erp_csv(tmp_path / "in" / "a.csv", FILE_A)
    (tmp_path / "in" / "empty.csv").write_bytes(b"")
    code = main(["--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "out"),
                 "--start-date", "07-10-2026", "--end-date", "12-10-2026", "--non-interactive"])
    assert code == EXIT_AUDIT_FAILED
    (folder,) = (tmp_path / "out").iterdir()
    rows = _rows(next(folder.glob("*.xlsx")))
    assert [r[0] for r in rows[1:]] == ["a.csv", GRAND_TOTAL_LABEL] and rows[-1][3] == 334.25
    assert list((tmp_path / "in").iterdir()) == []
    assert [p.name for p in (tmp_path / "failures").iterdir()] == ["empty.csv"]       # not in processed
    (archive,) = (tmp_path / "processed").iterdir()
    assert archive.name == folder.name and [p.name for p in archive.iterdir()] == ["a.csv"]


# ------------------------------------------------------------------ file lifecycle
def _period(tmp_path, *extra):
    return main(["--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "out"),
                 "--start-date", "07-10-2026", "--end-date", "12-10-2026", "--non-interactive", *extra])


def _grand(tmp_path):
    reports = sorted((tmp_path / "out").glob("Period_*/*.xlsx"))
    return [_rows(p)[-1] for p in reports]


def test_processed_files_move_to_processed_period_folder(tmp_path):
    a = write_erp_csv(tmp_path / "in" / "a.csv", FILE_A)
    write_erp_csv(tmp_path / "in" / "b.csv", FILE_B)
    original = a.read_bytes()
    assert _period(tmp_path) == EXIT_OK
    (report_dir,) = (tmp_path / "out").iterdir()
    (archive,) = (tmp_path / "processed").iterdir()
    assert report_dir.name.startswith("Period_") and archive.name == report_dir.name   # same stamp
    assert sorted(p.name for p in archive.iterdir()) == ["a.csv", "b.csv"]
    assert (archive / "a.csv").read_bytes() == original                                # content unchanged
    assert list((tmp_path / "in").iterdir()) == []
    assert _grand(tmp_path) == [[GRAND_TOTAL_LABEL, 350.4, -13.25, 337.15]]           # report still correct
    events = [__import__("json").loads(l)["event"]
              for l in (tmp_path / "state" / "period_journal.jsonl").read_text(encoding="utf-8").splitlines()]
    assert events == ["RUN_STARTED", "COMMIT", "MOVED", "MOVED", "RUN_FINISHED"]
    assert not (tmp_path / "state" / "processing_journal.jsonl").exists()             # Daily's journal untouched


def test_rejected_rows_file_goes_to_failures(tmp_path):
    write_erp_csv(tmp_path / "in" / "bad.csv", [*FILE_A, Txn("xx-10-2026", S, "5.00", "99")])
    assert _period(tmp_path) == EXIT_AUDIT_FAILED
    assert [p.name for p in (tmp_path / "failures").iterdir()] == ["bad.csv"]
    assert not (tmp_path / "processed").exists() or not any((tmp_path / "processed").rglob("*.csv"))


def test_rerun_with_same_file_counts_it_again(tmp_path):
    a = write_erp_csv(tmp_path / "in" / "a.csv", FILE_A)
    content = a.read_bytes()
    assert _period(tmp_path) == EXIT_OK
    (tmp_path / "in" / "a.csv").write_bytes(content)                 # same file supplied again
    (tmp_path / "in" / "a_copy.csv").write_bytes(content)            # ... and under another name
    write_erp_csv(tmp_path / "in" / "b.csv", FILE_B)
    assert _period(tmp_path) == EXIT_OK
    first, second = _grand(tmp_path)
    assert first == [GRAND_TOTAL_LABEL, 350.3, -16.05, 334.25]
    assert second == [GRAND_TOTAL_LABEL, 700.7, -29.3, 671.4]        # a + a_copy + b: all counted
    assert not (tmp_path / "duplicates").exists()
    assert list((tmp_path / "in").iterdir()) == []
    archived = sorted(p.name for p in (tmp_path / "processed").rglob("*.csv"))
    assert archived == ["a.csv", "a.csv", "a_copy.csv", "b.csv"]     # all moved, nothing overwritten


def test_identical_files_in_one_run_both_counted(tmp_path):
    a = write_erp_csv(tmp_path / "in" / "a.csv", FILE_A)
    (tmp_path / "in" / "z.csv").write_bytes(a.read_bytes())
    assert _period(tmp_path) == EXIT_OK
    assert _grand(tmp_path) == [[GRAND_TOTAL_LABEL, 700.6, -32.1, 668.5]]
    assert sorted(p.name for p in (tmp_path / "processed").rglob("*.csv")) == ["a.csv", "z.csv"]


def test_repeated_rows_in_period_are_all_counted(tmp_path):
    same = Txn("09-10-2026", S, "10.00", "4")
    write_erp_csv(tmp_path / "in" / "a.csv", [same, same, same])          # same document number x3
    assert _period(tmp_path) == EXIT_OK
    assert _grand(tmp_path) == [[GRAND_TOTAL_LABEL, 30, 0, 30]]


def test_daily_history_does_not_block_period_and_vice_versa(tmp_path):
    a = write_erp_csv(tmp_path / "in" / "a.csv", FILE_A)
    content = a.read_bytes()
    assert main(["--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "out"),
                 "--ref-date", "12/10/2026", "--non-interactive"]) == EXIT_OK          # Daily
    (tmp_path / "in" / "a.csv").write_bytes(content)
    assert _period(tmp_path) == EXIT_OK                                                 # Period counts it
    assert _grand(tmp_path) == [[GRAND_TOTAL_LABEL, 350.3, -16.05, 334.25]]
    assert not (tmp_path / "duplicates").exists()
    (tmp_path / "in" / "a.csv").write_bytes(content)
    assert main(["--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "out"),
                 "--ref-date", "12/10/2026", "--non-interactive"]) == EXIT_OK          # Daily as before
    assert len(list((tmp_path / "processed").glob("Run_*"))) == 2

