"""CLI behaviour: the reference date is requested once and applied to all files."""

from sales_etl.cli import EXIT_AUDIT_FAILED, EXIT_NO_INPUT, EXIT_OK, EXIT_USAGE, main
from tests.conftest import S, Txn, write_erp_csv


def _prompts(answers):
    asked = []

    def fake_input(prompt):
        asked.append(prompt)
        return answers.pop(0)

    return fake_input, asked


def test_prompts_once_for_date_and_moves_files(tmp_path, capsys):
    for name in ("a.csv", "b.csv", "c.csv"):
        write_erp_csv(tmp_path / "in" / name, [Txn("08-10-2026", S, "1", doc_no=name)])
    fake_input, asked = _prompts(["1", "08/10/2026"])  # task menu, then the date
    code = main(["--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "out")],
                input_fn=fake_input)
    assert code == EXIT_OK
    assert asked == ["Select task number: ", "اكتب تاريخ أمس (DD/MM/YYYY): "]
    out = capsys.readouterr().out
    assert "Thursday 08 October 2026 (ISO 2026-10-08)" in out
    assert "Files processed this run : 3" in out and "Reconciliation audit     : PASS" in out
    assert "RUN RESULT               : SUCCEEDED - report generated" in out
    runs = list((tmp_path / "processed").iterdir())
    assert len(runs) == 1 and runs[0].name.startswith("Run_")
    assert sorted(p.name for p in runs[0].iterdir()) == ["a.csv", "b.csv", "c.csv"]
    assert "Report folder" in out and "[OK] report" in out and "[OK] audit_json" in out
    assert list((tmp_path / "in").iterdir()) == []


def test_no_move_mode_leaves_input_untouched(tmp_path):
    write_erp_csv(tmp_path / "in" / "a.csv", [Txn("08-10-2026", S, "1")])
    code = main(["--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "out"),
                 "--ref-date", "08/10/2026", "--no-move", "--non-interactive"])
    assert code == EXIT_OK
    assert [p.name for p in (tmp_path / "in").iterdir()] == ["a.csv"]
    assert not (tmp_path / "processed").exists() and not (tmp_path / "state").exists()


def test_invalid_dates_are_re_prompted_until_valid(tmp_path, capsys):
    write_erp_csv(tmp_path / "in" / "a.csv", [Txn("18-10-2026", S, "5"), Txn("19-10-2026", S, "7", "2")])
    fake_input, asked = _prompts(["1", "31/02/2026", "2026/10/18", "", "abc", "18/10/2026"])
    code = main(["--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "out")],
                input_fn=fake_input)
    assert code == EXIT_OK and len(asked) == 6
    out = capsys.readouterr().out
    assert out.count("التاريخ غير صحيح") == 4 and "Sunday 18 October 2026" in out
    line = next(ln for ln in out.splitlines() if "إجمالي مبيعات أمس" in ln and "Grand Total" in ln)
    assert line.split()[-1] == "5.00"  # 18/10 counted, 19/10 excluded


def test_cancelled_date_prompt_processes_nothing(tmp_path):
    write_erp_csv(tmp_path / "in" / "a.csv", [Txn("18-10-2026", S, "5")])

    def eof(prompt):
        raise EOFError

    code = main(["--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "out"), "--task", "1"],
                input_fn=eof)
    assert code == EXIT_USAGE
    assert [p.name for p in (tmp_path / "in").iterdir()] == ["a.csv"] and not (tmp_path / "out").exists()


def test_main_py_runs_directly_like_vscode_run_button(tmp_path):
    """`python sales_etl/__main__.py` from an unrelated folder: import works, date asked, files moved."""
    import subprocess
    import sys
    from pathlib import Path

    main_py = Path(__file__).resolve().parents[1] / "sales_etl" / "__main__.py"
    write_erp_csv(tmp_path / "in" / "a.csv", [Txn("18-10-2026", S, "5"), Txn("17-10-2026", S, "3", "2")])
    proc = subprocess.run(
        [sys.executable, str(main_py), "--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "out")],
        input="1\n99/99/2026\n18/10/2026\n", capture_output=True, text=True, encoding="utf-8",
        cwd=tmp_path, env={**__import__("os").environ, "PYTHONIOENCODING": "utf-8"}, timeout=120,
    )
    assert "ModuleNotFoundError" not in proc.stderr, proc.stderr
    assert proc.returncode == 0, proc.stdout + proc.stderr
    assert "Available Tasks:\n1. Daily\n2. Period\n3. Customer\n4. Representative\n5. Sector & Representative Sales Analysis\n" in proc.stdout
    assert "اكتب تاريخ أمس (DD/MM/YYYY):" in proc.stdout and "التاريخ غير صحيح" in proc.stdout
    assert "RUN RESULT               : SUCCEEDED - report generated" in proc.stdout
    (run_folder,) = (tmp_path / "processed").iterdir()
    assert [p.name for p in run_folder.iterdir()] == ["a.csv"]


def test_non_interactive_requires_arguments(tmp_path):
    assert main(["--non-interactive", "--input-dir", str(tmp_path)]) == EXIT_USAGE


def test_bad_date_argument_is_usage_error(tmp_path):
    assert main(["--input-dir", str(tmp_path), "--ref-date", "2026/10/08", "--non-interactive"]) == EXIT_USAGE


def test_audit_failure_sets_exit_code(tmp_path):
    (tmp_path / "in").mkdir()
    (tmp_path / "in" / "empty.csv").write_bytes(b"")
    code = main(["--input-dir", str(tmp_path / "in"), "--ref-date", "08/10/2026",
                 "--output-dir", str(tmp_path / "out"), "--non-interactive"])
    assert code == EXIT_AUDIT_FAILED
    assert [p.name for p in (tmp_path / "failures").iterdir()] == ["empty.csv"]


def test_empty_input_folder_is_not_success(tmp_path, capsys):
    (tmp_path / "in").mkdir()
    code = main(["--input-dir", str(tmp_path / "in"), "--ref-date", "18/10/2026",
                 "--output-dir", str(tmp_path / "out"), "--non-interactive"])
    assert code == EXIT_NO_INPUT
    assert "NO REPORT - there were no CSV files in the input folder" in capsys.readouterr().out
