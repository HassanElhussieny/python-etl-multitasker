"""Built-in practical self-test of the desktop window: ``--self-test <report.json> --sandbox <dir> [--sample-csv <csv>]``.

Works from the launcher (``Python ETL Multitasker.pyw``) and from the packaged .exe, so the packaged build can be
checked, not only the source. It opens the real window and goes through every page and button, measuring times.

Safety: it never starts an ETL job on the configured folders. Reports and History are only READ. "Open report /
folder" are recorded instead of launched (Excel never opens a production report). Settings are saved only to the
sandbox. The background task is a mock (``python -c sleep``). The one real engine run (with ``--sample-csv``) uses
a second window whose output, state and staging folders are all inside the sandbox.
"""

from __future__ import annotations

import json
import os
import sys
import time
import traceback
from pathlib import Path


def run_selftest(argv: list[str]) -> int:
    args = dict(zip(argv[::2], argv[1::2])) if len(argv) % 2 == 0 else {}
    out = Path(args.get("--self-test", "selftest.json"))
    sandbox = Path(args.get("--sandbox", out.parent / "selftest_sandbox"))
    sandbox.mkdir(parents=True, exist_ok=True)
    os.environ["ETL_GUI_SETTINGS"] = str(sandbox / "settings.json")      # never the user's AppData files
    os.environ["ETL_GUI_HISTORY"] = str(sandbox / "jobs_history.jsonl")
    os.environ["ETL_GUI_STAGING"] = str(sandbox / "staging")

    from PySide6.QtCore import QObject, SIGNAL, qInstallMessageHandler
    from PySide6.QtWidgets import QFileDialog, QMessageBox, QPushButton, QTableWidget

    from gui.app import create_application
    from gui.services.jobs import Job, Status
    from gui.services.settings import AppSettings, load_settings
    from gui.widgets import common
    from gui.windows.main_window import PAGES, MainWindow

    report: dict = {"frozen": bool(getattr(sys, "frozen", False)), "executable": sys.executable,
                    "checks": [], "timings_ms": {}, "errors": [], "qt_warnings": []}

    def check(name: str, ok: bool, detail: object = "") -> None:
        report["checks"].append({"name": name, "ok": bool(ok), "detail": str(detail)[:500]})

    sys.excepthook = lambda t, v, tb: report["errors"].append("".join(traceback.format_exception(t, v, tb)))
    qInstallMessageHandler(lambda mode, ctx, msg: report["qt_warnings"].append(str(msg)[:300]))

    opened: list[str] = []

    class _RecordOnly:                                # open_path() records instead of launching Excel/Explorer
        @staticmethod
        def openUrl(url):  # noqa: N802 - Qt name
            opened.append(url.toLocalFile())
            return True

    common.QDesktopServices = _RecordOnly
    app = create_application([])
    settings = load_settings()
    report["settings"] = {"output_dir": settings.output_dir, "state_dir": settings.state_dir,
                          "project_root": settings.project_root, "python": settings.python}
    check("optimization included (fill_table, folder_signature)",
          hasattr(common, "fill_table") and hasattr(common, "folder_signature"))
    try:
        w = MainWindow(settings)
        w.show()
        app.processEvents()
        keys = [k for k, _ in PAGES]

        def nav(key: str) -> float:
            t = time.perf_counter()
            w.nav.setCurrentRow(keys.index(key))
            app.processEvents()
            return (time.perf_counter() - t) * 1000

        def wait(cond, seconds: float) -> bool:
            end = time.monotonic() + seconds
            while time.monotonic() < end:
                app.processEvents()
                if cond():
                    return True
                time.sleep(0.02)
            return cond()

        q = w.queue
        sigs = ("job_changed(QString)", "job_added(QString)", "output(QString,QString)")
        receivers = {s: q.receivers(SIGNAL(s)) for s in sigs}
        widgets = {k: len(p.findChildren(QObject)) for k, p in w.pages.items()}

        # ---- expected data, read independently from the folders
        outdir = Path(settings.output_dir)
        xlsx = [f for d in outdir.iterdir() if d.is_dir() for f in d.glob("*.xlsx")
                if not f.name.startswith("~$")] if outdir.is_dir() else []
        runs = [d for d in outdir.iterdir() if d.is_dir()] if outdir.is_dir() else []

        # ---- navigation: open, leave, return, several rounds
        rounds = []
        for r in range(5):
            row = {}
            for key in ["reports", "history", "monitor", "settings", "run"]:
                row[key] = round(nav(key), 1)
                if w.stack.currentWidget() is not w.pages[key]:
                    check(f"round {r + 1}: {key} is the visible page", False)
            rounds.append(row)
        report["timings_ms"]["navigation_rounds"] = rounds
        check("every click showed the requested page (5 rounds x 5 pages)",
              not [c for c in report["checks"] if c["name"].startswith("round") and not c["ok"]])
        repeat = [v for rnd in rounds[1:] for v in rnd.values()]
        check("repeat visits fast (< 500 ms each)", max(repeat) < 500, f"max {max(repeat):.1f} ms")
        check("reports page lists every report in output", w.pages["reports"].table.rowCount() == len(xlsx),
              f"{w.pages['reports'].table.rowCount()} rows / {len(xlsx)} files")
        check("history page lists every run folder", w.pages["history"].table.rowCount() == len(runs),
              f"{w.pages['history'].table.rowCount()} rows / {len(runs)} folders")

        # ---- Reports page controls
        nav("reports")
        rp = w.pages["reports"]
        t = time.perf_counter()
        [b for b in rp.findChildren(QPushButton) if b.text() == "تحديث"][0].click()
        app.processEvents()
        report["timings_ms"]["reports_refresh_button"] = round((time.perf_counter() - t) * 1000, 1)
        check("reports refresh button", rp.table.rowCount() == len(xlsx))
        if xlsx:
            term = Path(rp.table.item(0, 0).text()).stem[:6]
            rp.search.setText(term)
            app.processEvents()
            shown = [r for r in range(rp.table.rowCount()) if not rp.table.isRowHidden(r)]
            expect = [r for r in range(rp.table.rowCount())
                      if term.lower() in " ".join(rp.table.item(r, c).text().lower() for c in range(3))]
            check("reports search filters rows", shown == expect and len(shown) >= 1, f"'{term}': {len(shown)} rows")
            rp.search.setText("")
            app.processEvents()
            check("reports search cleared shows all", all(not rp.table.isRowHidden(r) for r in range(rp.table.rowCount())))
            rp.table.setCurrentCell(0, 0)
            path = rp.table.item(0, 0).data(256)
            [b for b in rp.findChildren(QPushButton) if b.text() == "فتح التقرير"][0].click()
            [b for b in rp.findChildren(QPushButton) if b.text() == "فتح المجلد"][0].click()
            check("reports open report / open folder (recorded, not launched)",
                  opened[-2:] == [str(Path(path)).replace("\\", "/"), str(Path(path).parent).replace("\\", "/")]
                  or [Path(x) for x in opened[-2:]] == [Path(path), Path(path).parent], opened[-2:])

        # ---- History page controls
        nav("history")
        hp = w.pages["history"]
        [b for b in hp.findChildren(QPushButton) if b.text() == "تحديث"][0].click()
        app.processEvents()
        check("history refresh button", hp.table.rowCount() == len(runs))
        if runs:
            hp.table.setCurrentCell(0, 0)
            app.processEvents()
            check("history row shows its files and reports", hp.files.count() >= 1, hp.title.text())
            n = len(opened)
            [b for b in hp.findChildren(QPushButton) if b.text() == "فتح مجلد التشغيل"][0].click()
            check("history open run folder (recorded)", len(opened) == n + 1 and Path(opened[-1]).is_dir(), opened[-1:])

        # ---- Settings page controls (saved only to the sandbox settings file)
        nav("settings")
        sp = w.pages["settings"]
        sp.reset()
        defaults = AppSettings(project_root=settings.project_root)
        check("settings: restore defaults fills the fields",
              sp.fields["output_dir"].value() == defaults.output_dir and sp.fields["python"].value() == defaults.python)
        sp.save()
        check("settings: save writes the (sandbox) settings file", (sandbox / "settings.json").is_file())

        # ---- Logs page
        nav("monitor")
        mp = w.pages["monitor"]
        check("logs page has its controls", {"إلغاء / إيقاف المهمة المحددة", "مسح المهام المنتهية من القائمة"}
              <= {b.text() for b in mp.findChildren(QPushButton)})

        # ---- navigation while a background task runs (mock job, not ETL)
        job = q.add(Job("3", ["-c", "import time; time.sleep(5)"], True, program=settings.python))
        started = wait(lambda: job.status is Status.RUNNING, 30)
        busy = []
        while job.status is Status.RUNNING and len(busy) < 400:
            for key in keys:
                busy.append(nav(key))
        done = wait(lambda: job.status is Status.COMPLETED, 30)
        report["timings_ms"]["navigation_during_task_max"] = round(max(busy), 1) if busy else None
        check("navigation while a task runs stays fast (< 500 ms)", started and done and busy and max(busy) < 500,
              f"{len(busy)} clicks, max {max(busy) if busy else 0:.1f} ms, job {job.status.value}")
        nav("monitor")
        check("logs page lists the background task", mp.table.rowCount() >= 1)
        [b for b in mp.findChildren(QPushButton) if b.text() == "مسح المهام المنتهية من القائمة"][0].click()
        app.processEvents()
        check("logs: clear finished tasks", mp.table.rowCount() == 0)

        # ---- nothing duplicated after all this navigation
        check("no duplicate signal connections", {s: q.receivers(SIGNAL(s)) for s in sigs} == receivers)
        check("no duplicate widgets on any page",
              {k: len(p.findChildren(QObject)) for k, p in w.pages.items()} == widgets,
              {k: len(p.findChildren(QObject)) - widgets[k] for k, p in w.pages.items()})
        check("pages are reused (one instance each)", list(w.pages) == keys and
              len({id(p) for p in w.pages.values()}) == len(keys))

        # ---- the run screen with a real engine run, entirely in the sandbox
        sample = args.get("--sample-csv")
        if sample and Path(sample).is_file():
            box = AppSettings(project_root=settings.project_root, python=settings.python,
                              output_dir=str(sandbox / "output"), state_dir=str(sandbox / "state"),
                              input_dir=str(sandbox / "input"), collection_dir=str(sandbox))
            for d in ("output", "state", "input"):
                (sandbox / d).mkdir(exist_ok=True)
            w2 = MainWindow(box)
            w2.show()
            page = w2.pages["run"]
            page.task_list.setCurrentRow(2)                        # Customer
            check("run screen: selected task matches details", page.title.text() == page.task.title)
            page.add_files([Path(sample)])
            job2 = page.run()
            ok = job2 is not None and wait(lambda: job2.status not in (Status.PENDING, Status.RUNNING), 180)
            check("run screen: engine run on the sample file completes", ok and job2.status is Status.COMPLETED,
                  job2.note if job2 else "not started")
            check("run screen: report written in the (sandbox) output", bool(job2 and job2.reports and
                  all(Path(r).is_file() and Path(sandbox / "output") in Path(r).parents for r in job2.reports)))
            QFileDialog.getExistingDirectory = staticmethod(lambda *a, **k: str(sandbox / "downloads"))
            QMessageBox.warning = staticmethod(lambda *a, **k: None)
            page.all_btn.click()
            check("run screen: «تحميل كل التقارير» copies every report",
                  sorted(p.name for p in (sandbox / "downloads").glob("*.xlsx")) ==
                  sorted(Path(r).name for r in job2.reports), list((sandbox / "downloads").glob("*")))
            n = len(opened)
            [b for b in page.results.findChildren(QPushButton) if b.text() == "فتح مجلد المخرجات"][0].click()
            check("run screen: open output folder (recorded)", len(opened) == n + 1, opened[-1:])
            w2.close()
        w.close()
    except Exception:  # noqa: BLE001 - reported, never hidden
        report["errors"].append(traceback.format_exc())
    report["passed"] = sum(c["ok"] for c in report["checks"])
    report["failed"] = sum(not c["ok"] for c in report["checks"])
    report["ok"] = report["failed"] == 0 and not report["errors"]
    out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0 if report["ok"] else 1
