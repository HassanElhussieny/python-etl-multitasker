"""History and reports, rebuilt from what the engine already writes (read only):

* every run folder in ``output`` (its name says the task and the execution date),
* the journals in ``state`` (RUN_STARTED / COMMIT / MOVED / RUN_FINISHED with the run's status and files),
* Task 1's ``run_manifest.json`` and ``audit_*.json`` (audit status), ``RUN_INTERRUPTED.txt`` of interrupted runs,
* the GUI's own job log (``jobs_history.jsonl``), when the run was started from the desktop app.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from gui.services.catalog import TASKS, task_for_folder
from gui.services.jobs import history_file


@dataclass
class Run:
    folder: Path
    task_title: str
    task_key: str
    when: datetime
    status: str                      # SUCCEEDED / COMPLETED_WITH_PROBLEMS / FAILED / INTERRUPTED / read-only ...
    audit: str = ""
    files: list[dict] = field(default_factory=list)      # journal: name, disposition, destination
    reports: list[Path] = field(default_factory=list)
    source: str = ""                 # "journal", "manifest", "gui", "folder"


@dataclass
class Report:
    path: Path
    run_folder: str
    task_title: str
    modified: datetime
    size: int


STATUS_AR = {"SUCCEEDED": "ناجح", "COMPLETED_WITH_PROBLEMS": "اكتمل مع مشاكل", "FAILED": "فشل",
             "INTERRUPTED": "مُقاطَع", "READ_ONLY": "قراءة فقط (بدون سجل)", "UNKNOWN": "غير معروف",
             "RUNNING_OR_ABANDONED": "لم يُسجَّل انتهاؤه"}


def _journal_runs(state_dir: Path) -> dict[str, dict]:
    """run_id -> {status, files, task} from every journal in state (malformed lines are skipped)."""
    runs: dict[str, dict] = {}
    if not state_dir.is_dir():
        return runs
    for path in state_dir.glob("*.jsonl"):
        try:
            lines = path.read_text(encoding="utf-8").splitlines()
        except OSError:
            continue
        for line in lines:
            try:
                e = json.loads(line)
            except ValueError:
                continue
            run_id = e.get("run_id")
            if not run_id:
                continue
            r = runs.setdefault(run_id, {"status": None, "files": [], "journal": path.name})
            if e.get("event") == "RUN_FINISHED":
                r["status"] = e.get("status")
            elif e.get("event") == "COMMIT":
                r["files"] = [{"name": f.get("file_name"), "disposition": f.get("disposition"),
                               "destination": f.get("destination")} for f in e.get("files", [])]
            elif e.get("event") == "RUN_STARTED" and not r["files"]:
                r["files"] = [{"name": f.get("file_name"), "disposition": "", "destination": ""}
                              for f in e.get("files", [])]
    return runs


def _gui_jobs() -> dict[str, dict]:
    """output folder -> the GUI job record that produced it."""
    out: dict[str, dict] = {}
    try:
        lines = history_file().read_text(encoding="utf-8").splitlines()
    except OSError:
        return out
    for line in lines:
        try:
            rec = json.loads(line)
        except ValueError:
            continue
        for d in rec.get("output_dirs", []):
            out[Path(d).name] = rec
    return out


def load_runs(output_dir: Path, state_dir: Path) -> list[Run]:
    if not output_dir.is_dir():
        return []
    journal = _journal_runs(state_dir)
    gui = _gui_jobs()
    runs = []
    for folder in output_dir.iterdir():
        if not folder.is_dir():
            continue
        task = task_for_folder(folder.name)
        title = task.title if task else "مخرجات قديمة / غير معروفة"
        if task and task.key in ("5-1", "5-2"):
            title = "تحليل القطاعات والمناديب (المهمة 5)"
            if any(folder.glob("Customers_Not_Worked_*.xlsx")):
                title += " — أمس والأيام السابقة"
        try:
            when = datetime.fromtimestamp(folder.stat().st_mtime)
        except OSError:
            continue
        run = Run(folder, title, task.key if task else "", when, "UNKNOWN", source="folder")
        j = journal.get(folder.name)
        if j:
            run.status, run.files, run.source = j["status"] or "RUNNING_OR_ABANDONED", j["files"], "journal"
        manifest = folder / "run_manifest.json"
        if manifest.is_file():
            try:
                m = json.loads(manifest.read_text(encoding="utf-8"))
                run.status = m.get("status") or run.status
                run.source = "manifest"
            except (OSError, ValueError):
                pass
        if (folder / "RUN_INTERRUPTED.txt").is_file():
            run.status = "INTERRUPTED"
        for audit in folder.glob("audit_*.json"):
            try:
                run.audit = str(json.loads(audit.read_text(encoding="utf-8")).get("audit_status", ""))
            except (OSError, ValueError, AttributeError):
                pass
        if run.status == "UNKNOWN":
            g = gui.get(folder.name)
            if g:                                    # a desktop run: its engine exit code says how it ended
                run.source = "gui"
                run.status = "SUCCEEDED" if g.get("exit_code") == 0 else "FAILED"
                run.files = [{"name": Path(x).name, "disposition": "SOURCE", "destination": x}
                             for x in g.get("sources", [])]
            elif task is not None:
                run.status = "READ_ONLY"            # no journal: --no-move or Task 5 option 3
        run.reports = sorted(f for f in folder.glob("*.xlsx") if not f.name.startswith("~$"))
        runs.append(run)
    return sorted(runs, key=lambda r: r.when, reverse=True)


def load_reports(output_dir: Path) -> list[Report]:
    reports = []
    if not output_dir.is_dir():
        return reports
    for folder in output_dir.iterdir():
        if not folder.is_dir():
            continue
        task = task_for_folder(folder.name)
        for f in folder.glob("*.xlsx"):
            if f.name.startswith("~$"):
                continue
            try:
                st = f.stat()
            except OSError:
                continue
            reports.append(Report(f, folder.name, task.title if task else "—", datetime.fromtimestamp(st.st_mtime),
                                  st.st_size))
    return sorted(reports, key=lambda r: r.modified, reverse=True)


def task_titles() -> list[tuple[str, str]]:
    return [(t.key, t.title) for t in TASKS]
