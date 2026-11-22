"""The job queue: each job is one run of the existing command line in its own process (QProcess).

Scheduling rules (the engine's lock and lifecycle are never bypassed or changed):

* a job that moves files (managed) runs alone: only when nothing else runs and ``state/etl.lock`` is free;
* read-only jobs (``--no-move``, Task 5 option 3) may run together (up to ``max_parallel_readonly``), never beside a
  managed job (it moves the input files they read) and never while the lock is held by another process;
* jobs start in the order they were added; a waiting managed job holds back the jobs behind it;
* a busy lock (e.g. a run started from the terminal) keeps the job Pending with a clear note - it is retried.

Statuses: Pending, Running, Completed, Failed, Cancelled. There is no progress percentage: the engine does not
report one, so only the elapsed time and the live output are shown.
"""

from __future__ import annotations

import codecs
import itertools
import json
import os
import re
from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Callable

from PySide6.QtCore import QObject, QProcess, QProcessEnvironment, QTimer, Signal

from gui.services.catalog import BY_KEY, matches_output, moves_files
from gui.services.lock import LockState, probe_lock
from gui.services.settings import AppSettings, app_data_dir
from gui.services.staging import changed_sources, remove_staging


class Status(Enum):
    PENDING = "Pending"
    RUNNING = "Running"
    COMPLETED = "Completed"
    FAILED = "Failed"
    CANCELLED = "Cancelled"


STATUS_AR = {Status.PENDING: "بالانتظار", Status.RUNNING: "قيد التشغيل", Status.COMPLETED: "مكتملة",
             Status.FAILED: "فشلت", Status.CANCELLED: "أُلغيت"}
#: the command line's exit codes (sales_etl.cli)
EXIT_TEXT = {0: "اكتملت بنجاح", 1: "اكتملت مع مشاكل أو فشل التدقيق — راجع أسطر FAIL",
             2: "خطأ في المعطيات، أو القفل مشغول بعملية أخرى", 3: "لا توجد ملفات إدخال — لم يُنشأ تقرير"}
_ids = itertools.count(1)
_XLSX = re.compile(r"([A-Za-z]:\\[^\r\n]*?\.xlsx)")


def history_file() -> Path:
    override = os.environ.get("ETL_GUI_HISTORY")
    return Path(override) if override else app_data_dir() / "jobs_history.jsonl"


@dataclass
class Job:
    task_key: str
    arguments: list[str]
    read_only: bool
    files_at_enqueue: list[str] = field(default_factory=list)
    id: str = field(default_factory=lambda: f"J{datetime.now():%y%m%d}-{next(_ids):04d}")
    status: Status = Status.PENDING
    created: datetime = field(default_factory=datetime.now)
    started: datetime | None = None
    finished: datetime | None = None
    exit_code: int | None = None
    note: str = ""
    log: list[str] = field(default_factory=list)
    output_dirs: list[str] = field(default_factory=list)
    reports: list[str] = field(default_factory=list)
    program: str | None = None        # tests may run another program; normally the settings' Python
    staging_dir: str | None = None    # desktop run: the folder of copied input files (removed when it ends)
    sources: list[tuple[str, str]] = field(default_factory=list)   # the user's original files and their SHA-256

    @property
    def title(self) -> str:
        return BY_KEY[self.task_key].title if self.task_key in BY_KEY else self.task_key

    @property
    def moves_files(self) -> bool:
        return self.task_key in BY_KEY and moves_files(self.task_key, self.read_only)

    @property
    def mode_text(self) -> str:
        return "ينقل الملفات بعد التحقق" if self.moves_files else "قراءة فقط"

    def elapsed(self) -> float:
        if not self.started:
            return 0.0
        return ((self.finished or datetime.now()) - self.started).total_seconds()

    @property
    def text(self) -> str:
        return "".join(self.log)

    def record(self) -> dict:
        return {"id": self.id, "task": self.task_key, "title": self.title, "read_only": self.read_only,
                "status": self.status.value, "exit_code": self.exit_code, "note": self.note,
                "created": self.created.isoformat(timespec="seconds"),
                "started": self.started.isoformat(timespec="seconds") if self.started else None,
                "finished": self.finished.isoformat(timespec="seconds") if self.finished else None,
                "arguments": self.arguments, "output_dirs": self.output_dirs, "reports": self.reports,
                "files_at_enqueue": self.files_at_enqueue, "sources": [s for s, _ in self.sources]}


class JobQueue(QObject):
    job_added = Signal(str)
    job_changed = Signal(str)
    output = Signal(str, str)

    def __init__(self, settings: AppSettings, probe: Callable[[Path], LockState] = probe_lock,
                 parent: QObject | None = None) -> None:
        super().__init__(parent)
        self.settings = settings
        self.probe = probe
        self.jobs: list[Job] = []
        self._procs: dict[str, QProcess] = {}
        self._finished_procs: list[QProcess] = []
        self._decoders: dict[str, dict[str, codecs.IncrementalDecoder]] = {}
        self._before: dict[str, set[str]] = {}
        self._timer = QTimer(self)
        self._timer.setInterval(1500)
        self._timer.timeout.connect(self.schedule)
        self._timer.start()

    # ------------------------------------------------------------------ public
    def add(self, job: Job) -> Job:
        self.jobs.append(job)
        self.job_added.emit(job.id)
        self.schedule()
        return job

    def get(self, job_id: str) -> Job | None:
        return next((j for j in self.jobs if j.id == job_id), None)

    def running(self) -> list[Job]:
        return [j for j in self.jobs if j.status is Status.RUNNING]

    def counts(self) -> dict[Status, int]:
        return {s: sum(1 for j in self.jobs if j.status is s) for s in Status}

    def cancel(self, job_id: str) -> None:
        job = self.get(job_id)
        if job is None:
            return
        if job.status is Status.PENDING:
            job.status, job.finished = Status.CANCELLED, datetime.now()
            job.note = "أُلغيت قبل البدء"
            self._release_staging(job)
            self._persist(job)
            self.job_changed.emit(job.id)
        elif job.status is Status.RUNNING and job.id in self._procs:
            job.note = ("أُوقفت أثناء التشغيل. آلية السجل في المحرك تضمن السلامة: إن لم يُسجَّل COMMIT تبقى الملفات "
                        "في input، وإن سُجِّل يُكمل التشغيل التالي النقل.") if job.moves_files else "أُوقفت أثناء التشغيل"
            job.status = Status.CANCELLED            # set before the kill: finished() keeps it
            self._procs[job.id].kill()

    def clear_finished(self) -> None:
        self.jobs = [j for j in self.jobs if j.status in (Status.PENDING, Status.RUNNING)]
        self.job_changed.emit("")

    def lock_state(self) -> LockState:
        return self.probe(self.settings.lock_file)

    # ------------------------------------------------------------------ scheduling
    def schedule(self) -> None:
        running = self.running()
        managed_running = any(j.moves_files for j in running)
        readonly_running = sum(1 for j in running if not j.moves_files)
        lock = None
        for job in [j for j in self.jobs if j.status is Status.PENDING]:
            if managed_running:
                return
            if lock is None:
                lock = self.lock_state()
            if lock is LockState.BUSY:
                self._wait(job, "القفل العام (state/etl.lock) مشغول بعملية أخرى — قد تكون تشغيلًا من سطر الأوامر. "
                                "ستبدأ المهمة تلقائيًا عند تحريره.")
                return
            if job.moves_files:
                if running:
                    self._wait(job, "بانتظار انتهاء المهام الجارية (مهمة تنقل الملفات تعمل منفردة).")
                    return
                self._start(job)
                return                                   # alone: nothing else starts beside it
            if readonly_running >= self.settings.max_parallel_readonly:
                self._wait(job, "بانتظار مكان: الحد الأقصى للمهام المتوازية للقراءة فقط.")
                return
            self._start(job)
            readonly_running += 1
            running = self.running()

    def _wait(self, job: Job, note: str) -> None:
        if job.note != note:
            job.note = note
            self.job_changed.emit(job.id)

    def _start(self, job: Job) -> None:
        out = self.settings.path("output_dir")
        self._before[job.id] = {p.name for p in out.iterdir()} if out.is_dir() else set()
        proc = QProcess(self)
        env = QProcessEnvironment.systemEnvironment()
        env.insert("PYTHONIOENCODING", "utf-8")
        env.insert("PYTHONUTF8", "1")
        proc.setProcessEnvironment(env)
        proc.setWorkingDirectory(self.settings.project_root)
        proc.setProgram(job.program or self.settings.python)
        proc.setArguments(job.arguments)
        self._decoders[job.id] = {"out": codecs.getincrementaldecoder("utf-8")("replace"),
                                  "err": codecs.getincrementaldecoder("utf-8")("replace")}
        proc.readyReadStandardOutput.connect(lambda j=job, p=proc: self._read(j, p, "out"))
        proc.readyReadStandardError.connect(lambda j=job, p=proc: self._read(j, p, "err"))
        proc.finished.connect(lambda code, status, j=job: self._finished(j, code, status))
        proc.errorOccurred.connect(lambda error, j=job: self._error(j, error))
        self._procs[job.id] = proc
        job.status, job.started, job.note = Status.RUNNING, datetime.now(), ""
        job.log.append(f"$ {job.program or self.settings.python} {' '.join(job.arguments)}\n\n")
        self.job_changed.emit(job.id)
        proc.start()

    def _read(self, job: Job, proc: QProcess, channel: str) -> None:
        data = bytes(proc.readAllStandardOutput() if channel == "out" else proc.readAllStandardError())
        text = self._decoders[job.id][channel].decode(data)
        if text:
            job.log.append(text)
            self.output.emit(job.id, text)

    def _error(self, job: Job, error: QProcess.ProcessError) -> None:
        if error == QProcess.ProcessError.FailedToStart:
            job.status, job.finished = Status.FAILED, datetime.now()
            job.note = f"تعذّر تشغيل المفسّر: {job.program or self.settings.python} — راجع الإعدادات."
            self._release_staging(job)
            self._cleanup(job)
            self._persist(job)
            self.job_changed.emit(job.id)
            self.schedule()

    def _finished(self, job: Job, code: int, status: QProcess.ExitStatus) -> None:
        proc = self._procs.get(job.id)
        if proc is not None:
            self._read(job, proc, "out")
            self._read(job, proc, "err")
        job.finished = datetime.now()
        job.exit_code = code if status == QProcess.ExitStatus.NormalExit else None
        if job.status is not Status.CANCELLED:
            if status != QProcess.ExitStatus.NormalExit:
                job.status, job.note = Status.FAILED, "توقفت العملية بشكل غير طبيعي"
            else:
                job.status = Status.COMPLETED if code == 0 else Status.FAILED
                job.note = EXIT_TEXT.get(code, f"رمز الخروج {code}")
        self._collect_outputs(job)
        self._release_staging(job)
        self._cleanup(job)
        self._persist(job)
        self.job_changed.emit(job.id)
        self.schedule()

    def _collect_outputs(self, job: Job) -> None:
        out = self.settings.path("output_dir")
        task = BY_KEY.get(job.task_key)
        before = self._before.pop(job.id, set())
        if out.is_dir() and task is not None:
            new = sorted(p for p in out.iterdir() if p.is_dir() and p.name not in before and matches_output(task, p.name))
            job.output_dirs = [str(p) for p in new]
            job.reports = [str(f) for d in new for f in sorted(d.glob("*.xlsx")) if not f.name.startswith("~$")]
        for m in _XLSX.findall(job.text):          # report paths the engine printed (inside the output folder only)
            p = Path(m.strip())
            try:
                inside = p.resolve().is_relative_to(out.resolve())
            except OSError:
                inside = False
            known = {Path(r).resolve() for r in job.reports}          # 8.3 short and long names are one file
            if inside and p.is_file() and p.resolve() not in known:
                job.reports.append(str(p))

    def _release_staging(self, job: Job) -> None:
        """After the process has ended: check the user's originals are unchanged, then delete the staged copies
        (the engine ran with --no-move, so nothing else refers to them)."""
        changed = changed_sources(job.sources)
        if changed:
            job.note += f" | تحذير: تغيّر ملف أصلي أثناء التشغيل: {', '.join(changed)}"
        if job.staging_dir:
            remove_staging(job.staging_dir)

    def _cleanup(self, job: Job) -> None:
        # The finished QProcess stays a child of the queue (freed with it). deleteLater() here, combined with
        # the queue's own deletion, aborts the interpreter on PySide6 6.12 / Python 3.14 (reproduced in testing).
        proc = self._procs.pop(job.id, None)
        self._decoders.pop(job.id, None)
        if proc is not None:
            self._finished_procs.append(proc)

    def _persist(self, job: Job) -> None:
        try:
            path = history_file()
            path.parent.mkdir(parents=True, exist_ok=True)
            with path.open("a", encoding="utf-8") as fh:
                fh.write(json.dumps(job.record(), ensure_ascii=False) + "\n")
        except OSError:
            pass

    def shutdown(self, kill: bool = False) -> None:
        self._timer.stop()
        for job_id, proc in list(self._procs.items()):
            if kill:
                self.cancel(job_id)
            proc.waitForFinished(30000)
