"""Durable bookkeeping for input files: lock, write-ahead journal, safe moves, recovery.

Protocol for one managed run (see ``managed.py``):

1. take the run lock (released by the OS even if the process dies);
2. ``recover()`` finishes moves that an interrupted run had already committed;
3. outputs are written and verified;
4. ONE ``COMMIT`` journal line (fsync'd) records every file's disposition and
   destination - this is the point after which a file counts as processed;
5. files are moved (never overwriting) and each move is journalled as ``MOVED``.

A crash before step 4 leaves every file in ``input`` with no commit, so it is
simply processed again (the abandoned run folder is labelled as such). After a
crash in step 5, ``recover()`` confirms moves that did complete; files still in
``input`` are never archived behind the user's back - they are processed again.
The journal is a trace only: it never causes a file to be skipped.
"""

from __future__ import annotations

import errno
import json
import os
import shutil
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

from sales_etl.ingest import sha256_file

JOURNAL_NAME = "processing_journal.jsonl"
LOCK_NAME = "etl.lock"

# dispositions
PROCESSED = "PROCESSED"
FAILED = "FAILED"
RETAINED = "RETAINED"  # left in input (system-level problem); never committed


class RunLockedError(RuntimeError):
    """Another run currently holds the lock."""


class JournalCorruptError(RuntimeError):
    """A journal line other than a torn final line cannot be parsed."""


class MoveError(RuntimeError):
    """A file could not be moved/verified; the source is left in place."""


# ------------------------------------------------------------------------- lock
class RunLock:
    """Exclusive, non-blocking OS lock on ``state/etl.lock``."""

    def __init__(self, path: Path) -> None:
        self.path = path
        self._handle: Any = None

    def __enter__(self) -> RunLock:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        handle = open(self.path, "a+b")  # noqa: SIM115 - held for the lock's lifetime
        try:
            if sys.platform == "win32":
                import msvcrt

                handle.seek(0)
                msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl

                fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError as exc:
            handle.close()
            raise RunLockedError(f"another ETL run is in progress (lock: {self.path})") from exc
        self._handle = handle
        return self

    def __exit__(self, *exc_info: object) -> None:
        handle, self._handle = self._handle, None
        if handle is None:
            return
        try:
            if sys.platform == "win32":
                import msvcrt

                handle.seek(0)
                msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                import fcntl

                fcntl.flock(handle.fileno(), fcntl.LOCK_UN)
        finally:
            handle.close()


# ---------------------------------------------------------------------- journal
@dataclass(frozen=True)
class CommittedFile:
    run_id: str
    run_dir: str
    sha256: str
    file_name: str
    source: str
    destination: str
    disposition: str
    reason: str

    @property
    def key(self) -> tuple[str, str, str]:
        return (self.run_id, self.sha256, self.source)


@dataclass
class JournalState:
    commits: list[CommittedFile] = field(default_factory=list)
    closed: set[tuple[str, str, str]] = field(default_factory=set)  # moved or abandoned
    started: dict[str, dict[str, Any]] = field(default_factory=dict)
    finished: set[str] = field(default_factory=set)
    abandoned_runs: set[str] = field(default_factory=set)
    committed_runs: set[str] = field(default_factory=set)
    torn_tail: bool = False

    def pending_moves(self) -> list[CommittedFile]:
        return [c for c in self.commits if c.key not in self.closed]

    def interrupted_runs(self) -> list[str]:
        return [run for run in self.started
                if run not in self.finished and run not in self.abandoned_runs]


class Journal:
    """Append-only JSON-lines journal; every append is flushed and fsync'd."""

    def __init__(self, path: Path) -> None:
        self.path = path
        path.parent.mkdir(parents=True, exist_ok=True)

    def append(self, event: str, **data: Any) -> dict[str, Any]:
        record = {"event": event, "ts": datetime.now().isoformat(timespec="seconds"), **data}
        line = json.dumps(record, ensure_ascii=False, default=str)
        needs_newline = False
        if self.path.exists() and self.path.stat().st_size:
            with open(self.path, "rb") as handle:
                handle.seek(-1, os.SEEK_END)
                needs_newline = handle.read(1) != b"\n"  # previous write was torn
        with open(self.path, "a", encoding="utf-8", newline="\n") as handle:
            handle.write(("\n" if needs_newline else "") + line + "\n")
            handle.flush()
            os.fsync(handle.fileno())
        return record

    def records(self) -> tuple[list[dict[str, Any]], bool]:
        """All records, plus whether a torn (unparseable) final line was ignored."""
        if not self.path.exists():
            return [], False
        lines = self.path.read_text(encoding="utf-8").splitlines()
        records: list[dict[str, Any]] = []
        torn = False
        for index, line in enumerate(lines):
            if not line.strip():
                continue
            try:
                records.append(json.loads(line))
            except json.JSONDecodeError as exc:
                is_last = all(not rest.strip() for rest in lines[index + 1:])
                if is_last:
                    torn = True  # interrupted mid-write: the record never took effect
                    continue
                if any(rest.strip() for rest in lines[index + 1:]) and _is_torn_then_resumed(line):
                    torn = True
                    continue
                raise JournalCorruptError(f"{self.path} line {index + 1}: {exc}") from exc
        return records, torn

    def state(self) -> JournalState:
        records, torn = self.records()
        state = JournalState(torn_tail=torn)
        for rec in records:
            event = rec.get("event")
            if event == "RUN_STARTED":
                state.started[rec["run_id"]] = rec
            elif event == "RUN_FINISHED":
                state.finished.add(rec["run_id"])
            elif event == "RUN_ABANDONED":
                state.abandoned_runs.add(rec["run_id"])
            elif event == "COMMIT":
                state.committed_runs.add(rec["run_id"])
                for item in rec["files"]:
                    state.commits.append(CommittedFile(
                        run_id=rec["run_id"], run_dir=rec["run_dir"], sha256=item["sha256"],
                        file_name=item["file_name"], source=item["source"],
                        destination=item["destination"], disposition=item["disposition"],
                        reason=item.get("reason", ""),
                    ))
            elif event in ("MOVED", "MOVE_ABANDONED"):
                state.closed.add((rec["run_id"], rec["sha256"], rec["source"]))
        return state


def _is_torn_then_resumed(line: str) -> bool:
    """A torn write followed by later appends (append() always starts a new line)."""
    return not line.rstrip().endswith("}")


# ------------------------------------------------------------------------ moves
def unique_destination(folder: Path, file_name: str, sha256: str, now: datetime,
                       reserved: set[Path] | None = None) -> Path:
    """``folder/file_name`` or, if taken, ``stem__<timestamp>_<sha8>[_n].ext`` - never an existing path."""
    reserved = reserved or set()
    candidate = folder / file_name
    stem, suffix = Path(file_name).stem, Path(file_name).suffix
    attempt = 0
    while candidate.exists() or candidate in reserved:
        attempt += 1
        tail = f"_{attempt}" if attempt > 1 else ""
        candidate = folder / f"{stem}__{now:%Y%m%d-%H%M%S}_{sha256[:8]}{tail}{suffix}"
    return candidate


def move_file(source: Path, destination: Path, sha256: str) -> Path:
    """Move without ever overwriting, then verify the destination's SHA-256."""
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        raise MoveError(f"destination already exists: {destination}")
    if sha256_file(source) != sha256:
        raise MoveError(f"source content changed since it was processed: {source}")
    try:
        os.rename(source, destination)  # atomic on one volume; never overwrites on Windows
    except OSError as exc:
        if exc.errno != errno.EXDEV:
            raise MoveError(f"cannot move {source} -> {destination}: {exc}") from exc
        with open(source, "rb") as src, open(destination, "xb") as dst:  # cross-volume
            shutil.copyfileobj(src, dst)
            dst.flush()
            os.fsync(dst.fileno())
        if sha256_file(destination) != sha256:
            raise MoveError(f"copy verification failed: {destination}")  # source kept
        os.unlink(source)
    if sha256_file(destination) != sha256:
        raise MoveError(f"moved file does not match the recorded SHA-256: {destination}")
    return destination


# --------------------------------------------------------------------- recovery
@dataclass(frozen=True)
class RecoveryAction:
    run_id: str
    file_name: str
    action: str
    detail: str


def recover(journal: Journal, now: datetime) -> list[RecoveryAction]:
    """Finish what interrupted runs left behind. Never deletes un-verified data."""
    state = journal.state()
    actions: list[RecoveryAction] = []
    for commit in state.pending_moves():
        source, destination = Path(commit.source), Path(commit.destination)
        dst_ok = destination.exists() and sha256_file(destination) == commit.sha256
        base = {"run_id": commit.run_id, "sha256": commit.sha256, "source": commit.source}
        if source.exists():
            # Still in input: it is NOT archived behind the user's back - the current run
            # processes it again (managed.py closes this pending move as superseded).
            actions.append(RecoveryAction(commit.run_id, commit.file_name, "LEFT_IN_INPUT",
                                          "earlier move never completed; file will be processed in this run"))
        elif dst_ok:
            journal.append("MOVED", **base, destination=str(destination), recovered=True)
            actions.append(RecoveryAction(commit.run_id, commit.file_name, "MOVE_CONFIRMED",
                                          f"already at {destination}; journal updated"))
        else:
            reason = "source missing and no verified copy at the destination"
            journal.append("MOVE_ABANDONED", **base, destination=str(destination), error=reason)
            actions.append(RecoveryAction(commit.run_id, commit.file_name, "NEEDS_ATTENTION",
                                          f"{reason}; nothing was moved or deleted"))

    for run_id in state.interrupted_runs():
        info = state.started[run_id]
        committed = run_id in state.committed_runs
        journal.append("RUN_ABANDONED", run_id=run_id, committed=committed)
        run_dir = Path(info.get("run_dir", ""))
        note = run_dir / "RUN_INTERRUPTED.txt"
        if run_dir.is_dir() and not note.exists():
            note.write_text(
                (f"This run ({run_id}) was interrupted after its results were committed.\n"
                 "Its report is valid. Files it could not archive stayed in input and were "
                 "processed again by a later run.\n")
                if committed else
                (f"This run ({run_id}) was interrupted BEFORE its results were committed.\n"
                 "Its outputs are VOID. No file was moved; every input file was processed again "
                 "by a later run.\n"),
                encoding="utf-8",
            )
        actions.append(RecoveryAction(
            run_id, "", "RUN_INTERRUPTED_AFTER_COMMIT" if committed else "RUN_VOIDED",
            "report valid; unarchived files are processed again" if committed
            else "outputs void; its files stayed in input and are reprocessed"))
    return actions
