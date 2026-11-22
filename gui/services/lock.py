"""Is the engine's global run lock (``state/etl.lock``) held right now?

The probe uses the same non-blocking OS lock the engine uses and releases it at once, so it never holds the lock
and never waits. A missing lock file means free (the probe does not create it). The lock itself is never
removed, bypassed or changed: a busy lock only makes the GUI wait and say so.
"""

from __future__ import annotations

import sys
from enum import Enum
from pathlib import Path


class LockState(Enum):
    FREE = "free"
    BUSY = "busy"
    UNKNOWN = "unknown"


def probe_lock(lock_file: Path) -> LockState:
    if not lock_file.is_file():
        return LockState.FREE
    try:
        handle = open(lock_file, "r+b")  # noqa: SIM115 - closed below
    except OSError:
        return LockState.UNKNOWN
    try:
        if sys.platform == "win32":
            import msvcrt

            handle.seek(0)
            try:
                msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
            except OSError:
                return LockState.BUSY
            handle.seek(0)
            msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
        else:
            import fcntl

            try:
                fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            except OSError:
                return LockState.BUSY
            fcntl.flock(handle.fileno(), fcntl.LOCK_UN)
        return LockState.FREE
    finally:
        handle.close()
