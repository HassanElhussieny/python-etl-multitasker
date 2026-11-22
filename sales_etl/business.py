"""Company-specific business configuration: customer sectors, representatives, Cover groups and targets, and the
responsibles of the customer report (Task 5).

These values are business data, not code, so they are not written in the source. They are read, in order, from:

1. the file named by the environment variable ``SALES_ETL_BUSINESS_CONFIG``;
2. ``config/business.local.json`` - the real values, kept only on the production machine (git-ignored);
3. ``config/business.example.json`` - synthetic example values shipped with the repository (also used by the tests).

The file is JSON with the keys of :class:`BusinessConfig`; see ``config/business.example.json``.
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path

from sales_etl.config import PROJECT_ROOT

ENV_VAR = "SALES_ETL_BUSINESS_CONFIG"
CONFIG_DIR = PROJECT_ROOT / "config"
LOCAL_FILE = CONFIG_DIR / "business.local.json"
EXAMPLE_FILE = CONFIG_DIR / "business.example.json"


class BusinessConfigError(RuntimeError):
    """The business configuration file is missing or malformed."""


@dataclass(frozen=True)
class BusinessConfig:
    sectors: tuple[str, ...]
    representatives: tuple[str, ...]
    excluded_representative: str
    cover_groups: tuple[tuple[str, ...], ...]
    direct_addition: str
    cover_targets: dict[str, int]
    customer_responsibles: tuple[str, ...]
    source: Path


def config_path() -> Path:
    override = os.environ.get(ENV_VAR)
    if override:
        return Path(override)
    return LOCAL_FILE if LOCAL_FILE.is_file() else EXAMPLE_FILE


def load_business_config(path: Path | None = None) -> BusinessConfig:
    path = Path(path) if path else config_path()
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise BusinessConfigError(f"cannot read the business configuration {path}: {exc}") from exc

    def names(key: str) -> tuple[str, ...]:
        value = raw.get(key)
        if not isinstance(value, list) or not all(isinstance(x, str) for x in value):
            raise BusinessConfigError(f"{path}: '{key}' must be a list of names")
        return tuple(value)

    def text(key: str) -> str:
        value = raw.get(key)
        if not isinstance(value, str):
            raise BusinessConfigError(f"{path}: '{key}' must be a name")
        return value

    groups = raw.get("cover_groups")
    if not isinstance(groups, list) or len(groups) != 2 or not all(isinstance(g, list) for g in groups):
        raise BusinessConfigError(f"{path}: 'cover_groups' must be two lists of names")
    targets = raw.get("cover_targets")
    if not isinstance(targets, dict) or not all(isinstance(v, int) for v in targets.values()):
        raise BusinessConfigError(f"{path}: 'cover_targets' must map names to whole numbers")
    return BusinessConfig(
        sectors=names("sectors"),
        representatives=names("representatives"),
        excluded_representative=text("excluded_representative"),
        cover_groups=tuple(tuple(str(x) for x in g) for g in groups),
        direct_addition=text("direct_addition"),
        cover_targets={str(k): int(v) for k, v in targets.items()},
        customer_responsibles=names("customer_responsibles"),
        source=path,
    )


BUSINESS = load_business_config()
