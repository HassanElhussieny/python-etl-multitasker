"""Task 4: sales, returns and net per sales representative over every transaction in the input files.

In the ERP layout the columns used are:

* B ``نوع المستند``  -> مبيعات / مرتجعات
* I ``المندوب``      -> representative code (one report row per code)
* J (blank header)  -> representative name: the column right after the code
* T ``صافى``         -> amount, summed with its original sign

The customer columns (F, G) are not used. No date filtering. The implementation (and the
file lifecycle) is shared with task 3: see ``grouped.py``.
"""

from __future__ import annotations

from datetime import datetime
from typing import Any

from sales_etl.config import COL_REP, TOTAL_COLUMNS, PipelineConfig
from sales_etl.grouped import GroupResult, GroupSpec, group_table, run_group_totals

COL_REP_CODE = "كود المندوب"
COL_REP_NAME = "اسم المندوب"
REPRESENTATIVE_COLUMNS: tuple[str, ...] = (COL_REP_CODE, COL_REP_NAME, *TOTAL_COLUMNS.values())
REPRESENTATIVE = GroupSpec(task="Representative", code_column=COL_REP, columns=REPRESENTATIVE_COLUMNS,
                           report_name="representative_totals.xlsx")


def representative_table(result: GroupResult) -> list[list[Any]]:
    return group_table(result)


def run_representative_totals(config: PipelineConfig, *, now: datetime | None = None) -> GroupResult:
    """Read-only: per-representative totals over every CSV in ``config.input_dir``; nothing is moved."""
    return run_group_totals(REPRESENTATIVE, config, now=now)
