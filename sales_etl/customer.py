"""Task 3: sales, returns and net per customer over every transaction in the input files.

In the ERP layout the columns used are:

* B ``نوع المستند``  -> مبيعات / مرتجعات
* F ``العميل``       -> customer code (one report row per code)
* G (blank header)  -> customer name: the column right after the code
* T ``صافى``         -> amount, summed with its original sign

No date filtering. The implementation (and the file lifecycle) is shared with task 4: see
``grouped.py``.
"""

from __future__ import annotations

from datetime import datetime
from typing import Any

from sales_etl.config import COL_CUSTOMER, TOTAL_COLUMNS, PipelineConfig
from sales_etl.grouped import GroupResult, GroupSpec, group_table, run_group_totals

COL_CUSTOMER_CODE = "كود العميل"
COL_CUSTOMER_NAME = "اسم العميل"
CUSTOMER_COLUMNS: tuple[str, ...] = (COL_CUSTOMER_CODE, COL_CUSTOMER_NAME, *TOTAL_COLUMNS.values())
CUSTOMER = GroupSpec(task="Customer", code_column=COL_CUSTOMER, columns=CUSTOMER_COLUMNS,
                     report_name="customer_totals.xlsx")
CustomerResult = GroupResult


def customer_table(result: GroupResult) -> list[list[Any]]:
    return group_table(result)


def run_customer_totals(config: PipelineConfig, *, now: datetime | None = None) -> GroupResult:
    """Read-only: per-customer totals over every CSV in ``config.input_dir``; nothing is moved."""
    return run_group_totals(CUSTOMER, config, now=now)
