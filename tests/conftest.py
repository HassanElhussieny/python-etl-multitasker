"""Builders for synthetic CSV files in the exact layout of the ERP export."""

from __future__ import annotations

import csv
import io
import os
from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest

# Tests always use the synthetic business configuration (fictional names/targets), never a local real one.
os.environ["SALES_ETL_BUSINESS_CONFIG"] = str(Path(__file__).resolve().parents[1] / "config" / "business.example.json")

from sales_etl.config import PipelineConfig  # noqa: E402
from sales_etl.pipeline import run_pipeline  # noqa: E402

SYNTHETIC_DATA = Path(__file__).parent / "data"      # synthetic fixtures only
S, R = "مبيعات", "مرتجعات"

# Header exactly as in the ERP export: 21 cells, blank headers at 4, 7, 10 and 21.
ERP_HEADER = [
    "التاريخ", "نوع المستند", "رقم المستند", "", "بيان", "العميل", "", " الحساب الفرعي",
    "المندوب", "", "اجمالى", "خصم 1", "خصم صيدلى", "خصم موزع", "خصم نقدى", "ض.قيمة مضا",
    "ضريبة 2", "ض.ا-ت صناع", "ضريبة 4", "صافى", "",
]


def money(value: Decimal | str | int) -> str:
    """ERP style: right-aligned, thousands separators, two decimals.

    Non-numeric strings are passed through unchanged (to build invalid rows).
    """
    try:
        return f"{Decimal(str(value).replace(',', '')):>31,.2f}"
    except Exception:  # noqa: BLE001
        return str(value)


@dataclass
class Txn:
    date: str
    doc_type: str
    net: str
    doc_no: str = "1"
    customer: str = "400100"
    series: str = "93"
    gross: str | None = None
    pharm: str = "0"
    dist: str = "0"
    cash: str = "0"

    def cells(self) -> list[str]:
        gross = self.gross if self.gross is not None else self.net
        return [
            self.date, self.doc_type, self.doc_no, self.series, "", self.customer,
            f"عميل {self.customer}", "", "2", "مندوب 2", money(gross), money(0), money(self.pharm),
            money(self.dist), money(self.cash), money(0), money(0), money(0), money(0),
            money(self.net), "",
        ]


def erp_csv_text(
    rows: list[Txn | list[str]],
    *,
    header: list[str] | None = None,
    delimiter: str = ";",
    footer: bool = True,
    footer_net: Decimal | None = None,
    footer_count: int | None = None,
    period: tuple[str, str] = ("26-09-2026", "19-10-2026"),
    footer_labels: tuple[str, str] = ("عدد المستندات", "الاجمالى العام"),
) -> str:
    out = io.StringIO()
    writer = csv.writer(out, delimiter=delimiter, lineterminator="\r\n")
    writer.writerow(["", "", "", "تقرير المبيعات والمرتجعات"])
    writer.writerow(["", "", "من", period[0], "الى", period[1]])
    writer.writerow(["", "", "اسم المندوب", ""])
    writer.writerow([])
    writer.writerow([])
    writer.writerow(["الفرع", "4", "فرع تجريبي", ""])
    writer.writerow([" المخزن", "1", "مخزن عام", ""])
    writer.writerow(header or ERP_HEADER)
    net_total = Decimal("0")
    for row in rows:
        if isinstance(row, Txn):
            writer.writerow(row.cells())
            try:
                net_total += Decimal(row.net.replace(",", ""))
            except Exception:  # noqa: BLE001
                pass
        else:
            writer.writerow(row)
    if footer:
        count = footer_count if footer_count is not None else len(rows)
        writer.writerow([footer_labels[0], str(count), ""])
        total = footer_net if footer_net is not None else net_total
        writer.writerow([footer_labels[1], *[""] * 18, money(total), ""])
    return out.getvalue()


def write_erp_csv(path: Path, rows: list[Txn | list[str]], *, encoding: str = "cp1256", **kwargs) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(erp_csv_text(rows, **kwargs).encode(encoding))
    return path


def branch_a_rows() -> list[Txn]:
    """The 235 SYNTHETIC transactions of tests/data/branch_a_compact.txt (see make_synthetic_branch.py)."""
    rows = []
    for line in (SYNTHETIC_DATA / "branch_a_compact.txt").read_text(encoding="utf-8").splitlines():
        if not line or line.startswith("#"):
            continue
        d, t, doc, series, cust, gross, pharm, dist, cash, net = line.split("|")
        rows.append(Txn(d, S if t == "S" else R, net, doc, cust, series, gross, pharm, dist, cash))
    return rows


@pytest.fixture
def run(tmp_path: Path):
    """Run the pipeline on ``tmp_path/input`` with a reference date (DD/MM/YYYY)."""
    from sales_etl.dates import parse_reference_date

    def _run(ref: str = "08/10/2026"):
        config = PipelineConfig(
            input_dir=tmp_path / "input", output_dir=tmp_path / "output",
            reference_date=parse_reference_date(ref),
        )
        return run_pipeline(config)

    (tmp_path / "input").mkdir(exist_ok=True)
    return _run


def outcome(result, name: str):
    return next(o for o in result.outcomes if o.file_name == name)


def check(result, name: str, scope: str | None = None):
    return [c for c in result.checks if c.check == name and (scope is None or c.scope == scope)]


REF = date(2026, 10, 8)


# ------------------------------------------------------------------ Task 5 option 1: the collection workbook
COLLECTION_HEADER = ["م", "كود العميل", "اسم العميل", "المندوب", "قيمة الفواتير المفتوحة", "المستحق 7-10",
                     "المحصل من مستحق 7-10", "الشيكات المرتدة", "المرتجعات", "مصروفات وتسويات", "المتبقي من 7-10",
                     "المستحق الجديد", "المحصل على الجديد"]


def write_collection(path: Path, rows, header=COLLECTION_HEADER) -> Path:
    """The real layout of التحصيل/<workbook>: title row, empty row, header on row 3 (B code, F المستحق 7-10,
    L المستحق الجديد), the customers (code, F, L), then a total and a note row without a code."""
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "القديم والجديد"
    ws.append([None, "متابعة التحصيل: القديم والجديد"])
    ws.append([])
    ws.append(list(header))
    for i, (code, old, new) in enumerate(rows, start=1):
        line = [None] * len(header)
        line[0], line[1], line[2], line[3] = i, code, f"عميل {code}", "مندوب"
        if len(line) > 11:
            line[5], line[11] = old, new
        ws.append(line)
    ws.append([None, None, "الإجمالي"])
    ws.append([None, None, "منها مستحق 7-10"])
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


@pytest.fixture(autouse=True)
def _no_real_collection_folder(tmp_path: Path, monkeypatch):
    """No test ever reads the project's real التحصيل folder: the command's default points into the test's own
    temporary folder (tmp_path / التحصيل)."""
    from sales_etl import cli

    monkeypatch.setattr(cli, "DEFAULT_COLLECTION_DIR", tmp_path / "التحصيل")
