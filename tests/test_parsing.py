"""Unit tests: dates, amounts, Arabic header matching, encoding/delimiter detection."""

from datetime import date
from decimal import Decimal

import pytest

from sales_etl.amounts import AmountParseError, parse_amount
from sales_etl.config import DEFAULT_DELIMITERS, DEFAULT_ENCODINGS
from sales_etl.dates import DateParseError, describe_date, parse_date, parse_reference_date
from sales_etl.ingest import read_raw_table
from sales_etl.models import FileProcessingError
from tests.conftest import ERP_HEADER, R, S, Txn, erp_csv_text, write_erp_csv


# ------------------------------------------------------------------ dates (scenario 10)
@pytest.mark.parametrize("text", ["08/10/2026", "08-10-2026", "8/10/2026", "08.10.2026", "2026-10-08",
                                  " 08/10/2026 ", "٠٨/١٠/٢٠٢٦"])
def test_reference_date_is_day_first(text):
    assert parse_reference_date(text) == date(2026, 10, 8)  # 8 October, never 10 August


def test_reference_date_description_is_unambiguous():
    assert describe_date(date(2026, 10, 8)) == "Thursday 08 October 2026 (ISO 2026-10-08)"


@pytest.mark.parametrize("text", ["", "abc", "31/02/2026", "13/13/2026", "08/10/26", "2026/10/08",
                                  "08/10-2026", "10/25/2026", "08/10/2026 10:00"])
def test_invalid_reference_dates_rejected(text):
    with pytest.raises(DateParseError):
        parse_reference_date(text)


def test_transaction_date_accepts_erp_dash_format_and_time_suffix():
    assert parse_date("26-09-2026") == date(2026, 9, 26)
    assert parse_date("26/09/2026 00:00:00") == date(2026, 9, 26)


# ------------------------------------------------------------------ amounts
@pytest.mark.parametrize("text,expected", [
    ("                       6,660.00", Decimal("6660.00")),
    ("-8,400.00", Decimal("-8400.00")),
    ("-166,660.00", Decimal("-166660.00")),
    ("0.00", Decimal("0.00")),
    ("1105646.62", Decimal("1105646.62")),
    ("٦٬٦٦٠٫٠٠", Decimal("6660.00")),
])
def test_amounts_parse_exactly(text, expected):
    assert parse_amount(text) == expected


@pytest.mark.parametrize("text", ["", "   ", "abc", "12,34.00", "1.234,56", "1,2345.00", "--5", "5-", "N/A"])
def test_invalid_amounts_rejected_not_zeroed(text):
    with pytest.raises(AmountParseError):
        parse_amount(text)


# ------------------------------------------------------------------ ingestion
def _read(path):
    return read_raw_table(path, encodings=DEFAULT_ENCODINGS, delimiters=DEFAULT_DELIMITERS)


@pytest.mark.parametrize("encoding,expected", [("cp1256", "cp1256"), ("utf-8", "utf-8"),
                                               ("utf-8-sig", "utf-8-sig"), ("utf-16", "utf-16")])
def test_encoding_detected_per_file(tmp_path, encoding, expected):
    path = write_erp_csv(tmp_path / "f.csv", [Txn("08-10-2026", S, "10")], encoding=encoding)
    raw = _read(path)
    assert raw.encoding == expected
    assert raw.header_cells[0] == "التاريخ"


@pytest.mark.parametrize("delimiter", [";", ",", "\t"])
def test_delimiter_detected_even_when_amounts_contain_commas(tmp_path, delimiter):
    path = write_erp_csv(tmp_path / "f.csv", [Txn("08-10-2026", S, "6660")], delimiter=delimiter)
    raw = _read(path)
    assert raw.delimiter == delimiter
    assert raw.header_line_no == 8  # after the 7-line ERP preamble


def test_header_located_after_preamble_and_blank_headers_named(tmp_path):
    raw = _read(write_erp_csv(tmp_path / "f.csv", [Txn("08-10-2026", S, "1")]))
    assert raw.columns[raw.column_index["صافى"]] == "صافى"
    assert raw.column_index["صافى"] == 19
    assert raw.columns[3].startswith("[عمود بدون عنوان")
    assert any(i.code == "BLANK_HEADERS" for i in raw.issues)


def test_arabic_spelling_variants_match_required_columns(tmp_path):
    header = list(ERP_HEADER)
    header[19] = "صافي"  # ي instead of ى
    header[0] = "التـاريخ"  # with tatweel
    raw = _read(write_erp_csv(tmp_path / "f.csv", [Txn("08-10-2026", S, "1")], header=header))
    assert raw.column_index["صافى"] == 19 and raw.column_index["التاريخ"] == 0
    assert raw.columns[19] == "صافي"  # original name preserved


def test_duplicate_required_column_is_a_file_failure(tmp_path):
    header = list(ERP_HEADER)
    header[10] = "صافى"  # a second صافى column
    path = write_erp_csv(tmp_path / "f.csv", [Txn("08-10-2026", S, "1")], header=header)
    with pytest.raises(FileProcessingError) as exc:
        _read(path)
    assert exc.value.code == "AMBIGUOUS_REQUIRED_COLUMN"


def test_missing_required_column_is_a_file_failure(tmp_path):
    header = list(ERP_HEADER)
    header[19] = "المبلغ"
    path = write_erp_csv(tmp_path / "f.csv", [Txn("08-10-2026", S, "1")], header=header)
    with pytest.raises(FileProcessingError) as exc:
        _read(path)
    assert exc.value.code == "MISSING_REQUIRED_COLUMNS" and "صافى" in exc.value.message


def test_empty_file_is_a_file_failure(tmp_path):
    path = tmp_path / "empty.csv"
    path.write_bytes(b"")
    with pytest.raises(FileProcessingError) as exc:
        _read(path)
    assert exc.value.code == "EMPTY_FILE"


def test_wrong_forced_encoding_reports_encoding_error(tmp_path):
    path = write_erp_csv(tmp_path / "f.csv", [Txn("08-10-2026", R, "-1")], encoding="cp1256")
    with pytest.raises(FileProcessingError) as exc:
        read_raw_table(path, encodings=DEFAULT_ENCODINGS, delimiters=DEFAULT_DELIMITERS,
                       forced_encoding="utf-8")
    assert exc.value.code == "ENCODING_ERROR"


def test_text_builder_matches_original_layout():
    text = erp_csv_text([Txn("26-09-2026", S, "4014.53", "2089", "400919")])
    data_line = text.splitlines()[8]
    assert data_line.startswith("26-09-2026;مبيعات;2089;93;;400919;")
    assert data_line.count(";") == 20
