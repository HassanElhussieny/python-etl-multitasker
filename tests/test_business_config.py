"""The company-specific business configuration (sales_etl/business.py): where it is read from, its validation, and
that the tests run on the synthetic example only (never on a local file with real names)."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from sales_etl import analysis, business, fixed_customers
from sales_etl.business import BusinessConfigError, load_business_config

EXAMPLE = Path(business.__file__).resolve().parents[1] / "config" / "business.example.json"


def test_tests_use_the_synthetic_example():
    assert business.BUSINESS.source.resolve() == EXAMPLE
    assert analysis.REPRESENTATIVES == business.BUSINESS.representatives
    assert fixed_customers.CUSTOMER_RESPONSIBLES == business.BUSINESS.customer_responsibles


def test_example_structure_matches_the_reports():
    cfg = load_business_config(EXAMPLE)
    assert len(cfg.sectors) == 8 and len(cfg.representatives) == 17 and len(cfg.customer_responsibles) == 11
    assert cfg.excluded_representative in cfg.representatives
    named = [n for g in cfg.cover_groups for n in g]
    assert len(named) == len(set(named)) == 16 and set(named) | {cfg.excluded_representative} == set(cfg.representatives)
    assert cfg.direct_addition not in cfg.representatives and cfg.direct_addition in cfg.cover_targets
    assert all(isinstance(v, int) and v > 0 for v in cfg.cover_targets.values())


def test_environment_variable_has_priority(tmp_path, monkeypatch):
    other = tmp_path / "other.json"
    other.write_text(EXAMPLE.read_text(encoding="utf-8"), encoding="utf-8")
    monkeypatch.setenv(business.ENV_VAR, str(other))
    assert business.config_path() == other


def test_local_file_is_used_before_the_example(tmp_path, monkeypatch):
    monkeypatch.delenv(business.ENV_VAR, raising=False)
    local = tmp_path / "business.local.json"
    monkeypatch.setattr(business, "LOCAL_FILE", local)
    monkeypatch.setattr(business, "EXAMPLE_FILE", EXAMPLE)
    assert business.config_path() == EXAMPLE                     # no local file: the example
    local.write_text("{}", encoding="utf-8")
    assert business.config_path() == local                       # local file present: it wins


@pytest.mark.parametrize("change, message", [
    (lambda d: d.pop("sectors"), "sectors"),
    (lambda d: d.update(representatives="x"), "representatives"),
    (lambda d: d.update(cover_groups=[["a"]]), "cover_groups"),
    (lambda d: d.update(cover_targets={"a": "ten"}), "cover_targets"),
    (lambda d: d.update(direct_addition=3), "direct_addition"),
])
def test_malformed_configuration_is_refused(tmp_path, change, message):
    data = json.loads(EXAMPLE.read_text(encoding="utf-8"))
    change(data)
    bad = tmp_path / "bad.json"
    bad.write_text(json.dumps(data), encoding="utf-8")
    with pytest.raises(BusinessConfigError, match=message):
        load_business_config(bad)


def test_missing_file_is_a_clear_error(tmp_path):
    with pytest.raises(BusinessConfigError, match="cannot read"):
        load_business_config(tmp_path / "missing.json")
