import pandas as pd

from modules.comparison_engine import ComparisonEngine, generate_master_comparison


def test_compare_supports_unlimited_dynamic_suppliers_and_metrics():
    data = pd.DataFrame([
        {"item_no": "1", "description": "Desk", "unit": "ea", "quantity": 2, "unit_rate": 100, "total_amount": 200, "vendor": "Alpha"},
        {"item_no": "1", "description": "Desk", "unit": "ea", "quantity": 2, "unit_rate": 90, "total_amount": 180, "vendor": "Beta"},
        {"item_no": "1", "description": "Desk", "unit": "ea", "quantity": 2, "unit_rate": None, "total_amount": None, "vendor": "Gamma"},
    ])

    result = ComparisonEngine().compare(data)

    row = result.iloc[0]
    assert row["lowest_unit_rate"] == 90
    assert row["lowest_total_amount"] == 180
    assert row["lowest_supplier"] == "Beta"
    assert row["missing_prices"] == 1
    assert row["number_of_quotations"] == 3
    assert round(row["variance_percent"], 2) == 11.11
    assert {"Alpha total_amount", "Beta total_amount", "Gamma total_amount"}.issubset(result.columns)


def test_missing_item_number_matches_by_description_and_unit():
    data = pd.DataFrame([
        {"item_no": None, "description": "Generic service", "unit": "lot", "quantity": 1, "unit_rate": 50, "total_amount": 50, "vendor": "A"},
        {"item_no": "", "description": "Generic service", "unit": "lot", "quantity": 1, "unit_rate": 45, "total_amount": 45, "vendor": "B"},
    ])

    result = generate_master_comparison(data)

    assert len(result) == 1
    assert result.iloc[0]["lowest_supplier"] == "B"
    assert result.iloc[0]["number_of_quotations"] == 2


def test_configurable_matching_rules_can_ignore_item_number():
    data = pd.DataFrame([
        {"item_no": "A-1", "description": "Common item", "unit": "kg", "quantity": 3, "unit_rate": 10, "total_amount": 30, "vendor": "A"},
        {"item_no": "B-9", "description": "Common item", "unit": "kg", "quantity": 3, "unit_rate": 9, "total_amount": 27, "vendor": "B"},
    ])

    result = ComparisonEngine(matching_rules=("description", "unit")).compare(data)

    assert len(result) == 1
    assert result.iloc[0]["lowest_total_amount"] == 27
