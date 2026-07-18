import pandas as pd

from modules.comparison_engine import build_comparison, dashboard_metrics, detect_outliers, missing_items, package_summary, supplier_ranking


def test_build_comparison_shows_lowest_price_and_missing_prices():
    data = pd.DataFrame([
        {"item_no": "1", "description": "Desk", "unit": "ea", "quantity": 2, "unit_rate": 100, "total_amount": 200, "supplier": "Supplier 1"},
        {"item_no": "1", "description": "Desk", "unit": "ea", "quantity": 2, "unit_rate": 90, "total_amount": 180, "supplier": "Supplier 2"},
        {"item_no": "1", "description": "Desk", "unit": "ea", "quantity": 2, "unit_rate": None, "total_amount": None, "supplier": "Supplier 3"},
    ])

    result = build_comparison(data)

    row = result.iloc[0]
    assert row["lowest_supplier"] == "Supplier 2"
    assert row["lowest_price"] == 180
    assert row["missing_prices"] == 1
    assert {"Supplier 1 price", "Supplier 2 price", "Supplier 3 price"}.issubset(result.columns)


def test_missing_item_number_matches_by_description_and_unit():
    data = pd.DataFrame([
        {"item_no": None, "description": "Generic service", "unit": "lot", "quantity": 1, "unit_rate": 50, "total_amount": 50, "supplier": "Supplier 1"},
        {"item_no": "", "description": "Generic service", "unit": "lot", "quantity": 1, "unit_rate": 45, "total_amount": 45, "supplier": "Supplier 2"},
    ])

    result = build_comparison(data)

    assert len(result) == 1
    assert result.iloc[0]["lowest_supplier"] == "Supplier 2"


def test_supplier_ranking_orders_by_total_and_missing_prices():
    data = pd.DataFrame([
        {"item_no": "1", "description": "A", "unit": "ea", "quantity": 1, "unit_rate": 20, "total_amount": 20, "supplier": "Supplier 1"},
        {"item_no": "1", "description": "A", "unit": "ea", "quantity": 1, "unit_rate": 10, "total_amount": 10, "supplier": "Supplier 2"},
        {"item_no": "2", "description": "B", "unit": "ea", "quantity": 1, "unit_rate": None, "total_amount": None, "supplier": "Supplier 2"},
    ])

    ranking = supplier_ranking(data)

    assert ranking.iloc[0]["supplier"] == "Supplier 2"
    assert ranking.iloc[0]["rank"] == 1
    assert ranking.iloc[0]["missing_prices"] == 1



def test_outliers_can_be_excluded_without_removing_raw_rows():
    data = pd.DataFrame([
        {"item_no": "1.01", "description": "A", "unit": "ea", "quantity": 1, "unit_rate": 10, "total_amount": 10, "supplier": "Supplier 1"},
        {"item_no": "1.01", "description": "A", "unit": "ea", "quantity": 1, "unit_rate": 11, "total_amount": 11, "supplier": "Supplier 2"},
        {"item_no": "1.01", "description": "A", "unit": "ea", "quantity": 1, "unit_rate": 100, "total_amount": 100, "supplier": "Supplier 3"},
    ])

    outliers = detect_outliers(data)
    excluded_key = outliers.iloc[0]["outlier_key"]
    comparison = build_comparison(data, {excluded_key})

    assert len(outliers) == 1
    assert comparison.iloc[0]["outlier_count"] == 1
    assert comparison.iloc[0]["excluded_outliers"] == 1
    assert data["total_amount"].sum() == 121


def test_package_missing_and_dashboard_summaries_are_generic():
    data = pd.DataFrame([
        {"item_no": "A-1", "description": "Alpha", "unit": "ea", "quantity": 1, "unit_rate": 10, "total_amount": 10, "supplier": "Supplier 1"},
        {"item_no": "A-1", "description": "Alpha", "unit": "ea", "quantity": 1, "unit_rate": None, "total_amount": None, "supplier": "Supplier 2"},
    ])

    packages = package_summary(data)
    missing = missing_items(data)
    metrics = dashboard_metrics(data, uploaded_file_count=2)

    assert packages.iloc[0]["package"] == "a"
    assert len(missing) == 1
    assert metrics["total_uploaded_files"] == 2
    assert metrics["missing_prices"] == 1


def test_comparison_matches_by_package_item_description_and_unit():
    data = pd.DataFrame([
        {"package": "HVAC", "item_no": "1", "description": "Fan", "unit": "ea", "quantity": 1, "unit_rate": 10, "total_amount": 10, "supplier": "Hazek", "file_name": "HVAC.xlsx", "worksheet": "BOQ"},
        {"package": "Fire Alarm", "item_no": "1", "description": "Fan", "unit": "ea", "quantity": 1, "unit_rate": 20, "total_amount": 20, "supplier": "Hazek", "file_name": "Fire Alarm.xlsx", "worksheet": "BOQ"},
        {"package": "HVAC", "item_no": "1", "description": "Fan", "unit": "ea", "quantity": 1, "unit_rate": 9, "total_amount": 9, "supplier": "Orascom", "file_name": "HVAC.xlsx", "worksheet": "BOQ"},
        {"package": "HVAC", "item_no": "1", "description": "Different fan", "unit": "ea", "quantity": 1, "unit_rate": 8, "total_amount": 8, "supplier": "Other", "file_name": "HVAC.xlsx", "worksheet": "BOQ"},
    ])

    result = build_comparison(data)

    assert len(result) == 3
    hvac_fan = result[(result["package"] == "HVAC") & (result["description"] == "Fan")].iloc[0]
    assert hvac_fan["lowest_supplier"] == "Orascom"
    assert "file_name" in result.columns
    assert "worksheet" in result.columns
