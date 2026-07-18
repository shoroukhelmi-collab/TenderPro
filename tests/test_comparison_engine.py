import pandas as pd

from modules.comparison_engine import build_comparison, supplier_ranking


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
