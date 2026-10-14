"""Regression tests for primary-key detection and Excel error values."""

import openpyxl

from app.analyzer import analyze_excel


def _write_xlsx(path, sheets: dict[str, list[tuple]]) -> str:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for row in rows:
            ws.append(row)
    wb.save(path)
    return str(path)


def _sheet(analysis, table_name):
    return next(s for s in analysis.sheets if s.table_name == table_name)


def _pk_names(sheet):
    return [c.python_name for c in sheet.columns if c.is_primary_key]


def test_unique_id_column_is_primary_key(tmp_path):
    path = _write_xlsx(tmp_path / "a.xlsx", {
        "products": [("productid", "name"), (1, "Chai"), (2, "Chang")],
    })
    assert _pk_names(_sheet(analyze_excel(path), "products")) == ["productid"]


def test_duplicate_values_are_not_primary_key(tmp_path):
    # Junction table: employeeid matches the PK name pattern but repeats.
    path = _write_xlsx(tmp_path / "a.xlsx", {
        "employeeterritories": [
            ("employeeid", "territoryid"),
            (1, "A"), (1, "A"), (2, "B"),
        ],
    })
    sheet = _sheet(analyze_excel(path), "employeeterritories")
    # No unique column -> generator synthesizes an autoincrement id.
    assert _pk_names(sheet) == []


def test_excel_error_values_count_as_empty(tmp_path):
    path = _write_xlsx(tmp_path / "a.xlsx", {
        "orders": [
            ("orderid", "shipregion"),
            (1, "RJ"), (2, "#N/A"), (3, "#DIV/0!"),
        ],
    })
    col = next(
        c for c in _sheet(analyze_excel(path), "orders").columns
        if c.python_name == "shipregion"
    )
    assert col.constraint.nullable is True
    assert "#N/A" not in col.sample_values
