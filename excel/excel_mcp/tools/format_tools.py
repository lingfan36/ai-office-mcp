"""Conditional Formatting tools (2 operations)."""
import json
from excel_mcp.core.excel_com import get_excel_app, find_workbook, find_sheet, SnapshotUndo
from excel_mcp.core.constants import xlCellValue, xlExpression, xlColorScale, xlDataBar, xlIconSet


def add_conditional_format(address: str, rule_type: str, formula: str = None,
                            operator: str = "greater", value: float = None, value2: float = None,
                            font_color: str = None, bg_color: str = None,
                            sheet: str = None, workbook: str = None) -> str:
    """Add a conditional formatting rule.

    rule_type: 'cell_value' | 'formula' | 'color_scale' | 'data_bar' | 'icon_set'
    operator (for cell_value): 'greater'|'less'|'equal'|'between'|'not_between'|'greater_equal'|'less_equal'
    Colors as '#RRGGBB'.
    """
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        rng = ws.Range(address)

        def _color(h):
            h = h.lstrip("#")
            r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
            return r + g * 256 + b * 65536

        op_map = {"greater": 5, "less": 6, "equal": 3, "between": 1,
                  "not_between": 2, "greater_equal": 7, "less_equal": 8}

        if rule_type == "cell_value":
            op = op_map.get(operator, 5)
            kwargs = {"Type": xlCellValue, "Operator": op, "Formula1": str(value)}
            if value2 is not None:
                kwargs["Formula2"] = str(value2)
            fc = rng.FormatConditions.Add(**kwargs)
            if font_color:
                fc.Font.Color = _color(font_color)
            if bg_color:
                fc.Interior.Color = _color(bg_color)

        elif rule_type == "formula":
            fc = rng.FormatConditions.Add(Type=xlExpression, Formula1=formula)
            if font_color:
                fc.Font.Color = _color(font_color)
            if bg_color:
                fc.Interior.Color = _color(bg_color)

        elif rule_type == "color_scale":
            rng.FormatConditions.AddColorScale(ColorScaleType=3)

        elif rule_type == "data_bar":
            rng.FormatConditions.AddDatabar()

        elif rule_type == "icon_set":
            rng.FormatConditions.AddIconSetCondition()

        return json.dumps({"success": True, "address": address, "rule_type": rule_type})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def clear_conditional_formats(address: str, sheet: str = None, workbook: str = None) -> str:
    """Remove all conditional formatting rules from a range."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        ws.Range(address).FormatConditions.Delete()
        return json.dumps({"success": True, "cleared": address})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})
