"""Calculation Mode tools (3 operations)."""
import json
from excel_mcp.core.excel_com import get_excel_app, find_workbook
from excel_mcp.core.constants import xlCalculationAutomatic, xlCalculationManual, xlCalculationSemiautomatic


def get_calculation_mode(workbook: str = None) -> str:
    """Get the current calculation mode of the workbook."""
    try:
        app = get_excel_app()
        mode_map = {xlCalculationAutomatic: "automatic", xlCalculationManual: "manual",
                    xlCalculationSemiautomatic: "semiautomatic"}
        mode = mode_map.get(app.Calculation, "unknown")
        return json.dumps({"success": True, "mode": mode, "raw": app.Calculation})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def set_calculation_mode(mode: str) -> str:
    """Set calculation mode. mode: 'automatic' | 'manual' | 'semiautomatic'."""
    try:
        app = get_excel_app()
        mode_map = {"automatic": xlCalculationAutomatic, "manual": xlCalculationManual,
                    "semiautomatic": xlCalculationSemiautomatic}
        if mode not in mode_map:
            raise ValueError(f"Invalid mode: {mode!r}. Use 'automatic', 'manual', or 'semiautomatic'.")
        app.Calculation = mode_map[mode]
        return json.dumps({"success": True, "mode": mode})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def calculate(scope: str = "workbook", sheet: str = None, address: str = None,
              workbook: str = None) -> str:
    """Force recalculation. scope: 'workbook' | 'sheet' | 'range'."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        if scope == "workbook":
            try:
                wb.Calculate()
            except Exception:
                app.CalculateFull()  # WPS fallback
        elif scope == "sheet":
            from excel_mcp.core.excel_com import find_sheet
            find_sheet(wb, sheet).Calculate()
        elif scope == "range" and address and sheet:
            from excel_mcp.core.excel_com import find_sheet
            find_sheet(wb, sheet).Range(address).Calculate()
        else:
            app.Calculate()
        return json.dumps({"success": True, "scope": scope})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})
