"""VBA Macro tools (6 operations). Requires .xlsm workbooks."""
import json
from excel_mcp.core.excel_com import get_excel_app, find_workbook


def list_vba_modules(workbook: str = None) -> str:
    """List all VBA components (modules, class modules, user forms) in a workbook."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        modules = []
        for comp in wb.VBProject.VBComponents:
            type_map = {1: "module", 2: "class_module", 3: "user_form", 100: "sheet_or_workbook"}
            modules.append({
                "name": comp.Name,
                "type": type_map.get(comp.Type, "unknown"),
                "lines": comp.CodeModule.CountOfLines,
            })
        return json.dumps({"success": True, "modules": modules, "count": len(modules)})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def view_vba_code(module_name: str, workbook: str = None) -> str:
    """Retrieve the source code of a VBA module."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        comp = wb.VBProject.VBComponents(module_name)
        code = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
        return json.dumps({"success": True, "module": module_name, "code": code,
                           "lines": comp.CodeModule.CountOfLines})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def import_vba_module(module_name: str, code: str, workbook: str = None) -> str:
    """Create or replace a standard VBA module with the given code."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        vbp = wb.VBProject
        # Remove existing module with same name if present
        for comp in list(vbp.VBComponents):
            if comp.Name == module_name and comp.Type == 1:
                vbp.VBComponents.Remove(comp)
                break
        new_comp = vbp.VBComponents.Add(1)  # vbext_ct_StdModule
        new_comp.Name = module_name
        new_comp.CodeModule.AddFromString(code)
        return json.dumps({"success": True, "module": module_name,
                           "lines": new_comp.CodeModule.CountOfLines})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def update_vba_code(module_name: str, code: str, workbook: str = None) -> str:
    """Replace the entire code of an existing VBA module."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        comp = wb.VBProject.VBComponents(module_name)
        cm = comp.CodeModule
        if cm.CountOfLines > 0:
            cm.DeleteLines(1, cm.CountOfLines)
        cm.AddFromString(code)
        return json.dumps({"success": True, "module": module_name, "lines": cm.CountOfLines})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def delete_vba_module(module_name: str, workbook: str = None) -> str:
    """Delete a VBA component (module, class module, or user form)."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        comp = wb.VBProject.VBComponents(module_name)
        wb.VBProject.VBComponents.Remove(comp)
        return json.dumps({"success": True, "deleted": module_name})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def run_macro(macro_name: str, args: list = None, workbook: str = None) -> str:
    """Execute a VBA macro by name. macro_name can include module: 'Module1.MyMacro'."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        full_name = f"'{wb.Name}'!{macro_name}"
        if args:
            result = app.Run(full_name, *args)
        else:
            result = app.Run(full_name)
        return json.dumps({"success": True, "macro": macro_name, "result": str(result) if result is not None else None})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})
