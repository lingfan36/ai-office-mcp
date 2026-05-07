"""VBA macro tools for open Word documents via COM."""

import json
import sys


async def word_live_list_macros(filename: str = None) -> str:
    """List all VBA macros (Sub procedures) accessible from an open Word document.

    Scans Normal.dotm (global macros) and the document's own VBA project.

    Args:
        filename: Document name or path (None = active document, used to access its project).

    Returns:
        JSON array of macros with module name and procedure name.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "VBA tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        macros = []

        def _scan_project(vba_project, scope):
            try:
                for i in range(1, vba_project.VBComponents.Count + 1):
                    comp = vba_project.VBComponents(i)
                    cm = comp.CodeModule
                    line = 1
                    total_lines = cm.CountOfLines
                    while line <= total_lines:
                        try:
                            proc_name = cm.ProcOfLine(line, 0)  # 0 = vbext_pk_Proc
                            if proc_name:
                                # Check it's a Sub (not a Function or Property)
                                start = cm.ProcStartLine(proc_name, 0)
                                code_line = cm.Lines(start, 1).strip()
                                proc_type = "Sub" if code_line.lower().startswith("sub ") else "Function"
                                entry = {
                                    "scope": scope,
                                    "module": comp.Name,
                                    "macro": proc_name,
                                    "type": proc_type,
                                    "call": f"{comp.Name}.{proc_name}",
                                }
                                if entry not in macros:
                                    macros.append(entry)
                                # Skip to end of this procedure
                                proc_lines = cm.ProcCountLines(proc_name, 0)
                                line += max(proc_lines, 1)
                            else:
                                line += 1
                        except Exception:
                            line += 1
            except Exception:
                pass

        # Document project
        try:
            _scan_project(doc.VBProject, "document")
        except Exception:
            pass

        # Normal template project
        try:
            _scan_project(app.NormalTemplate.VBProject, "Normal.dotm")
        except Exception:
            pass

        return json.dumps({
            "success": True,
            "count": len(macros),
            "macros": macros,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_run_macro(
    macro_name: str,
    filename: str = None,
    args: list = None,
) -> str:
    """Run a VBA macro in the context of an open Word document.

    Args:
        macro_name: Macro name, optionally qualified as "Module.MacroName".
                    Use word_live_list_macros to discover available macros.
        filename: Document name or path (used to set ActiveDocument context).
        args: List of string arguments to pass to the macro (optional).

    Returns:
        JSON confirming the macro ran, or an error if it failed.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "VBA tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        if filename:
            doc = find_document(app, filename)
            doc.Activate()

        call_args = [macro_name] + [str(a) for a in (args or [])]
        app.Run(*call_args)

        return json.dumps({
            "success": True,
            "macro": macro_name,
            "args": args or [],
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_get_macro_code(
    module_name: str,
    filename: str = None,
    scope: str = "document",
) -> str:
    """Get the VBA source code of a module in an open Word document.

    Args:
        module_name: Name of the VBA module (from word_live_list_macros).
        filename: Document name or path (None = active document).
        scope: "document" to look in the document project, "normal" for Normal.dotm.

    Returns:
        JSON with the full module source code.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "VBA tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()

        if scope == "normal":
            project = app.NormalTemplate.VBProject
        else:
            doc = find_document(app, filename)
            project = doc.VBProject

        comp = project.VBComponents(module_name)
        code = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)

        return json.dumps({
            "success": True,
            "module": module_name,
            "scope": scope,
            "lines": comp.CodeModule.CountOfLines,
            "code": code,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})
