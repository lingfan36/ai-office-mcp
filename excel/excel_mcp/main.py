"""Excel MCP Server — entry point and tool registration."""
import os
import sys
from dotenv import load_dotenv
from fastmcp import FastMCP

load_dotenv()

IS_WINDOWS = sys.platform == "win32"

_MODE = "COM (Windows + Excel/WPS)" if IS_WINDOWS else "openpyxl (cross-platform, no Office required)"
_EXTRA = (
    "Full feature set: pivot tables, slicers, VBA macros, screenshots, window management. "
    if IS_WINDOWS else
    "Core feature set: workbooks, ranges, sheets, named ranges, formatting, validation, "
    "autofilter, data diff. Advanced features (pivot tables, slicers, VBA, screenshots) "
    "require Windows + Excel. "
)

mcp = FastMCP(
    "Excel MCP Server",
    instructions=(
        f"AI-driven Excel automation — running in {_MODE} mode. "
        + _EXTRA
        + "Write operations automatically save a snapshot; call excel_undo to revert. "
        "Always pass workbook name when multiple files are open. "
        "Use excel_diff to compare any two .xlsx files without opening them."
    ),
)

# ═══════════════════════════════════════════════════════════════════════════════
# WINDOWS mode — full COM-based feature set
# ═══════════════════════════════════════════════════════════════════════════════
if IS_WINDOWS:
    # ── File / session ────────────────────────────────────────────────────────
    from excel_mcp.tools.file_tools import (
        list_sessions, open_workbook, close_workbook,
        create_workbook, save_workbook, get_workbook_info,
        protect_workbook, unprotect_workbook, export_pdf,
    )
    mcp.tool()(list_sessions)
    mcp.tool()(open_workbook)
    mcp.tool()(close_workbook)
    mcp.tool()(create_workbook)
    mcp.tool()(save_workbook)
    mcp.tool()(get_workbook_info)
    mcp.tool()(protect_workbook)
    mcp.tool()(unprotect_workbook)
    mcp.tool()(export_pdf)

    # ── Ranges ────────────────────────────────────────────────────────────────
    from excel_mcp.tools.range_tools import (
        get_values, set_values, get_formulas, set_formulas,
        clear_range, copy_range,
        insert_rows, delete_rows, insert_columns, delete_columns,
        find_replace, sort_range,
        get_used_range, get_current_region, get_range_info,
        get_number_format, set_number_format, format_range,
        add_validation, remove_validation,
        merge_cells, unmerge_cells, autofit,
        add_hyperlink, remove_hyperlinks,
        set_cell_lock,
        set_row_height, set_column_width,
        hide_rows, show_rows, hide_columns, show_columns,
        freeze_panes, unfreeze_panes,
        apply_range_autofilter, clear_range_autofilter, toggle_autofilter,
        group_rows, ungroup_rows, group_columns, ungroup_columns,
        remove_duplicates,
    )
    mcp.tool()(get_values)
    mcp.tool()(set_values)
    mcp.tool()(get_formulas)
    mcp.tool()(set_formulas)
    mcp.tool()(clear_range)
    mcp.tool()(copy_range)
    mcp.tool()(insert_rows)
    mcp.tool()(delete_rows)
    mcp.tool()(insert_columns)
    mcp.tool()(delete_columns)
    mcp.tool()(find_replace)
    mcp.tool()(sort_range)
    mcp.tool()(get_used_range)
    mcp.tool()(get_current_region)
    mcp.tool()(get_range_info)
    mcp.tool()(get_number_format)
    mcp.tool()(set_number_format)
    mcp.tool()(format_range)
    mcp.tool()(add_validation)
    mcp.tool()(remove_validation)
    mcp.tool()(merge_cells)
    mcp.tool()(unmerge_cells)
    mcp.tool()(autofit)
    mcp.tool()(add_hyperlink)
    mcp.tool()(remove_hyperlinks)
    mcp.tool()(set_cell_lock)
    mcp.tool()(set_row_height)
    mcp.tool()(set_column_width)
    mcp.tool()(hide_rows)
    mcp.tool()(show_rows)
    mcp.tool()(hide_columns)
    mcp.tool()(show_columns)
    mcp.tool()(freeze_panes)
    mcp.tool()(unfreeze_panes)
    mcp.tool()(apply_range_autofilter)
    mcp.tool()(clear_range_autofilter)
    mcp.tool()(toggle_autofilter)
    mcp.tool()(group_rows)
    mcp.tool()(ungroup_rows)
    mcp.tool()(group_columns)
    mcp.tool()(ungroup_columns)
    mcp.tool()(remove_duplicates)

    # ── Sheets ────────────────────────────────────────────────────────────────
    from excel_mcp.tools.sheet_tools import (
        list_sheets, create_sheet, rename_sheet, copy_sheet,
        move_sheet, delete_sheet, set_tab_color, set_sheet_visibility,
        copy_sheet_to_file, protect_sheet, unprotect_sheet,
    )
    mcp.tool()(list_sheets)
    mcp.tool()(create_sheet)
    mcp.tool()(rename_sheet)
    mcp.tool()(copy_sheet)
    mcp.tool()(move_sheet)
    mcp.tool()(delete_sheet)
    mcp.tool()(set_tab_color)
    mcp.tool()(set_sheet_visibility)
    mcp.tool()(copy_sheet_to_file)
    mcp.tool()(protect_sheet)
    mcp.tool()(unprotect_sheet)

    # ── Excel Tables ──────────────────────────────────────────────────────────
    from excel_mcp.tools.table_tools import (
        list_tables, create_table, rename_table, resize_table, delete_table,
        apply_table_style, toggle_totals_row, set_column_total,
        append_rows_to_table, get_table_data,
        apply_table_filter, clear_table_filters, sort_table,
        add_table_column, delete_table_column,
    )
    mcp.tool()(list_tables)
    mcp.tool()(create_table)
    mcp.tool()(rename_table)
    mcp.tool()(resize_table)
    mcp.tool()(delete_table)
    mcp.tool()(apply_table_style)
    mcp.tool()(toggle_totals_row)
    mcp.tool()(set_column_total)
    mcp.tool()(append_rows_to_table)
    mcp.tool()(get_table_data)
    mcp.tool()(apply_table_filter)
    mcp.tool()(clear_table_filters)
    mcp.tool()(sort_table)
    mcp.tool()(add_table_column)
    mcp.tool()(delete_table_column)

    # ── PivotTables ───────────────────────────────────────────────────────────
    from excel_mcp.tools.pivot_tools import (
        list_pivot_tables, create_pivot_table, create_pivot_from_table,
        add_pivot_field, remove_pivot_field,
        set_pivot_aggregation, set_pivot_field_name, set_pivot_number_format,
        filter_pivot_field, sort_pivot_field,
        add_calculated_field, list_pivot_fields,
        get_pivot_data, refresh_pivot, set_pivot_layout, delete_pivot,
    )
    mcp.tool()(list_pivot_tables)
    mcp.tool()(create_pivot_table)
    mcp.tool()(create_pivot_from_table)
    mcp.tool()(add_pivot_field)
    mcp.tool()(remove_pivot_field)
    mcp.tool()(set_pivot_aggregation)
    mcp.tool()(set_pivot_field_name)
    mcp.tool()(set_pivot_number_format)
    mcp.tool()(filter_pivot_field)
    mcp.tool()(sort_pivot_field)
    mcp.tool()(add_calculated_field)
    mcp.tool()(list_pivot_fields)
    mcp.tool()(get_pivot_data)
    mcp.tool()(refresh_pivot)
    mcp.tool()(set_pivot_layout)
    mcp.tool()(delete_pivot)

    # ── Charts ────────────────────────────────────────────────────────────────
    from excel_mcp.tools.chart_tools import (
        list_charts, create_chart, create_chart_from_pivot,
        set_chart_type, set_chart_title, set_axis_title,
        set_axis_scale, set_axis_number_format, toggle_legend,
        configure_data_labels, add_trendline,
        set_data_source, fit_chart_to_range, move_chart, delete_chart,
    )
    mcp.tool()(list_charts)
    mcp.tool()(create_chart)
    mcp.tool()(create_chart_from_pivot)
    mcp.tool()(set_chart_type)
    mcp.tool()(set_chart_title)
    mcp.tool()(set_axis_title)
    mcp.tool()(set_axis_scale)
    mcp.tool()(set_axis_number_format)
    mcp.tool()(toggle_legend)
    mcp.tool()(configure_data_labels)
    mcp.tool()(add_trendline)
    mcp.tool()(set_data_source)
    mcp.tool()(fit_chart_to_range)
    mcp.tool()(move_chart)
    mcp.tool()(delete_chart)

    # ── Named Ranges ──────────────────────────────────────────────────────────
    from excel_mcp.tools.named_range_tools import (
        list_named_ranges, read_named_range, write_named_range,
        create_named_range, update_named_range, delete_named_range,
    )
    mcp.tool()(list_named_ranges)
    mcp.tool()(read_named_range)
    mcp.tool()(write_named_range)
    mcp.tool()(create_named_range)
    mcp.tool()(update_named_range)
    mcp.tool()(delete_named_range)

    # ── Conditional Formatting ────────────────────────────────────────────────
    from excel_mcp.tools.format_tools import (
        add_conditional_format, clear_conditional_formats,
    )
    mcp.tool()(add_conditional_format)
    mcp.tool()(clear_conditional_formats)

    # ── Slicers ───────────────────────────────────────────────────────────────
    from excel_mcp.tools.slicer_tools import (
        create_pivot_slicer, list_pivot_slicers,
        set_pivot_slicer_selection, delete_slicer,
        create_table_slicer, set_table_slicer_selection,
    )
    mcp.tool()(create_pivot_slicer)
    mcp.tool()(list_pivot_slicers)
    mcp.tool()(set_pivot_slicer_selection)
    mcp.tool()(delete_slicer)
    mcp.tool()(create_table_slicer)
    mcp.tool()(set_table_slicer_selection)

    # ── Calculation ───────────────────────────────────────────────────────────
    from excel_mcp.tools.calc_tools import (
        get_calculation_mode, set_calculation_mode, calculate,
    )
    mcp.tool()(get_calculation_mode)
    mcp.tool()(set_calculation_mode)
    mcp.tool()(calculate)

    # ── Screenshots ───────────────────────────────────────────────────────────
    from excel_mcp.tools.screenshot_tools import capture_range, capture_sheet
    mcp.tool()(capture_range)
    mcp.tool()(capture_sheet)

    # ── Window Management ─────────────────────────────────────────────────────
    from excel_mcp.tools.window_tools import (
        show_excel, hide_excel, get_window_info,
        set_window_state, set_window_position, arrange_window,
        set_status_bar, clear_status_bar,
    )
    mcp.tool()(show_excel)
    mcp.tool()(hide_excel)
    mcp.tool()(get_window_info)
    mcp.tool()(set_window_state)
    mcp.tool()(set_window_position)
    mcp.tool()(arrange_window)
    mcp.tool()(set_status_bar)
    mcp.tool()(clear_status_bar)

    # ── VBA ───────────────────────────────────────────────────────────────────
    from excel_mcp.tools.vba_tools import (
        list_vba_modules, view_vba_code, import_vba_module,
        update_vba_code, delete_vba_module, run_macro,
    )
    mcp.tool()(list_vba_modules)
    mcp.tool()(view_vba_code)
    mcp.tool()(import_vba_module)
    mcp.tool()(update_vba_code)
    mcp.tool()(delete_vba_module)
    mcp.tool()(run_macro)

    # ── Comments ─────────────────────────────────────────────────────────────
    from excel_mcp.tools.comment_tools import (
        add_comment, get_comments, delete_comment, show_hide_comment,
    )
    mcp.tool()(add_comment)
    mcp.tool()(get_comments)
    mcp.tool()(delete_comment)
    mcp.tool()(show_hide_comment)

    # ── Images ────────────────────────────────────────────────────────────────
    from excel_mcp.tools.image_tools import (
        insert_image, list_images, delete_image,
    )
    mcp.tool()(insert_image)
    mcp.tool()(list_images)
    mcp.tool()(delete_image)

    # ── Undo ─────────────────────────────────────────────────────────────────
    from excel_mcp.tools.undo_tools import (
        undo_last, list_undo_snapshots, clear_undo_history,
    )
    mcp.tool()(undo_last)
    mcp.tool()(list_undo_snapshots)
    mcp.tool()(clear_undo_history)

# ═══════════════════════════════════════════════════════════════════════════════
# CROSS-PLATFORM mode — openpyxl backend (macOS / Linux / Windows without Excel)
# ═══════════════════════════════════════════════════════════════════════════════
else:
    # ── File / session ────────────────────────────────────────────────────────
    from excel_mcp.tools.compat_file_tools import (
        list_sessions, open_workbook, close_workbook,
        create_workbook, save_workbook, get_workbook_info,
        protect_workbook, unprotect_workbook, export_pdf,
    )
    mcp.tool()(list_sessions)
    mcp.tool()(open_workbook)
    mcp.tool()(close_workbook)
    mcp.tool()(create_workbook)
    mcp.tool()(save_workbook)
    mcp.tool()(get_workbook_info)
    mcp.tool()(protect_workbook)
    mcp.tool()(unprotect_workbook)
    mcp.tool()(export_pdf)

    # ── Ranges ────────────────────────────────────────────────────────────────
    from excel_mcp.tools.compat_range_tools import (
        get_values, set_values, get_formulas, set_formulas,
        clear_range, copy_range,
        insert_rows, delete_rows, insert_columns, delete_columns,
        find_replace, sort_range,
        get_used_range, get_current_region, get_range_info,
        get_number_format, set_number_format, format_range,
        add_validation, remove_validation,
        merge_cells, unmerge_cells, autofit,
        add_hyperlink, remove_hyperlinks,
        set_cell_lock,
        set_row_height, set_column_width,
        hide_rows, show_rows, hide_columns, show_columns,
        freeze_panes, unfreeze_panes,
        apply_range_autofilter, clear_range_autofilter, toggle_autofilter,
        group_rows, ungroup_rows, group_columns, ungroup_columns,
        remove_duplicates,
    )
    mcp.tool()(get_values)
    mcp.tool()(set_values)
    mcp.tool()(get_formulas)
    mcp.tool()(set_formulas)
    mcp.tool()(clear_range)
    mcp.tool()(copy_range)
    mcp.tool()(insert_rows)
    mcp.tool()(delete_rows)
    mcp.tool()(insert_columns)
    mcp.tool()(delete_columns)
    mcp.tool()(find_replace)
    mcp.tool()(sort_range)
    mcp.tool()(get_used_range)
    mcp.tool()(get_current_region)
    mcp.tool()(get_range_info)
    mcp.tool()(get_number_format)
    mcp.tool()(set_number_format)
    mcp.tool()(format_range)
    mcp.tool()(add_validation)
    mcp.tool()(remove_validation)
    mcp.tool()(merge_cells)
    mcp.tool()(unmerge_cells)
    mcp.tool()(autofit)
    mcp.tool()(add_hyperlink)
    mcp.tool()(remove_hyperlinks)
    mcp.tool()(set_cell_lock)
    mcp.tool()(set_row_height)
    mcp.tool()(set_column_width)
    mcp.tool()(hide_rows)
    mcp.tool()(show_rows)
    mcp.tool()(hide_columns)
    mcp.tool()(show_columns)
    mcp.tool()(freeze_panes)
    mcp.tool()(unfreeze_panes)
    mcp.tool()(apply_range_autofilter)
    mcp.tool()(clear_range_autofilter)
    mcp.tool()(toggle_autofilter)
    mcp.tool()(group_rows)
    mcp.tool()(ungroup_rows)
    mcp.tool()(group_columns)
    mcp.tool()(ungroup_columns)
    mcp.tool()(remove_duplicates)

    # ── Sheets ────────────────────────────────────────────────────────────────
    from excel_mcp.tools.compat_sheet_tools import (
        list_sheets, create_sheet, rename_sheet, copy_sheet,
        move_sheet, delete_sheet, set_tab_color, set_sheet_visibility,
        copy_sheet_to_file, protect_sheet, unprotect_sheet,
    )
    mcp.tool()(list_sheets)
    mcp.tool()(create_sheet)
    mcp.tool()(rename_sheet)
    mcp.tool()(copy_sheet)
    mcp.tool()(move_sheet)
    mcp.tool()(delete_sheet)
    mcp.tool()(set_tab_color)
    mcp.tool()(set_sheet_visibility)
    mcp.tool()(copy_sheet_to_file)
    mcp.tool()(protect_sheet)
    mcp.tool()(unprotect_sheet)

    # ── Undo ─────────────────────────────────────────────────────────────────
    from excel_mcp.tools.compat_undo_tools import (
        undo_last, list_undo_snapshots, clear_undo_history,
    )
    mcp.tool()(undo_last)
    mcp.tool()(list_undo_snapshots)
    mcp.tool()(clear_undo_history)

# ═══════════════════════════════════════════════════════════════════════════════
# ALWAYS available — cross-platform tools
# ═══════════════════════════════════════════════════════════════════════════════
from excel_mcp.tools.diff_tools import excel_diff
mcp.tool()(excel_diff)


def run_server():
    transport = os.environ.get("MCP_TRANSPORT", "stdio")
    if transport == "streamable-http":
        mcp.run(
            transport="streamable-http",
            host=os.environ.get("MCP_HOST", "127.0.0.1"),
            port=int(os.environ.get("MCP_PORT", "8001")),
            path=os.environ.get("MCP_PATH", "/mcp"),
        )
    else:
        mcp.run(transport="stdio")


if __name__ == "__main__":
    run_server()
