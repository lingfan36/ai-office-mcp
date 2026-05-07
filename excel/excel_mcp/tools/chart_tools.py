"""Chart creation and configuration tools (29 operations)."""
import json
from excel_mcp.core.excel_com import get_excel_app, find_workbook, find_sheet, SnapshotUndo
from excel_mcp.core.constants import CHART_TYPES, TRENDLINE_TYPES


def _get_chart(ws, chart_name: str):
    for co in ws.ChartObjects():
        if co.Name == chart_name:
            return co.Chart
    raise ValueError(f"Chart not found: {chart_name!r}")


def _get_chart_object(ws, chart_name: str):
    for co in ws.ChartObjects():
        if co.Name == chart_name:
            return co
    raise ValueError(f"Chart not found: {chart_name!r}")


def list_charts(sheet: str = None, workbook: str = None) -> str:
    """List all charts on a sheet or the entire workbook."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        sheets = [find_sheet(wb, sheet)] if sheet else list(wb.Worksheets)
        results = []
        for ws in sheets:
            for co in ws.ChartObjects():
                c = co.Chart
                results.append({
                    "name": co.Name,
                    "sheet": ws.Name,
                    "chart_type": c.ChartType,
                    "left": co.Left, "top": co.Top,
                    "width": co.Width, "height": co.Height,
                })
        return json.dumps({"success": True, "charts": results, "count": len(results)})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def create_chart(source_address: str, chart_type: str = "column_clustered",
                 name: str = None, title: str = None,
                 left: float = None, top: float = None,
                 width: float = 400, height: float = 250,
                 sheet: str = None, workbook: str = None) -> str:
    """Create a chart from a data range.
    chart_type: column_clustered | bar_clustered | line | pie | scatter | area | doughnut | radar …
    """
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        src_rng = ws.Range(source_address)
        # position defaults: just below the data range
        if left is None:
            left = src_rng.Left
        if top is None:
            top = src_rng.Top + src_rng.Height + 10
        co = ws.ChartObjects().Add(left, top, width, height)
        if name:
            co.Name = name
        c = co.Chart
        c.SetSourceData(src_rng)
        c.ChartType = CHART_TYPES.get(chart_type, CHART_TYPES["column_clustered"])
        if title:
            c.HasTitle = True
            c.ChartTitle.Text = title
        return json.dumps({"success": True, "name": co.Name, "sheet": ws.Name,
                           "chart_type": chart_type, "title": title})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def create_chart_from_pivot(pivot_name: str, chart_type: str = "column_clustered",
                             name: str = None, title: str = None,
                             left: float = None, top: float = None,
                             width: float = 400, height: float = 250,
                             sheet: str = None, workbook: str = None) -> str:
    """Create a PivotChart from an existing PivotTable."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        pt = None
        for p in ws.PivotTables():
            if p.Name == pivot_name:
                pt = p
                break
        if pt is None:
            raise ValueError(f"PivotTable not found: {pivot_name!r}")
        src_rng = pt.TableRange1
        if left is None:
            left = src_rng.Left
        if top is None:
            top = src_rng.Top + src_rng.Height + 10
        co = ws.ChartObjects().Add(left, top, width, height)
        if name:
            co.Name = name
        c = co.Chart
        c.SetSourceData(src_rng)
        c.ChartType = CHART_TYPES.get(chart_type, CHART_TYPES["column_clustered"])
        if title:
            c.HasTitle = True
            c.ChartTitle.Text = title
        return json.dumps({"success": True, "name": co.Name, "pivot": pivot_name})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def set_chart_type(chart_name: str, chart_type: str, sheet: str = None, workbook: str = None) -> str:
    """Change the chart type."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        c = _get_chart(ws, chart_name)
        c.ChartType = CHART_TYPES.get(chart_type, CHART_TYPES["column_clustered"])
        return json.dumps({"success": True, "chart": chart_name, "type": chart_type})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def set_chart_title(chart_name: str, title: str, sheet: str = None, workbook: str = None) -> str:
    """Set or remove the chart title (pass title='' to remove)."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        c = _get_chart(ws, chart_name)
        if title:
            c.HasTitle = True
            c.ChartTitle.Text = title
        else:
            c.HasTitle = False
        return json.dumps({"success": True, "chart": chart_name, "title": title})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def set_axis_title(chart_name: str, axis: str, title: str,
                   sheet: str = None, workbook: str = None) -> str:
    """Set axis title. axis: 'x' | 'y' | 'z'."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        c = _get_chart(ws, chart_name)
        ax_map = {"x": (1, 1), "y": (2, 1), "z": (3, 1)}  # xlCategory/xlValue, xlPrimary
        ax_type, ax_group = ax_map.get(axis, (1, 1))
        ax = c.Axes(ax_type, ax_group)
        ax.HasTitle = True
        ax.AxisTitle.Text = title
        return json.dumps({"success": True, "chart": chart_name, "axis": axis, "title": title})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def set_axis_scale(chart_name: str, axis: str = "y",
                   min_val: float = None, max_val: float = None, unit: float = None,
                   sheet: str = None, workbook: str = None) -> str:
    """Configure axis scale (min, max, major unit)."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        c = _get_chart(ws, chart_name)
        ax = c.Axes(2 if axis == "y" else 1, 1)
        if min_val is not None:
            ax.MinimumScale = min_val
        if max_val is not None:
            ax.MaximumScale = max_val
        if unit is not None:
            ax.MajorUnit = unit
        return json.dumps({"success": True, "chart": chart_name, "axis": axis})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def set_axis_number_format(chart_name: str, axis: str = "y", number_format: str = "#,##0",
                           sheet: str = None, workbook: str = None) -> str:
    """Set number format on an axis."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        c = _get_chart(ws, chart_name)
        ax = c.Axes(2 if axis == "y" else 1, 1)
        ax.TickLabels.NumberFormat = number_format
        return json.dumps({"success": True, "chart": chart_name, "axis": axis, "format": number_format})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def toggle_legend(chart_name: str, show: bool = True, sheet: str = None, workbook: str = None) -> str:
    """Show or hide the chart legend."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        c = _get_chart(ws, chart_name)
        c.HasLegend = show
        return json.dumps({"success": True, "chart": chart_name, "legend": show})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def configure_data_labels(chart_name: str, show: bool = True, show_value: bool = True,
                          show_percent: bool = False, show_category: bool = False,
                          position: str = None, sheet: str = None, workbook: str = None) -> str:
    """Configure data labels on all series of a chart.
    position: 'center' | 'inside_end' | 'outside_end' | 'above' | 'below'
    """
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        c = _get_chart(ws, chart_name)
        pos_map = {"center": -4108, "inside_end": 2, "outside_end": 3, "above": -4162, "below": -4107}
        for s in c.SeriesCollection():
            s.HasDataLabels = show
            if show:
                dl = s.DataLabels()
                dl.ShowValue = show_value
                dl.ShowPercentage = show_percent
                dl.ShowCategoryName = show_category
                if position and position in pos_map:
                    dl.Position = pos_map[position]
        return json.dumps({"success": True, "chart": chart_name, "labels": show})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def add_trendline(chart_name: str, series_index: int = 1, trendline_type: str = "linear",
                  periods: int = 2, display_equation: bool = True, display_r_squared: bool = True,
                  sheet: str = None, workbook: str = None) -> str:
    """Add a trendline to a chart series.
    trendline_type: 'linear'|'exponential'|'logarithmic'|'polynomial'|'power'|'moving_average'
    """
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        c = _get_chart(ws, chart_name)
        series = c.SeriesCollection(series_index)
        tl_type = TRENDLINE_TYPES.get(trendline_type, TRENDLINE_TYPES["linear"])
        tl = series.Trendlines().Add(Type=tl_type)
        tl.DisplayEquation = display_equation
        tl.DisplayRSquared = display_r_squared
        if trendline_type == "moving_average":
            tl.Period = periods
        return json.dumps({"success": True, "chart": chart_name, "trendline": trendline_type})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def set_data_source(chart_name: str, source_address: str, sheet: str = None, workbook: str = None) -> str:
    """Change the data source range of a chart."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        c = _get_chart(ws, chart_name)
        c.SetSourceData(ws.Range(source_address))
        return json.dumps({"success": True, "chart": chart_name, "source": source_address})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def fit_chart_to_range(chart_name: str, address: str, sheet: str = None, workbook: str = None) -> str:
    """Align chart position and size to a cell range."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        co = _get_chart_object(ws, chart_name)
        rng = ws.Range(address)
        co.Left = rng.Left
        co.Top = rng.Top
        co.Width = rng.Width
        co.Height = rng.Height
        return json.dumps({"success": True, "chart": chart_name, "fitted_to": address})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def move_chart(chart_name: str, left: float, top: float, width: float = None, height: float = None,
               sheet: str = None, workbook: str = None) -> str:
    """Move (and optionally resize) a chart."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        co = _get_chart_object(ws, chart_name)
        co.Left = left
        co.Top = top
        if width is not None:
            co.Width = width
        if height is not None:
            co.Height = height
        return json.dumps({"success": True, "chart": chart_name})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def delete_chart(chart_name: str, sheet: str = None, workbook: str = None) -> str:
    """Delete a chart."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        _get_chart_object(ws, chart_name).Delete()
        return json.dumps({"success": True, "deleted": chart_name})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})
