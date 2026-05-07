"""Excel COM enumeration constants used across tools."""

# XlCalculation
xlCalculationAutomatic = -4105
xlCalculationManual = -4135
xlCalculationSemiautomatic = 2

# XlSaveAsAccessMode / format
xlOpenXMLWorkbook = 51        # .xlsx
xlOpenXMLWorkbookMacroEnabled = 52  # .xlsm
xlCSV = 6

# XlDirection
xlDown = -4121
xlUp = -4162
xlToRight = -4161
xlToLeft = -4159

# XlBorderIndex
xlEdgeLeft = 7
xlEdgeTop = 8
xlEdgeBottom = 9
xlEdgeRight = 10
xlInsideVertical = 11
xlInsideHorizontal = 12
xlDiagonalDown = 5
xlDiagonalUp = 6

# XlLineStyle
xlContinuous = 1
xlDash = -4115
xlDot = -4118
xlDouble = -4119
xlNone = -4142

# XlBorderWeight
xlThin = 2
xlMedium = -4138
xlThick = 4
xlHairline = 1

# XlHAlign
xlHAlignCenter = -4108
xlHAlignLeft = -4131
xlHAlignRight = -4152
xlHAlignGeneral = 1

# XlVAlign
xlVAlignBottom = -4107
xlVAlignCenter = -4108
xlVAlignTop = -4160

# XlSortOrder
xlAscending = 1
xlDescending = 2

# XlYesNoGuess
xlGuess = 0
xlYes = 1
xlNo = 2

# XlChartType
CHART_TYPES = {
    "column_clustered": 51,
    "column_stacked": 52,
    "column_stacked_100": 53,
    "bar_clustered": 57,
    "bar_stacked": 58,
    "bar_stacked_100": 59,
    "line": 4,
    "line_markers": 65,
    "line_stacked": 63,
    "line_stacked_markers": 66,
    "pie": 5,
    "pie_exploded": 69,
    "area": 1,
    "area_stacked": 76,
    "scatter": -4169,
    "scatter_lines": 74,
    "scatter_smooth": 73,
    "bubble": 15,
    "doughnut": -4120,
    "radar": -4151,
    "stock_ohlc": 59,
    "combo": -4169,
}

# XlTrendlineType
TRENDLINE_TYPES = {
    "linear": -4132,
    "exponential": 5,
    "logarithmic": -4133,
    "polynomial": 3,
    "power": 4,
    "moving_average": 6,
}

# XlPivotFieldOrientation
xlRowField = 1
xlColumnField = 2
xlPageField = 3
xlDataField = 4
xlHidden = 0

# XlConsolidationFunction
AGGREGATION = {
    "sum": -4157,
    "count": -4112,
    "average": -4106,
    "max": -4136,
    "min": -4139,
    "product": -4149,
    "count_numbers": -4113,
    "std_dev": -4155,
    "std_devp": -4156,
    "var": -4164,
    "varp": -4165,
}

# XlWindowState
xlNormal = -4143
xlMinimized = -4140
xlMaximized = -4137

# XlSheetVisibility
xlSheetVisible = -1
xlSheetHidden = 0
xlSheetVeryHidden = 2

# XlDVType (Data Validation)
xlValidateList = 3
xlValidateWholeNumber = 1
xlValidateDecimal = 2
xlValidateDate = 4
xlValidateTime = 5
xlValidateTextLength = 6
xlValidateCustom = 7

# XlFormatConditionType
xlCellValue = 1
xlExpression = 2
xlColorScale = 3
xlDataBar = 4
xlIconSet = 6
xlTop10 = 5
xlBlanksCondition = 10
xlNoBlanksCondition = 13
