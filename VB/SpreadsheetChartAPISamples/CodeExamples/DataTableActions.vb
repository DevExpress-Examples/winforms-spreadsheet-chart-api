Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Charts

Namespace SpreadsheetChartAPIActions
    Module DataTableActions
        Private Sub ShowDataTables(ByVal workbook As IWorkbook)
#Region "#ShowDataTable"
            Dim worksheet As Worksheet = workbook.Worksheets("chartTask5")
            workbook.Worksheets.ActiveWorksheet = worksheet

            'Create a chart and specify its location
            Dim chart As Chart = worksheet.Charts.Add(ChartType.Line, worksheet("B2:C8"))
            chart.TopLeftCell = worksheet.Cells("F2")
            chart.BottomRightCell = worksheet.Cells("L14")
            Dim dataTableOptions As DataTableOptions = chart.DataTable
            dataTableOptions.Visible = True
            dataTableOptions.ShowLegendKeys = False
#End Region
        End Sub

        Private Sub ChangeDataTableBorders(ByVal workbook As IWorkbook)
#Region "#ChangeDataTableBorders"
            Dim worksheet As Worksheet = workbook.Worksheets("chartTask5")
            workbook.Worksheets.ActiveWorksheet = worksheet

            'Create a chart and specify its location
            Dim chart As Chart = worksheet.Charts.Add(ChartType.Line, worksheet("B2:C8"))
            chart.TopLeftCell = worksheet.Cells("F2")
            chart.BottomRightCell = worksheet.Cells("L14")
            Dim dataTableOptions As DataTableOptions = chart.DataTable
            dataTableOptions.Visible = True
            dataTableOptions.ShowLegendKeys = False

            dataTableOptions.ShowVerticalBorder = False
            dataTableOptions.ShowOutline = False
#End Region
        End Sub

        Private Sub ChangeDataTableFont(ByVal workbook As IWorkbook)
#Region "#ChangeDataTableFont"
            Dim worksheet As Worksheet = workbook.Worksheets("chartTask5")
            workbook.Worksheets.ActiveWorksheet = worksheet

            'Create a chart and specify its location
            Dim chart As Chart = worksheet.Charts.Add(ChartType.Line, worksheet("B2:C8"))
            chart.TopLeftCell = worksheet.Cells("F2")
            chart.BottomRightCell = worksheet.Cells("L14")
            Dim dataTableOptions As DataTableOptions = chart.DataTable
            dataTableOptions.Visible = True
            dataTableOptions.ShowLegendKeys = False

            dataTableOptions.Font.Name = "Helvetica"
            dataTableOptions.Font.Size = 12
#End Region
        End Sub
    End Module
End Namespace
