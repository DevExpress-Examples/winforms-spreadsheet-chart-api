Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Charts
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace SpreadsheetChartAPIActions

    Public Module Protection

        Private Sub ProtectChart(ByVal workbook As DevExpress.Spreadsheet.IWorkbook)
'#Region "#ProtectChart"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("chartTask3")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Create a chart and specify its location.
            Dim chart As DevExpress.Spreadsheet.Charts.Chart = worksheet.Charts.Add(DevExpress.Spreadsheet.Charts.ChartType.ColumnClustered, worksheet("B2:D4"))
            chart.TopLeftCell = worksheet.Cells("H2")
            chart.BottomRightCell = worksheet.Cells("N14")
            ' Specify the chart style.
            chart.Style = DevExpress.Spreadsheet.Charts.ChartStyle.ColorDark
            ' Apply the chart protection.
            chart.Options.Protection = DevExpress.Spreadsheet.Charts.ChartProtection.All
'#End Region  ' #ProtectChart
        End Sub
    End Module
End Namespace
