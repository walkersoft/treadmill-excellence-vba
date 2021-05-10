Attribute VB_Name = "Module1"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    ActiveChart.PivotLayout.PivotTable.PivotFields("Date").AutoGroup
    ActiveSheet.ChartObjects("Chart 4").Activate
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    ActiveSheet.ChartObjects("TotalTime").Activate
    ActiveChart.FullSeriesCollection(2).Select
    ActiveChart.PivotLayout.PivotTable.PivotFields("Months").Orientation = xlHidden
    ActiveChart.PivotLayout.PivotTable.PivotFields("Date").AutoGroup
    ActiveWindow.SmallScroll Down:=9
End Sub
