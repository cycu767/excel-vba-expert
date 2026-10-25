Attribute VB_Name = "MCharts"
Option Explicit
' Creates/updates ONE named embedded chart. Use only on an authorized working copy.
Public Sub UpdateSummaryChart(ByVal sheet As Excel.Worksheet, _
                              ByVal source As Excel.Range, _
                              ByVal chartName As String)
    Dim item As Excel.ChartObject
    Dim chartObject As Excel.ChartObject
    If sheet Is Nothing Then Err.Raise 5, "UpdateSummaryChart", "Sheet required."
    If source Is Nothing Then Err.Raise 5, "UpdateSummaryChart", "Source required."
    If Not (source.Parent Is sheet) Then _
        Err.Raise 5, "UpdateSummaryChart", "This example requires data on the target sheet."
    If source.Areas.Count <> 1 Or source.Rows.Count < 2 Or source.Columns.Count <> 2 Then _
        Err.Raise 5, "UpdateSummaryChart", "Provide header plus data in two columns."
    If Len(Trim$(chartName)) = 0 Then Err.Raise 5, "UpdateSummaryChart", "Chart name required."
    For Each item In sheet.ChartObjects
        If StrComp(item.Name, chartName, vbBinaryCompare) = 0 Then
            Set chartObject = item
            Exit For
        End If
    Next item
    If chartObject Is Nothing Then
        Set chartObject = sheet.ChartObjects.Add(Left:=300, Top:=20, Width:=420, Height:=260)
        chartObject.Name = chartName
    End If
    With chartObject.Chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=source, PlotBy:=xlColumns
        .HasTitle = True
        .ChartTitle.Text = "Summary"
        .HasLegend = False
    End With
    ' A failure may leave partial formatting/new chart: caller must inspect the copy.
End Sub
