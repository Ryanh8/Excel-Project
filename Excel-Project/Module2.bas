Attribute VB_Name = "Module2"

'code for chart on the output sheet
Sub MakeAChart()
 Dim chtChart As Chart
   'Create a new chart.
   Set chtChart = Charts.Add
   Set chtChart = chtChart.Location(Where:=xlLocationAsObject, Name:="Output")
   With chtChart
      .ChartType = xlColumnClustered
      'Link to the source data range.
      .SetSourceData Source:=Sheets("Charts").Range("L15:M16"), _
         PlotBy:=xlRows
         'code that adds titles and formats
      .HasTitle = True
      'set axes and title values
      .ChartTitle.Text = "Task Status"
      .Axes(xlCategory, xlPrimary).HasTitle = True
      .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Status"
      .Axes(xlValue, xlPrimary).HasTitle = True
      .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "number of tasks"
   End With
End Sub
Sub ClearDateRange()
Attribute ClearDateRange.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ClearDateRange Macro
'
    'clear all button that clears all row input up to the 1001th row
    Range("A2:J1001").Select
    Selection.ClearContents
    Range("B6").Select
End Sub

'this corresponds to a button on the output spreadsheet that will sort tasks by earliest to latest date
Sub sortDate()
' sort by date
    Sheets("Output").Select
    'sort C column by ascending order (earliest date to latest date)
    Range("A1").CurrentRegion.Sort Key1:=Range("C1"), order1:=xlAscending, Header:=xlYes
End Sub

Sub sortDate2()
' sort by date
    Sheets("Date Range").Select
    'sort C column by ascending order (earliest date to latest date)
    Range("A1").CurrentRegion.Sort Key1:=Range("C1"), order1:=xlAscending, Header:=xlYes
End Sub
'this corresponds to a button on the output spreadsheet that will sort tasks by highest to lowest priority calculation
Sub sortPriority()
' sort by date
    Sheets("Output").Select
    'sort H column by descending order (highest priority ranking to lowest priority ranking
    Range("A1").CurrentRegion.Sort Key1:=Range("H1"), order1:=xlDescending, Header:=xlYes
End Sub
