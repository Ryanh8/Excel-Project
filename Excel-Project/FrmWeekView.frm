VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmWeekView 
   Caption         =   "Tasks in a Week"
   ClientHeight    =   3450
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   8280.001
   OleObjectBlob   =   "FrmWeekView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmWeekView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ClearOk_Click()
    'find dates - idea is to get int values from worksheet and dates within that range will open in new page.
    Dim startDateUp As Date
    Dim StartDate As String
    Dim startyr As String
    Dim startmnth As String
    Dim startday As String

    'if date values are not valid or not numeric, error messages occur.
    If ((IsNumeric(startdayin.Value) = False) Or (IsNumeric(startmnthin.Value) = False) Or (IsNumeric(startyrin.Value) = False)) Then
        MsgBox "Please enter numerical values for the start date. Try Again."
        Exit Sub
    End If
    If (startyrin.Value > 99) Then
        MsgBox "Please format the Year as the last 2 digits, Please Try Again." & vbCrLf & "i.e. 2020 should have 20 in the input box."
        Exit Sub
    End If
    If ((startdayin.Value > 31) Or (startmnthin.Value > 12)) Then
        MsgBox "Please enter valid day and month values. Try Again."
        Exit Sub
    End If
    
    'get dates from userform and assign to variables
    startday = startdayin.Value
    startmnth = startmnthin.Value
    startyr = startyrin.Value
    
    'make date into a string
    StartDate = (startyr + "/" + startmnth + "/" + startday)
    'convert string date to date variable
    startDateUp = CDate(StartDate)
    
    'declare variables for each day for up to a week after the start date
    Dim StartAdd1 As Date 'startdate + 1 day
    Dim StartAdd2 As Date 'startdate + 2 days
    Dim StartAdd3 As Date 'startdate + 3 days
    Dim StartAdd4 As Date
    Dim StartAdd5 As Date
    Dim StartAdd6 As Date
    'assign values to each variable by adding 1-6 days to the start date for a week range
    StartAdd1 = DateAdd("d", 1, startDateUp)
    StartAdd2 = DateAdd("d", 2, startDateUp)
    StartAdd3 = DateAdd("d", 3, startDateUp)
    StartAdd4 = DateAdd("d", 4, startDateUp)
    StartAdd5 = DateAdd("d", 5, startDateUp)
    StartAdd6 = DateAdd("d", 6, startDateUp)
    
 'check number of filled rows
    a = Worksheets("Output").Cells(Rows.Count, 1).End(xlUp).Row
    Dim copyrange As String
    ' for second row to last row
    For i = 2 To a
    ' if row meets condition, which is start date, or any date that is between 1 to 6 days from start (matching the startadd variable values)
    If ((Worksheets("Output").Cells(i, 3).Value = startDateUp) Or (Worksheets("Output").Cells(i, 3).Value = StartAdd1) Or (Worksheets("Output").Cells(i, 3).Value = StartAdd2) Or (Worksheets("Output").Cells(i, 3).Value = StartAdd3) Or (Worksheets("Output").Cells(i, 3).Value = StartAdd4) Or (Worksheets("Output").Cells(i, 3).Value = StartAdd5) Or (Worksheets("Output").Cells(i, 3).Value = StartAdd6)) Then
        Let copyrange = "A" & i & ":" & "J" & i
        'row copied
        Worksheets("Output").Range(copyrange).Copy
        Worksheets("Date Range").Activate
        ' find number of filled roles for date range spreadsheet
        b = Worksheets("Date Range").Cells(Rows.Count, 1).End(xlUp).Row
        'select next blank cell
        Worksheets("Date Range").Cells(b + 1, 1).Select
        Sheets("Date Range").Paste
        'paste row
        Worksheets("Output").Activate
    End If
        
    Next
Application.CutCopyMode = False
ThisWorkbook.Worksheets("Output").Cells(1, 1).Select

'transfer dates to the date range spreadsheet from the output spreadsheet
Set WB = ThisWorkbook
Set Ws = WB.Worksheets("Date Range")
Ws.Cells(1, "M") = startyr
Ws.Cells(1, "O") = startmnth
Ws.Cells(1, "Q") = startday
Ws.Cells(1, "K") = "Start"
End Sub

Private Sub endday_Change()
    
End Sub

Private Sub ClearRowLabel_Click()

End Sub

Private Sub EndDateLabel_Click()

End Sub

Private Sub enddayin_Change()

End Sub

Private Sub StartDateLabel_Click()

End Sub

Private Sub startdayin_Change()

End Sub
