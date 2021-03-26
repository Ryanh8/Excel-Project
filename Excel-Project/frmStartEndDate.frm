VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStartEndDate 
   Caption         =   "UserForm1"
   ClientHeight    =   2460
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   6360
   OleObjectBlob   =   "frmStartEndDate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStartEndDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TextBox4_Change()

End Sub

Private Sub txtstartday_Change()
    
End Sub

Private Sub txtstartyear_Change()

End Sub

'code that runs to update the start and end date (not the same as due date) cells while calculating the number of days between them
'transferred values eventually appear on the Gantt Chart
Private Sub UpdateStartEndDate_Click()
Set WB = ThisWorkbook
    Set Ws = WB.Worksheets("Schedual")
    Dim StartDate As Date
    Dim Enddate As Date
    Dim updateRow As Date
    Dim daydiff As Integer
    
    'if date values are not valid or non numeric, create an error message and end the execution until it is fixed.
    If ((txtstartmonth.Value > 12) Or (txtstartday.Value > 31) Or (IsNumeric(txtstartyear.Value) = False) Or (IsNumeric(txtstartmonth.Value) = False) Or (IsNumeric(txtstartday.Value) = False)) Then
        MsgBox "Invalid Start Date Values, Please Try Again."
        Exit Sub
    End If
    If ((txtendmonth.Value > 12) Or (txtendday.Value > 31) Or (IsNumeric(txtendyear.Value) = False) Or (IsNumeric(txtendmonth.Value) = False) Or (IsNumeric(txtendday.Value) = False)) Then
        MsgBox "Invalid End Date Values, Please Try Again."
        Exit Sub
    End If
    If (txtstartyear.Value > 100) Or (txtendyear.Value > 100) Then
        MsgBox "Please format the Year as the last 2 digits, Please Try Again." & vbCrLf & "i.e. 2020 should have 20 in the input box."
        Exit Sub
    End If
    
    'set date values by converting string date variables to date variables
    StartDate = CDate(txtstartyear.Value + "/" + txtstartmonth.Value + "/" + txtstartday.Value) 'start date of task
    Enddate = CDate(txtendyear.Value + "/" + txtendmonth.Value + "/" + txtendday.Value) 'End date of task (not the same as due date)
    'based on what cell/row the mouse is on before they open the update start/end date button, that row will be the row edited with these values
    updateRow = ActiveCell.Row
    'calculates the diffrencebetween the dates
    daydiff = DateDiff("d", StartDate, Enddate)
    
    'the row determined by the active cell updates with the start and end date values of each task
    Ws.Cells(updateRow, "B") = StartDate
    Ws.Cells(updateRow, "c") = daydiff
    Ws.Cells(updateRow, "D") = Enddate


End Sub
