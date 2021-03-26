VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOutput 
   Caption         =   "Task Input"
   ClientHeight    =   5640
   ClientLeft      =   40
   ClientTop       =   280
   ClientWidth     =   10560
   OleObjectBlob   =   "frmOutput.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Team Term 8: Serena Pang, Aamina Anjum, Nafis Huq, Ryan Hu, Athena Ketabi
'Date: Nov. 29, 2020
'MSCI 100 Team Term Project Final
'Description: Final Decision Support System for Brandon's FYDP Team
'Input: Form that will prompt user to enter information about a responsibility/task to complete
'Output: Simple Organizational sheets that will organize input into respective categories or user-prompted sheets

Dim intRow As Integer

Private Sub AdditionalNotesIn_Change()

End Sub

Private Sub AdditionalNotesLabel_Click()

End Sub

Private Sub CancelBtn_Click()
    Unload Me
    End
End Sub

Private Sub DueDateDay_Change()

End Sub

Private Sub DueDateMnth_Change()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub lblDate2_Click()

End Sub

Private Sub lblTitle_Click()

End Sub

Private Sub PrioritizationLabel_Click()

End Sub

Private Sub PriorityFactor1In_Change()

End Sub

Private Sub TaskNameLabel_Click()

End Sub

Private Sub tempLabel_Click()

End Sub

'pull down lists

Private Sub UserForm_Initialize()
'set values for pull down list for category of task
    CategoryDrop.AddItem "Finding"
    CategoryDrop.AddItem "Planning"
    CategoryDrop.AddItem "Implementation/Testing"
'set values for pull down list for completion statment
    ComboBoxComplete.AddItem "yes"
    ComboBoxComplete.AddItem "no"
End Sub





'set executions when ok button is clicked

Private Sub OkBtn_Click()
'send input into workbook output (master spreadsheet)
    Set WB = ThisWorkbook
    Set Ws = WB.Worksheets("Output")
'declare counter that will take input to the next line
    intRow = 1
    'for each entry, put next row of data on the next empty row
    Do While (Ws.Cells(intRow, "A") <> "")
    intRow = intRow + 1
    Loop
    'if date values are invalid (not numbers or not in the right range), create error message
    If (DueDateYr.Value > 100) Then
        MsgBox "Please format the Year as the last 2 digits, Please Try Again." & vbCrLf & "i.e. 2020 should have 20 in the input box."
        '-1 from RowCount to not skip rows
        RowCount = RowCount - 1
        Exit Sub
    End If
    If ((DueDateMnth.Value > 12) Or (DueDateDay.Value > 31) Or (IsNumeric(DueDateYr.Value) = False) Or (IsNumeric(DueDateDay.Value) = False) Or (IsNumeric(DueDateMnth.Value) = False)) Then
        MsgBox "Invalid Date Values, Please Try Again."
        '-1 from RowCount to not skip rows
        RowCount = RowCount - 1
        Exit Sub
    End If
    'if priority factor is not within range of 1 to 3 or is not numeric, create error message
    If ((IsNumeric(PriorityFactor1In.Value) = False) Or (IsNumeric(PriorityFactor2In.Value) = False) Or (IsNumeric(PriorityFactor3In.Value) = False)) Then
        MsgBox "Please enter 1, 2, or 3 into priority ranking. Try Again."
        '-1 from RowCount to not skip rows
        RowCount = RowCount - 1
        Exit Sub
    End If
    If ((PriorityFactor1In > 3) Or (PriorityFactor2In > 3) Or (PriorityFactor3In > 3) Or (PriorityFactor1In < 1) Or (PriorityFactor2In < 1) Or (PriorityFactor3In < 1)) Then
        MsgBox "Please enter 1, 2, or 3 into priority ranking. Try Again."
        '-1 from RowCount to not skip rows
        RowCount = RowCount - 1
        Exit Sub
    End If
    'if approximate time is not numerical or less than 0, create error message
    If ((IsNumeric(ApproxTimeIn.Value) = False) Or (ApproxTimeIn.Value < 0)) Then
        MsgBox "The approximate time value is not valid. Try Again."
        '-1 from RowCount to not skip rows
        RowCount = RowCount - 1
        Exit Sub
    End If
    
    'declare variables needed
    Dim finalDate As Date
    'priority variables will calculate the importance of the task and aid in task organization
    Dim prioritycount As Integer
    Dim priorityfact1 As Integer
    Dim priorityfact2 As Integer
    Dim priorityfact3 As Integer
    Dim timepriorityfact As Integer
    
    timepriorityfact = 0
    timepriorityfact = ApproxTimeIn.Value
    
    'set priority factor depending on the time to complete the task (later added to the computer generated priority calculation)
    If (timepriorityfact < 6) Then
        timepriorityfact = 1
    End If
    If ((timepriorityfact >= 6) And (timepriorityfact <= 12)) Then
        timepriorityfact = 2
    End If
    If (timepriorityfact > 12) Then
        timepriorityfact = 3
    End If
    
    'set priority rankings to 0 before assigning them values
    priorityfact1 = 0
    priorityfact2 = 0
    priorityfact3 = 0
    prioritycount = 0
    
    Set WB = ThisWorkbook
    Set Ws = WB.Worksheets("Output")
    'transfer values from input userform to specific rows and columns and assign variables values for later calculations
    Ws.Cells(intRow, "A") = TaskNameInput.Value
    Ws.Cells(intRow, "D") = PriorityFactor1In.Value
    priorityfact1 = PriorityFactor1In.Value
    Ws.Cells(intRow, "E") = PriorityFactor2In.Value
    priorityfact2 = PriorityFactor2In.Value
    Ws.Cells(intRow, "F") = PriorityFactor3In.Value
    priorityfact3 = PriorityFactor3In.Value
    Ws.Cells(intRow, "I") = AdditionalNotesIn.Value
    Ws.Cells(intRow, "G") = ApproxTimeIn.Value
    Ws.Cells(intRow, "B") = CategoryDrop.Value
    'calculate priority importance of task for later organization through previously assigned variables
    prioritycount = priorityfact1 + priorityfact2 + priorityfact3 + timepriorityfact
    Ws.Cells(intRow, "H") = prioritycount
    'create final date value from values in userform
    finalDate = CDate(DueDateYr.Value + "/" + DueDateMnth.Value + "/" + DueDateDay.Value)
    Ws.Cells(intRow, "C") = finalDate
    'add yes/no values for completion
    Ws.Cells(intRow, "J") = ComboBoxComplete.Value
    

    
End Sub



Private Sub TaskNameInput_Change()

End Sub

Private Sub UserForm_Click()

End Sub

