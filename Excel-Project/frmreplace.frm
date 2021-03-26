VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmreplace 
   Caption         =   "Edit Cell "
   ClientHeight    =   2025
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   6800
   OleObjectBlob   =   "frmreplace.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmreplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Label1_Click()

End Sub

Private Sub replaceData_Change()

End Sub
'when ok button clicked for editing cell values, replace value in indicated row and column with the replaced data value.
'the only way for the already populated sheet values to be edited is through this button.
Private Sub rplaceokbtn_Click()
Set WB = ThisWorkbook
Set Ws = WB.Worksheets("output")
'if user tries to delete the first row (header) or inputs a non-number value, an error message occurs
If ((TextBoxRow.Value < 2) Or (IsNumeric(TextBoxRow.Value) = False)) Then
        MsgBox "You have inputted an invalid value."
        Exit Sub
End If
'if column number is numeric, error message
If (IsNumeric(TextBoxcolumn.Value) = True) Then
        MsgBox "Please input a letter for the column. i.e. J"
        Exit Sub
End If
'if column number is the computer generated priority calculation, an error message occurs.
If (TextBoxcolumn.Value = "H") Then
        MsgBox "You cannot change values in this column."
        Exit Sub
End If
'if the column chosen is J and the replacement value is not yes or no, an error message occurs.
If (TextBoxcolumn.Value = "J") Then
        'MsgBox (replaceData.Value)
        If Not ((replaceData.Value = "no") Or (replaceData.Value = "yes")) Then
            MsgBox "For this column, enter yes or no."
        Exit Sub
        End If
End If
'if the column chosen is E and the replacement value is not within the priority calculation range (1,2,3), an error message occurs.
If (TextBoxcolumn.Value = "E") Then
        'MsgBox (replaceData.Value)
        If Not ((replaceData.Value = 1) Or (replaceData.Value = 2) Or (replaceData.Value = 2)) Then
            MsgBox "For this column, enter 1, 2, or 3"
        Exit Sub
        End If
End If
'if the column chosen is F and the replacement value is not within the priority calculation range (1,2,3), an error message occurs.
If (TextBoxcolumn.Value = "F") Then
        'MsgBox (replaceData.Value)
        If Not ((replaceData.Value = 1) Or (replaceData.Value = 2) Or (replaceData.Value = 2)) Then
            MsgBox "For this column, enter 1, 2, or 3"
        Exit Sub
        End If
End If
'if the column chosen is G and the replacement value is not within the priority calculation range (1,2,3), an error message occurs.
If (TextBoxcolumn.Value = "G") Then
        'MsgBox (replaceData.Value)
        If Not ((replaceData.Value = 1) Or (replaceData.Value = 2) Or (replaceData.Value = 2)) Then
            MsgBox "For this column, enter 1, 2, or 3"
        Exit Sub
        End If
End If
'if the column chosen is B and the replacement value is not within the category/broad stroke range (Planning, Finding, or Implementation/Testing), an error message occurs.
If (TextBoxcolumn.Value = "B") Then
        'MsgBox (replaceData.Value)
        If Not ((replaceData.Value = "Planning") Or (replaceData.Value = "Finding") Or (replaceData.Value = "Implementation/Testing")) Then
            MsgBox "For this column, enter Planning, Finding, or Implementation/Testing"
        Exit Sub
        End If
End If
'if the column chosen is C and the replacement value is not a date value, an error message occurs.
If (TextBoxcolumn.Value = "C") Then
        'MsgBox (replaceData.Value)
        If (IsDate(replaceData.Value) = False) Then
            MsgBox "For this column, enter a valid date (YYYY-MM-DD)"
        Exit Sub
        End If
End If
'replace value in the specified cell with this data in the box.
Ws.Cells(TextBoxRow.Value, TextBoxcolumn.Value) = replaceData.Value

End Sub

Private Sub TextBoxRow_Change()

End Sub

Private Sub UserForm_Click()

End Sub
