VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmclearrow 
   Caption         =   "Clear Row"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   4560
   OleObjectBlob   =   "frmclearrow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmclearrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub ClearOk_Click()
    Set WB = ThisWorkbook
    Set Ws = WB.Worksheets("output")
    Dim deleteRow As String
    'create an error message for invalid rows (trying to delete header or entering non number value)
    If ((RowInput.Value < 2) Or (IsNumeric(RowInput.Value) = False)) Then
        MsgBox "You have inputted an invalid value."
        Exit Sub
    End If
    
    deleteRow = RowInput.Value
    
    'manually clear all contents for a row within columns A to J
    Cells(deleteRow, "a").ClearContents
    Cells(deleteRow, "b").ClearContents
    Cells(deleteRow, "c").ClearContents
    Cells(deleteRow, "d").ClearContents
    Cells(deleteRow, "e").ClearContents
    Cells(deleteRow, "f").ClearContents
    Cells(deleteRow, "g").ClearContents
    Cells(deleteRow, "h").ClearContents
    Cells(deleteRow, "i").ClearContents
    Cells(deleteRow, "j").ClearContents
   

End Sub

Private Sub ClearRowLabel_Click()

End Sub

Private Sub RowInput_Change()

End Sub

Private Sub RowNumberLabel_Click()

End Sub
