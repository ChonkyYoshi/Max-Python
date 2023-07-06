Attribute VB_Name = "PrepToolKit1"
Option Explicit

Sub UnhideSheet()
Dim sh As Worksheet

For Each sh In ActiveWorkbook.Worksheets
    sh.Visible = xlSheetVisible
Next sh
End Sub

Sub UnhideCol()
Dim sh As Worksheet

For Each sh In ActiveWorkbook.Worksheets
    If sh.Visible = xlSheetVisible Then
        sh.Columns.EntireColumn.Hidden = False
    End If
Next sh
End Sub

Sub UnhideRow()
Dim sh As Worksheet

For Each sh In ActiveWorkbook.Worksheets
    If sh.Visible = xlSheetVisible Then
        sh.Rows.EntireRow.Hidden = False
    End If
Next sh
End Sub
