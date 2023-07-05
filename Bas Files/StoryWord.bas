Attribute VB_Name = "PrepToolKit1"
Option Explicit

Sub Story()
Dim Tb As Table

On Error Resume Next
ActiveDocument.Range.Font.Hidden = True
For Each Tb In ActiveDocument.Tables
    If Left(Tb.Cell(1, 4).Range.Text, 11) = "Translation" Then
        Tb.Cell(1, 4).Column.Select
        Selection.Font.Hidden = False
    With Selection.Find
        .Text = "%<?@>%"
        .Format = True
        .Replacement.Font.Hidden = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    End If
Next Tb

ActiveDocument.SaveAs2 FileName:=ActiveDocument.Path & "/Prep_" & Right(ActiveDocument.Name, Len(ActiveDocument.Name) - 5)
ActiveDocument.Close (wdDoNotSaveChanges)
End Sub
