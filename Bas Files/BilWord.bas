Attribute VB_Name = "PrepToolKit1"
Option Explicit

Sub Bil_Tables()
Dim tp As Long
Dim Tb As Table
Dim c As Cell

On Error Resume Next
Application.ScreenUpdating = False
Application.DisplayAlerts = wdAlertsNone

For Each Tb in ActiveDocument.Tables
	For Each c In Tb.Range.Cells
		For tp = 1 To c.Range.Paragraphs.Count
			c.Range.Paragraphs.Add
			c.Range.Paragraphs(c.Range.Paragraphs.Count).Range.FormattedText = c.Range.Paragraphs(tp).Range.FormattedText
			c.Range.Paragraphs(tp).Range.Font.Hidden = True
		Next tp
		c.Range.Paragraphs(tp).Range.Characters(InStrRev(c.Range.Paragraphs(tp).Range, Chr(13))).Delete
	Next c
Next Tb

ActiveDocument.SaveAs2 FileName:=ActiveDocument.Path & "/Temp_" & ActiveDocument.Name
ActiveDocument.Close (wdDoNotSaveChanges)

End Sub

Sub Bil_Text()
Dim p As Paragraph
Dim Tb As Table

On Error Resume Next
Application.ScreenUpdating = False

For Each Tb In ActiveDocument.Tables
    Tb.Rows.WrapAroundText = True
Next Tb

For Each p In ActiveDocument.Paragraphs
    If p.Range.Information(wdWithInTable) = False Then
        p.Range.Find.Execute FindText:="^t", ReplaceWith:=" ", Replace:=wdReplaceAll
        p.Range.Paragraphs.Add p.Range
        p.Previous.Range.FormattedText = p.Range.FormattedText
        p.Previous.Range.Font.Hidden = True
        ActiveDocument.Range(p.Previous.Range.start, p.Range.End).ConvertToTable Separator:=wdSeparateByParagraphs, NumRows:=1, NumColumns:=2
	end if
next p

ActiveDocument.Content.Find.Execute FindText:="^p^p", ReplaceWith:="^p", Replace:=wdReplaceAll

For Each Tb In ActiveDocument.Tables
    Tb.Select
    Selection.SplitTable
Next Tb

For Each Tb In ActiveDocument.Tables
    Tb.Rows.WrapAroundText = False
Next Tb

ActiveDocument.SaveAs2 FileName:=ActiveDocument.Path & "/Bil_" & Right(ActiveDocument.Name, Len(ActiveDocument.Name) -5)
ActiveDocument.Close (wdDoNotSaveChanges)
End Sub