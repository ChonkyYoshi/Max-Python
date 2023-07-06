Attribute VB_Name = "PrepToolKit1"
Option Explicit

Sub Unhide()
Dim sec As Section
Dim item As Variant
For Each sec In ActiveDocument.Sections
    sec.Range.Font.Hidden = False
    For Each item In sec.Headers
        item.Range.Font.Hidden = False
    Next item
    For Each item In sec.Footers
        item.Range.Font.Hidden = False
    Next item
    For Each item In ActiveDocument.Shapes
        item.Visible = msoTrue
        item.TextFrame.TextRange.Font.Hidden = False
    Next item
    For Each item In ActiveDocument.InlineShapes
        item.Range.Font.Hidden = False
    Next item
Next sec
End Sub
