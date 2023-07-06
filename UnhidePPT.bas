Attribute VB_Name = "PrepToolKit1"
Option Explicit

Sub UnhideSlide()
Dim sl As Slide
For Each sl In ActivePresentation.Slides
    sl.SlideShowTransition.Hidden = msoFalse
Next sl
End Sub

Sub UnhideShape()
Dim sh As Shape
Dim sl As Slide

For Each sl In ActivePresentation.Slides
    If sl.SlideShowTransition.Hidden = msoFalse then
        For Each sh In sl.Shapes
            sh.Visible = msoTrue
        Next sh
    End If
Next sl
End Sub
