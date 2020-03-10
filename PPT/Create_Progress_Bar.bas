Option Explicit
Sub Create_Progress_Bar()
    Dim intSlideNumber As Integer
    Dim s      As Shape
    Dim intLineHeight As Integer
    intLineHeight = 3
    Dim oPres  As Presentation
    Set oPres = ActivePresentation
    On Error Resume Next
    With oPres
        For intSlideNumber = 2 To .Slides.count
            .Slides(intSlideNumber).Shapes("Progress_Bar").Delete
            Set s = .Slides(intSlideNumber).Shapes.AddLine(Beginx:=0, BeginY:=.PageSetup.SlideHeight - (intLineHeight / 2), _
                Endx:=intSlideNumber * .PageSetup.SlideWidth / .Slides.count, EndY:=.PageSetup.SlideHeight - (intLineHeight / 2))
            s.line.Weight = intLineHeight
            s.line.BackColor.ObjectThemeColor = msoThemeColorAccent1
            s.line.ForeColor.ObjectThemeColor = msoThemeColorAccent1
            s.Name = "Progress_Bar"
        Next intSlideNumber:
    End With
    MsgBox "All Done. Created a progress bar on " & oPres.Slides.count - 1 & " slides. Theme color accent 1 was used."
End Sub