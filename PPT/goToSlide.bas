Option Explicit
' TODO check if user cancels dialog window
TODO
Sub goToSlide() 
    Dim slide_num As Integer
    Dim total_slides As Integer
    total_slides = ActivePresentation.Slides.Count
    slide_num = InputBox("Enter slide number between 1 and " & total_slides, "Go To Slide")
    If ((slide_num <= 0) Or (slide_num > total_slides)) Then
        go_to_slide
    ElseIf (slide_num <= total_slides) Then
        'MsgBox ("Jumping to slide #" & slide_num)
        ActiveWindow.View.goToSlide slide_num
    End If
End Sub
