Option Explicit
Sub Slide_Go_To()                                 ' checked 1/17/20
    'Go to a slide by number
    Dim slide_num As Integer
    Dim total_slides As Integer
    total_slides = ActivePresentation.Slides.count
    slide_num = InputBox("Enter slide number between 1 And " & total_slides, "Go To Slide")
    If ((slide_num <= 0) Or (slide_num > total_slides)) Then
        Slide_Go_To
    ElseIf (slide_num <= total_slides) Then
        ActiveWindow.View.goToSlide slide_num
    End If
End Sub