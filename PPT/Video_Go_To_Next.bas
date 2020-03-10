Option Explicit
Sub Video_Go_To_Next()                            ' checked 1/17/20
    'Go to the next slide that has a video

    Dim sld                                       As Slide
    Dim shp                                       As Shape
    Dim currentSlideNumber                        As Integer
    Dim startingSlideNumber                       As Integer
    startingSlideNumber = Application.ActiveWindow.View.Slide.SlideNumber
    For currentSlideNumber = startingSlideNumber To ActivePresentation.Slides.count ' iterate slides
        Set sld = Application.ActivePresentation.Slides(currentSlideNumber)
        For Each shp In sld.Shapes
            If shp.Type = msoMedia Then
                If shp.MediaType = ppMediaTypeMovie Then
                    ActiveWindow.View.goToSlide sld.SlideIndex
                    MsgBox "Video found On slide " & currentSlideNumber & " Shape: " & shp.Name & ". Move To the Next slide And run again To continue searching" ' outputs slide number and shape name
                    If Not CommandBars.GetPressedMso("SelectionPane") Then CommandBars.ExecuteMso ("SelectionPane") ' show selection pane
                    Exit Sub                      ' end program for user to do things
                End If
            End If                                ' end of if type is msomedia
        Next shp                                  ' end of iterate shapes
    Next                                          ' end of iterate slides
    MsgBox "No More Videos Found. Move To slide 1 To search again."
End Sub