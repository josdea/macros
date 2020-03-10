Option Explicit
Sub Shapes_Go_To_Next_Non_Placeholder()
    'Go to the next slide which has a non placeholder shape
    Dim sld                                       As Slide
    Dim shp                                       As Shape
    Dim currentSlideNumber                        As Integer
    Dim startingSlideNumber                       As Integer
    startingSlideNumber = Application.ActiveWindow.View.Slide.SlideNumber + 1
    For currentSlideNumber = startingSlideNumber To ActivePresentation.Slides.count ' iterate slides
        Set sld = Application.ActivePresentation.Slides(currentSlideNumber)
        
        For Each shp In sld.Shapes
            If shp.Type <> msoPlaceholder Then
                ActiveWindow.View.goToSlide sld.SlideIndex
                
                If currentSlideNumber = Application.ActiveWindow.View.Slide.SlideNumber Then
                    shp.Select
                End If
                
                MsgBox "Non-Placeholder found On slide " & currentSlideNumber & " Shape: " & shp.Name & ". Move To the Next slide And run again To continue searching" ' outputs slide number and shape name
                If Not CommandBars.GetPressedMso("SelectionPane") Then CommandBars.ExecuteMso ("SelectionPane") ' show selection pane
                Exit Sub                          ' end program for user to do things
            End If                                ' end of if type is msomedia
        Next shp                                  ' end of iterate shapes
    Next                                          ' end of iterate slides
    MsgBox "No More Images Found. Move To slide 1 To search again."
End Sub