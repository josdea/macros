Option Explicit

Sub goToNextVideo()
    Dim sld                                       As Slide
    Dim shp                                       As Shape
    Dim currentSlideNumber                        As Integer
    Dim startingSlideNumber                       As Integer
    
    startingSlideNumber = Application.ActiveWindow.View.Slide.SlideNumber
    
    For currentSlideNumber = startingSlideNumber To ActivePresentation.Slides.Count        ' iterate slides
        Set sld = Application.ActivePresentation.Slides(currentSlideNumber)
       
        Debug.Print "Slide " & sld.SlideNumber & " of " & ActivePresentation.Slides.Count
        
        For Each shp In sld.Shapes
            If shp.Type = msoMedia Then
                If shp.MediaType = ppMediaTypeMovie Then
                 sld.Select        'select current slide
                    MsgBox "Video found on slide " & currentSlideNumber & " Shape: " & shp.name & ". Move to the next slide and run again to continue searching"       ' outputs slide number and shape name
                    If Not CommandBars.GetPressedMso("SelectionPane") Then CommandBars.ExecuteMso ("SelectionPane")        ' show selection pane
                    Exit Sub        ' end program for user to do things
                End If
            End If        ' end of if type is msomedia
        Next shp        ' end of iterate shapes
    Next        ' end of iterate slides
    
    MsgBox "No More Videos Found. Move to slide 1 to search again."
    
End Sub
