Option Explicit
Sub Image_Go_To_Next()                            ' checked 1/17/20
    'Go to the next slide that has an image
    Dim sld                                       As Slide
    Dim shp                                       As Shape
    Dim currentSlideNumber                        As Integer
    Dim startingSlideNumber                       As Integer
    Dim boolImageFound As Boolean
    boolImageFound = False
    
    
    startingSlideNumber = Application.ActiveWindow.View.Slide.SlideNumber + 1
    If startingSlideNumber > ActivePresentation.Slides.count Then
        startingSlideNumber = 1
    End If
    
    For currentSlideNumber = startingSlideNumber To ActivePresentation.Slides.count ' iterate slides
        Set sld = Application.ActivePresentation.Slides(currentSlideNumber)
        For Each shp In sld.Shapes
            If shp.Type = msoPicture Then
                boolImageFound = True
            End If
            
            If shp.Type = msoPlaceholder Then
                If shp.PlaceholderFormat.ContainedType = msoPicture Then
                    boolImageFound = True
                End If
                
            End If
            
            If boolImageFound = True Then
                boolImageFound = False
                ActiveWindow.View.goToSlide sld.SlideIndex
                shp.Select
                MsgBox "Slide Number: " & currentSlideNumber & vbCrLf & "Shape Name: " & shp.Name, vbOKOnly, "Image Found" ' outputs slide number and shape name
                If Not CommandBars.GetPressedMso("SelectionPane") Then CommandBars.ExecuteMso ("SelectionPane") ' show selection pane
                Exit Sub
            End If
            
        Next shp                                  ' end of iterate shapes
    Next                                          ' end of iterate slides
    MsgBox "No More Images Found."
End Sub