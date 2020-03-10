Option Explicit
Sub Text_Go_To_Small_Text()                       ' checked 2/25/20
    'Go to the next slide that has text smaller than specified
    Dim sld                                       As Slide
    Dim shp                                       As Shape
    Dim currentSlideNumber                        As Integer
    Dim startingSlideNumber                       As Integer
    Dim fontSize As Integer
    Dim currentPresentation As Presentation: Set currentPresentation = ActivePresentation
    fontSize = currentPresentation.BuiltInDocumentProperties(21)
    fontSize = InputBox("Input font size to find text smaller than", "Font Size Smaller Than", fontSize)
    currentPresentation.BuiltInDocumentProperties(21) = fontSize
    startingSlideNumber = Application.ActiveWindow.View.Slide.SlideNumber
    For currentSlideNumber = startingSlideNumber To ActivePresentation.Slides.count ' iterate slides
        Set sld = Application.ActivePresentation.Slides(currentSlideNumber)
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.TextRange.Font.Size < fontSize And shp.TextFrame.HasText = msoTrue Then
                    ActiveWindow.View.goToSlide sld.SlideIndex
                    shp.Select
                    If MsgBox("Small text found On slide " & currentSlideNumber & " Shape: " & shp.Name & ". Fix and set to " & fontSize & "?", (vbYesNo + vbQuestion), "Set to size " & fontSize & "?") = vbYes Then
                        If shp.TextFrame.AutoSize <> ppAutoSizeShapeToFitText Or MsgBox("Shape does not autosize, do you want shape to auto scale?", (vbYesNo + vbQuestion), "Auto Size Shape?") = vbYes Then
                            If MsgBox("Word Wrap On?", (vbYesNo + vbQuestion), "Word Wrap?") = vbYes Then
                                shp.TextFrame.WordWrap = msoTrue
                            End If
                            If shp.LockAspectRatio <> msoTrue Or MsgBox("Keep aspect ratio when scaling?", (vbYesNo + vbQuestion), "Aspect Ratio?") = vbYes Then
                                shp.LockAspectRatio = msoTrue
                            End If
                            shp.TextFrame.AutoSize = ppAutoSizeShapeToFitText
                        End If
                        shp.TextFrame.TextRange.Font.Size = fontSize
                    End If
                    Exit Sub
                End If
            End If
        Next shp                                  ' end of iterate shapes
    Next                                          ' end of iterate slides
    MsgBox "No More Videos Found. Move To slide 1 To search again."
End Sub