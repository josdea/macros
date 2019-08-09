Option Explicit

Sub createNoteTextBoxesAllSlides()
    
    ' TODO run loop and watch for points sizes and dimension of first slide shapes, add resize functions for the other obejcts
    
    Debug.Print "Start of Create all TextBoxes on notes slides";
    ActiveWindow.ViewType = ppViewNotesPage
    Dim sld                                       As Slide        ' declare slide object
    
    For Each sld In ActivePresentation.Slides        ' iterate slides
        sld.Select
        Debug.Print "Slide " & sld.SlideNumber & " of " & ActivePresentation.Slides.Count
        
        iterateNotesPageShapes sld        ' Call to iterate shapes on slide
        
    Next sld        ' end of iterate slides
    
    Debug.Print "All done creating notes textboxes";
    If Not CommandBars.GetPressedMso("SelectionPane") Then CommandBars.ExecuteMso ("SelectionPane")
    MsgBox "All done creating textboxes on slides"
    
End Sub

Sub iterateNotesPageShapes(sld As Slide)
    
    Dim shp                                       As Shape        ' declare shape object
    Dim hasModuleTitle                            As Boolean
    Dim hasLearnerNotes                           As Boolean
    Dim hasObjective                              As Boolean
    Dim hasMinutes                                As Boolean
    
    hasModuleTitle = False
    hasLearnerNotes = False
    hasObjective = False
    hasMinutes = False
    
    For Each shp In sld.NotesPage.Shapes
        
        Select Case shp.name
            Case Is = "Objective"        ' objective
                hasObjective = True
                
                updateShapePosition shp, 5.5, 9.4        ' reposition the shape
                updateShapeSize shp, 2, 0.3        'resize the shape
                shp.TextFrame2.AutoSize = msoAutoSizeTextToFitShape        'autosize the font to fit the text box
                shp.TextFrame2.WordWrap = msoFalse        ' no wrapping
                shp.TextEffect.FontItalic = msoTrue
                shp.TextEffect.Alignment = msoTextEffectAlignmentRight
                
                
                ' updateSizePosition shp, 5.5, 0.3, 2, 0.3, "Objective", "", 11
            Case Is = "Minutes"
                hasMinutes = True
                
                updateShapePosition shp, 0, 9.4        ' reposition the shape
                updateShapeSize shp, 5.5, 0.3        'resize the shape
                shp.TextFrame2.AutoSize = msoAutoSizeTextToFitShape        'autosize the font to fit the text box
                shp.TextFrame2.WordWrap = msoFalse        ' no wrapping
                
                'updateSizePosition shp, 0, 676.8, 144, 21.6, "Minutes", "Mins", 11
            Case Is = "ModuleTitle"
                hasModuleTitle = True
                updateShapePosition shp, 1, 0        ' reposition the shape
                updateShapeSize shp, 5.5, 0.3        'resize the shape
                shp.TextFrame2.AutoSize = msoAutoSizeTextToFitShape        'autosize the font to fit the text box
                shp.TextFrame2.WordWrap = msoFalse        ' no wrapping
                shp.TextEffect.Alignment = msoTextEffectAlignmentCentered
                
                ' updateSizePosition shp, 0.75, 0, 6, 0.3, "ModuleTitle", "", 11
            Case Is = "LearnerNotes"
                hasLearnerNotes = True
                
                updateShapePosition shp, 0, 3        ' reposition the shape
                updateShapeSize shp, 4.75, 6.3        'resize the shape
                
                With shp
                    .TextFrame2.AutoSize = msoAutoSizeTextToFitShape        'autosize the font to fit the text box
                    .TextFrame2.WordWrap = msoTrue        ' no wrapping for slide number
                    .TextEffect.FontItalic = msoFalse
                    .TextEffect.FontBold = msoFalse
                    .TextEffect.fontSize = 11
                    .TextFrame.TextRange.Font.name = "+mn-lt"
                    .TextFrame.TextRange.Font.underline = False
                    
                End With
                
                'updateSizePosition shp, 0, 3, 4.75, 6.25, "LearnerNotes", "", 11
            Case Else
                
                If shp.Type = msoPlaceholder Then        ' check if its a placeholder
                Select Case shp.PlaceholderFormat.Type        ' type of placeholder
                    Case Is = ppPlaceholderFooter        ' footer shape
                        updateShapePosition shp, 0, 9.7        ' reposition the shape
                        updateShapeSize shp, 5.5, 0.3        'resize the shape
                        shp.TextFrame2.AutoSize = msoAutoSizeTextToFitShape        'autosize the font to fit the text box
                        shp.TextFrame2.WordWrap = msoFalse        ' no wrapping for footer
                        
                    Case Is = ppPlaceholderSlideNumber ' Slide Number
                        updateShapePosition shp, 5.5, 9.7        ' reposition the shape
                        updateShapeSize shp, 2, 0.3        'resize the shape
                        shp.TextFrame2.AutoSize = msoAutoSizeTextToFitShape        'autosize the font to fit the text box
                        shp.TextFrame2.WordWrap = msoFalse        ' no wrapping for slide number
                        shp.TextEffect.Alignment = msoTextEffectAlignmentRight
                        
                    Case Is = ppPlaceholderBody ' Presenter notes
                        updateShapePosition shp, 4.75, 0.7        ' reposition the shape
                        updateShapeSize shp, 2.75, 8.6        'resize the shape
                        With shp
                            .TextFrame2.AutoSize = msoAutoSizeTextToFitShape        'autosize the font to fit the text box
                            .TextFrame2.WordWrap = msoTrue        ' no wrapping for slide number
                            .TextEffect.FontItalic = msoFalse
                            .TextEffect.FontBold = msoFalse
                            .TextEffect.fontSize = 11
                            .TextFrame.TextRange.Font.name = "+mn-lt"
                            .TextFrame.TextRange.Font.underline = False
                            .Fill.ForeColor.RGB = RGB(255, 255, 255)
                        End With
                        
                    Case Is = ppPlaceholderTitle ' Slide Image Placeholder
                        
                        updateShapePosition shp, 0.5, 0.7        ' reposition the shape
                        updateShapeSize shp, 4, 2.25        'resize the shape
                        
                    Case Else
                        
                End Select
                
            End If
            
            ' Debug.Print "NOTE: should not reach this point unless there is another text placeholder on the slide TODO"
    End Select        ' end of case
    
    '    End If        ' end of textframe has text
    
Next shp

If hasObjective = False Then
    Set shp = sld.NotesPage.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=0, Top:=0, Width:=1, Height:=1)
    shp.name = "Objective"
    updateShapePosition shp, 5.5, 9.4        ' reposition the shape
    updateShapeSize shp, 2, 0.3        'resize the shape
    shp.TextFrame2.AutoSize = msoAutoSizeTextToFitShape        'autosize the font to fit the text box
    shp.TextFrame2.WordWrap = msoFalse        ' no wrapping for slide number
    shp.TextEffect.FontItalic = msoTrue
    shp.TextEffect.Alignment = msoTextEffectAlignmentRight
    'createSlideShapeTextbox sld, 5.5 * 72, 0.3, 2, 0.3, "Objective", "Object", 11
End If
If hasMinutes = False Then
Set shp = sld.NotesPage.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=0, Top:=0, Width:=1, Height:=1)
    shp.name = "Minutes"
                    updateShapePosition shp, 0, 9.4        ' reposition the shape
                updateShapeSize shp, 5.5, 0.3        'resize the shape
                shp.TextFrame2.AutoSize = msoAutoSizeTextToFitShape        'autosize the font to fit the text box
                shp.TextFrame2.WordWrap = msoFalse        ' no wrapping
    
    'createSlideShapeTextbox sld, 0, 676.8, 144, 21.6, "Minutes", "Mins", 11
End If

If hasModuleTitle = False Then
Set shp = sld.NotesPage.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=0, Top:=0, Width:=1, Height:=1)
    shp.name = "ModuleTitle"
                    updateShapePosition shp, 1, 0        ' reposition the shape
                updateShapeSize shp, 5.5, 0.3        'resize the shape
                shp.TextFrame2.AutoSize = msoAutoSizeTextToFitShape        'autosize the font to fit the text box
                shp.TextFrame2.WordWrap = msoFalse        ' no wrapping
                shp.TextEffect.Alignment = msoTextEffectAlignmentCentered
    '    createSlideShapeTextbox sld, 0.75, 0, 6, 0.3, "ModuleTitle", "Mod title", 11
End If
If hasLearnerNotes = False Then
Set shp = sld.NotesPage.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=0, Top:=0, Width:=1, Height:=1)
    shp.name = "LearnerNotes"
                    updateShapePosition shp, 0, 3        ' reposition the shape
                updateShapeSize shp, 4.75, 6.3        'resize the shape
                
                With shp
                    .TextFrame2.AutoSize = msoAutoSizeTextToFitShape        'autosize the font to fit the text box
                    .TextFrame2.WordWrap = msoTrue        ' no wrapping for slide number
                    .TextEffect.FontItalic = msoFalse
                    .TextEffect.FontBold = msoFalse
                    .TextEffect.fontSize = 11
                    .TextFrame.TextRange.Font.name = "+mn-lt"
                    .TextFrame.TextRange.Font.underline = False
                    
                End With
    '     createSlideShapeTextbox sld, 0, 3, 4.75, 6.25, "LearnerNotes", "Learner notes", 11
End If

End Sub

Sub createSlideShapeTextbox(sld As Slide, shpLeftPosition As Long, shpTopPosition As Long, shpWidth As Long, shpHeight As Long, shpName As String, shpText As String, shpFontSize As Integer)
    
    Dim shp                                       As Shape
    Set shp = sld.NotesPage.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=shpLeftPosition, Top:=shpTopPosition, Width:=shpWidth, Height:=shpWidth)
    With shp
        .TextFrame.TextRange.Text = shpText
        .name = shpName
        '.TextFrame2.AutoSize = msoAutoSizeTextToFitShape
        .Height = shpHeight
        .Width = shpWidth
        .TextFrame.TextRange.Font.size = shpFontSize
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
    
End Sub

Sub updateShapeSize(shp As Shape, shpWidth As Double, shpHeight As Double)
    
    With shp
        .Width = shpWidth * 72
        .Height = shpHeight * 72
    End With
    
End Sub

Sub updateShapePosition(shp As Shape, shpLeftPosition As Double, shpTopPosition As Double)
    With shp
        .Left = shpLeftPosition * 72
        .Top = shpTopPosition * 72
    End With
    
End Sub

Sub createShape(sld As Slide, shpName As String)
    
    Dim shp                                       As Shape
    Set shp = sld.NotesPage.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=0, Top:=0, Width:=1 * 72, Height:=1 * 72)
    shp.name = shpName
    
End Sub


