' Todo loop thorugh all layouts and custom layouts

Option Explicit

Sub main()
    Debug.Print "****Start of Main****"
    Debug.Print "Presentation: " & ActivePresentation.Name
    Call presentationActions(ActivePresentation)
    Call iterateSlides(ActivePresentation)
    Debug.Print "****End of Main****"
End Sub

Function iterateSlides(currentPresentation As Presentation)
    Dim sld                                       As Slide
    Dim currentSlideNumber                        As Integer
    Dim startingSlideNumber                       As Integer
    
    startingSlideNumber = 1
    
    If (currentPresentation.Application.ActiveWindow.View.Slide.SlideNumber <> 1) Then
        If (MsgBox("Run from current slide position? (Otherwise, it will run from the start)", (vbYesNo + vbQuestion), "Run?") = vbYes) Then
            startingSlideNumber = currentPresentation.Application.ActiveWindow.View.Slide.SlideNumber
        End If
    End If
    
    For currentSlideNumber = startingSlideNumber To currentPresentation.Slides.Count
        Set sld = currentPresentation.Slides(currentSlideNumber)
        '    For Each sld In currentPresentation.Slides
        Call getSectionName(currentPresentation, sld)
        Debug.Print " Slide " & sld.SlideNumber & " / " & ActivePresentation.Slides.Count
        Call slideActions(sld)
        Call iterateSlideShapes(sld)
        Call iterateNoteShapes(sld)
        Call iterateSlideComments(sld)
        ' Next sld
    Next        'Slide
End Function

Function iterateSlideShapes(sld As Slide)
    Dim shp                                       As Shape
    Dim shpCount                                  As Integer
    shpCount = 1
    For Each shp In sld.Shapes
        Debug.Print "  Slide Shape " & shpCount & " / " & sld.Shapes.Count & " Type: " & shp.Type & " (" & shp.Name & ")"
        shpCount = shpCount + 1
        Call slideShapeActions(sld, shp)
        Call slideShapeTypeSorter(sld, shp)
    Next shp
    
End Function

Function slideShapeTypeSorter(sld As Slide, shp As Shape)
    If shp.Type = msoPlaceholder Then        'shape is a placeholder
    'See here for all shape types https://docs.microsoft.com/en-us/office/vba/api/office.msoshapetype
    Call slidePlaceholderActions(sld, shp)
Else        'shape is not a placeholder
    Call slideNonPlaceholderActions(sld, shp)
End If        'End of Shape Type

End Function

Function iterateNoteShapes(sld As Slide)
    Dim shp                                       As Shape
    Dim shpCount                                  As Integer
    shpCount = 1
    
    For Each shp In sld.NotesPage.Shapes
        Debug.Print "   Note Shape " & shpCount & " / " & sld.NotesPage.Shapes.Count & " Type: " & shp.Type & " (" & shp.Name & ")"
        shpCount = shpCount + 1
        Call noteShapeActions(sld, shp)
        Call noteShapeTypeSorter(sld, shp)
    Next shp
    
End Function

Function noteShapeTypeSorter(sld As Slide, shp As Shape)
    If shp.Type = msoPlaceholder Then        'shape is a placeholder
    'See here for all shape types https://docs.microsoft.com/en-us/office/vba/api/office.msoshapetype
    Call notePlaceholderActions(sld, shp)
Else        'shape is not a placeholder
    Call noteNonPlaceholderActions(sld, shp)
End If        'End of Shape Type

End Function

Function iterateSlideComments(sld As Slide)
    Dim cmt As Comment
    Dim cmtCount                                  As Integer
    cmtCount = 1
    For Each cmt In sld.Comments
        Call slideCommentActions(sld, cmt)
    Next cmt
    
End Function

Function getSectionName(currentPresentation As Presentation, sld As Slide) As String
    If currentPresentation.SectionProperties.Count > 0 Then        'sections exist
    If (sld.SlideNumber = 1) Then        'First slide so output section info
    Debug.Print "Section " & sld.sectionIndex & " of " & currentPresentation.SectionProperties.Count & " (" & currentPresentation.SectionProperties.Name(sld.sectionIndex) & ")"
    Call sectionStartAction(currentPresentation, sld)
ElseIf (sld.sectionIndex <> currentPresentation.Slides(sld.SlideNumber - 1).sectionIndex) Then        'Not the first slide but section index is different than previous slide
Debug.Print "Section " & sld.sectionIndex & " of " & currentPresentation.SectionProperties.Count & " (" & currentPresentation.SectionProperties.Name(sld.sectionIndex) & ")"
Call sectionStartAction(currentPresentation, sld)
ElseIf (sld.SlideNumber = currentPresentation.Slides.Count) Then        'Last slide of the presentation so the last slide in a section
Call sectionEndAction(currentPresentation, sld)
ElseIf (sld.sectionIndex <> currentPresentation.Slides(sld.SlideNumber + 1).sectionIndex) Then
    Call sectionEndAction(currentPresentation, sld)
End If        'End of first slide or differing sections IF
Else        'There are no sections in current PPT
    If sld.SlideNumber = 1 Then        'Only display the following, once
    Debug.Print "No Sections Present in PPT"
End If        'End of display no sections once
End If        'End of IF there are sections

End Function

Option Explicit

Function presentationActions(currentPresentation As Presentation)
    'ENTIRE PRESENTATION
    
End Function

Function sectionStartAction(currentPresentation As Presentation, sld As Slide)
    'FIRST SLIDE AT THE BEGINNING A NEW SECTION
    
End Function

Function sectionEndAction(currentPresentation As Presentation, sld As Slide)
    'LAST SLIDE AT THE END OF SECTION
    
End Function

Function slideActions(sld As Slide)
    'EACH SLIDE
    
End Function

Function slideShapeActions(sld As Slide, shp As Shape)
    'EACH SHAPE OF A SLIDE
      
    
End Function

Function slideCommentActions(sld As Slide, cmt As Comment)
    'EACH COMMENT OF A SLIDE
      
    
End Function

Function slidePlaceholderActions(sld As Slide, shp As Shape)
    'EACH SLIDE PlACEHOLDER SHAPE
    
    Select Case shp.PlaceholderFormat.Type
        Case Is = ppPlaceholderTitle        'TITLE
            
        Case Is = ppPlaceholderObject        'CONTENT
            
        Case Is = ppPlaceholderDate        'DATE
            
        Case Is = ppPlaceholderFooter        'FOOTER
            
        Case Is = ppPlaceholderSlideNumber        'SLIDE NUMBER
            
        Case Is = ppPlaceholderBody        'TEXT
            
        Case Is = ppPlaceholderPicture        'PICTURE
            
        Case Is = ppPlaceholderChart        'CHART
            
        Case Is = ppPlaceholderTable        'TABLE
            
        Case Is = ppPlaceholderOrgChart        'SMARTART
            
        Case Is = ppPlaceholderMediaClip        'MEDIA CLIP
            
    End Select        'End of Placeholder Case Statement
    
End Function

Function slideNonPlaceholderActions(sld As Slide, shp As Shape)
    'EACH SLIDE NON-PLACEHOLDER SHAPE
    
    Select Case shp.Type
        Case Is = msoMedia        'MEDIA OBJECT
            If shp.MediaType = ppMediaTypeMovie Then        'VIDEO
            
        End If
    Case Is = msoTable        'TABLE
        
    Case Is = msoPicture        'PICTURE
                
    Case Is = msoAutoShape        'SHAPE
        
    Case Is = msoSmartArt        'SMARTART
        
    Case Is = msoChart        'CHART
        
    Case Is = msoTextBox        'TEXTBOX
        
End Select

End Function

Function noteShapeActions(sld As Slide, shp As Shape)
    'EACH SHAPE OF A SLIDE NOTES
    
End Function

Function notePlaceholderActions(sld As Slide, shp As Shape)
    'EACH NOTES PlACEHOLDER SHAPE
    
    Select Case shp.PlaceholderFormat.Type
        Case Is = ppPlaceholderTitle        'SLIDE IMAGE
            
        Case Is = ppPlaceholderDate        'DATE
            
        Case Is = ppPlaceholderFooter        'FOOTER
            
        Case Is = ppPlaceholderSlideNumber        'SLIDE NUMBER
            
        Case Is = ppPlaceholderBody        'PRESENTER NOTES
            
            
        Case Is = ppPlaceholderHeader
            
    End Select        'End of Placeholder Case Statement
    
End Function

Function noteNonPlaceholderActions(sld As Slide, shp As Shape)
    'EACH NOTES NON-PLACEHOLDER SHAPE
    
    Select Case shp.Type
        Case Is = msoTable        'TABLE
            
        Case Is = msoPicture        'PICTURE
            
        Case Is = msoAutoShape        'SHAPE
            
        Case Is = msoSmartArt        'SMARTART
            
        Case Is = msoChart        'CHART
            
        Case Is = msoTextBox        'TEXTBOX
            
    End Select
    
End Function

