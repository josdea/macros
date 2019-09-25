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
    Next shp
End Function
Function iterateNoteShapes(sld As Slide)
    Dim shp                                       As Shape
    Dim shpCount                                  As Integer
    shpCount = 1
    
    For Each shp In sld.NotesPage.Shapes
        Debug.Print "   Note Shape " & shpCount & " / " & sld.NotesPage.Shapes.Count & " Type: " & shp.Type & " (" & shp.Name & ")"
        shpCount = shpCount + 1
        Call noteShapeActions(sld, shp)
        
    Next shp
    
End Function
Function iterateSlideComments(sld As Slide)
    Dim cmt                                       As Comment
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
  
  If sld.Master.Name = "SlideMaster" Then
  'EACH SLIDE USING DEFAULT MASTER LAYOUT
  
    Select Case sld.Layout
    Case Is = ppLayoutTitle 'TITLE SLIDE
    
    Case Is = ppLayoutObject 'TITLE AND CONTENT
    
    Case Is = ppLayoutSectionHeader 'SECTION HEADER
    
    Case Is = ppLayoutTwoObjects 'TWO CONTENT OBJECTS
    
    Case Is = ppLayoutComparison 'COMPARISON - TWO OBJECTS TWO TEXT BOXES
    
    Case Is = ppLayoutTitleOnly 'JUST TITLE
    
    Case Is = ppLayoutBlank 'NO PLACEHOLDERS
    
    Case Is = ppLayoutContentWithCaption 'TITLE, CONTENT, AND TEXTBOX
    
    Case Is = ppLayoutPictureWithCaption ' TITLE, PICTURE, AND TEXTBOX
    
  End Select
  Else
    'EACH SLIDE USING A NON DEFAULT MASTER LAYOUT
  
  End If
  
End Function
Function slideShapeActions(sld As Slide, shp As Shape)
    'EACH SHAPE OF A SLIDE
    
    If shp.Type = msoPlaceholder Then
        'EACH SLIDE PlACEHOLDER SHAPE
        
        Select Case shp.PlaceholderFormat.Type
            Case Is = ppPlaceholderTitle        'EACH TITLE PLACEHOLDER
                
            Case Is = ppPlaceholderObject        'EACH CONTENT PLACEHOLDER
                
            Case Is = ppPlaceholderDate        'EACH DATE PLACEHOLDER
                
            Case Is = ppPlaceholderFooter        'EACH FOOTER PLACEHOLDER
                
            Case Is = ppPlaceholderSlideNumber        'EACH SLIDE NUMBER PLACEHOLDER
                
            Case Is = ppPlaceholderBody        'EACH TEXT PLACEHOLDER
                
            Case Is = ppPlaceholderPicture        'EACH PICTURE PLACEHOLDER
                
            Case Is = ppPlaceholderChart        'EACH CHART PLACEHOLDER
                
            Case Is = ppPlaceholderTable        'EACH TABLE PLACEHOLDER
                
            Case Is = ppPlaceholderOrgChart        'EACH SMARTART PLACEHOLDER
                
            Case Is = ppPlaceholderMediaClip        'EACH MEDIA CLIP PLACEHOLDER
                
        End Select        'End of Placeholder Case Statement
    Else
        'EACH SLIDE NON-PLACEHOLDER SHAPE
        
        Select Case shp.Type
            Case Is = msoMedia        'EACH MEDIA OBJECT NON-PLACEHOLDER
                
                If shp.MediaType = ppMediaTypeMovie Then        'EACH VIDEO NON-PLACEHOLDER
                
            End If
        Case Is = msoTable        'EACH TABLE NON-PLACEHOLDER
            
        Case Is = msoPicture        'EACH PICTURE NON-PLACEHOLDER
            
        Case Is = msoAutoShape        'EACH SHAPE NON-PLACEHOLDER
            
        Case Is = msoSmartArt        'EACH SMARTART NON-PLACEHOLDER
            
        Case Is = msoChart        'EACH CHART NON-PLACEHOLDER
            
        Case Is = msoTextBox        'EACH TEXTBOX NON-PLACEHOLDER
            
    End Select
End If
End Function
Function slideCommentActions(sld As Slide, cmt As Comment)
    'EACH COMMENT OF A SLIDE
    
End Function
Function noteShapeActions(sld As Slide, shp As Shape)
    'EACH SHAPE OF SLIDE NOTES
    
    If shp.Type = msoPlaceholder Then
        'EACH NOTES PlACEHOLDER SHAPE
        
        Select Case shp.PlaceholderFormat.Type
            Case Is = ppPlaceholderTitle        'EACH SLIDE IMAGE PLACEHOLDER
                
            Case Is = ppPlaceholderDate        'EACH DATE PLACEHOLDER
                
            Case Is = ppPlaceholderFooter        'EACH FOOTER PLACEHOLDER
                
            Case Is = ppPlaceholderSlideNumber        'EACH SLIDE NUMBER PLACEHOLDER
                
            Case Is = ppPlaceholderBody        'EACH PRESENTER NOTES PLACEHOLDER
                shp.Visible = True
                
            Case Is = ppPlaceholderHeader        'EACH HEADER PLACEHOLDER
                
        End Select
    Else
        'EACH NOTES NON-PLACEHOLDER SHAPE
        
        Select Case shp.Type
            Case Is = msoTable        'EACH TABLE NON-PLACEHOLDER
                
            Case Is = msoPicture        'EACH PICTURE NON-PLACEHOLDER
                
            Case Is = msoAutoShape        'EACH SHAPE NON-PLACEHOLDER
                
            Case Is = msoSmartArt        'EACH SMARTART NON-PLACEHOLDER
                
            Case Is = msoChart        'EACH CHART NON-PLACEHOLDER
                
            Case Is = msoTextBox        'EACH TEXTBOX NON-PLACEHOLDER
                
        End Select
    End If
End Function
