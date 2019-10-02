Dim debugDetails                                  As String
Option Explicit
Sub main()
    debugDetails = ""
    Debug.Print "****Start of Main****"
    Debug.Print "Presentation: " & ActivePresentation.Name
    debugDetails = debugDetails & "Presentation: " & ActivePresentation.Name & vbCrLf
    Call presentationActionsStart(ActivePresentation)
    Call iterateSlides(ActivePresentation)
    Call presentationActionsEnd(ActivePresentation)
    Debug.Print "****End of Main****"
    If (MsgBox("Export Section, Slide, And Shape Details To Desktop?", (vbYesNo + vbQuestion), "Export To File?") = vbYes) Then
        Call writeFile(debugDetails)
    End If
End Sub
Function iterateSlides(currentPresentation        As Presentation)
    Dim sld                                       As Slide
    Dim currentSlideNumber                        As Integer
    Dim startingSlideNumber                       As Integer
    
    startingSlideNumber = 1
    
    If (currentPresentation.Application.ActiveWindow.View.Slide.SlideNumber <> 1) Then
        If (MsgBox("Run from slide 1? (Otherwise, it will run from current position)", (vbYesNo + vbQuestion), "Run?") = vbNo) Then
            startingSlideNumber = currentPresentation.Application.ActiveWindow.View.Slide.SlideNumber
        End If
    End If
    
    For currentSlideNumber = startingSlideNumber To currentPresentation.Slides.count
        Set sld = currentPresentation.Slides(currentSlideNumber)
        Call getSectionName(currentPresentation, sld)
        Debug.Print " Slide " & sld.SlideNumber & " / " & ActivePresentation.Slides.count
        debugDetails = debugDetails & " Slide " & sld.SlideNumber & " / " & ActivePresentation.Slides.count & vbCrLf
        Call slideActions(sld)
        Call iterateSlideShapes(sld)
        Call iterateNoteShapes(sld)
        Call iterateSlideComments(sld)
    Next        'Slide
End Function
Function iterateSlideShapes(sld                   As Slide)
    Dim shp                                       As Shape
    Dim shpCount                                  As Integer
    shpCount = 1
    For Each shp In sld.Shapes
        Debug.Print "  Slide Shape " & shpCount & " / " & sld.Shapes.count & " Type: " & shp.Type & " (" & shp.Name & ")"
        debugDetails = debugDetails & "  Slide Shape " & shpCount & " / " & sld.Shapes.count & " Type: " & shp.Type & " (" & shp.Name & ")" & vbCrLf
        shpCount = shpCount + 1
        Call slideShapeActions(sld, shp)
        If shp.Type = msoGroup Then
            Call iterateGroupedSlideShapes(sld, shp)
        End If
    Next shp
End Function
Function iterateGroupedSlideShapes(sld, shp)
    Dim x                                         As Integer
    Dim shp2                                      As Shape
    For x = 1 To shp.GroupItems.count
        If shp.GroupItems(x).Type = msoGroup Then
            Call iterateGroupedSlideShapes(sld, shp.GroupItems(x))
        Else
            Debug.Print "   Grouped Slide Shape " & x & " / " & shp.GroupItems.count & " Type: " & shp.Type & " (" & shp.GroupItems(x).Name & ")"
            debugDetails = debugDetails & "   Grouped Slide Shape " & x & " / " & shp.GroupItems.count & " Type: " & shp.Type & " (" & shp.GroupItems(x).Name & ")" & vbCrLf
            Call slideShapeActions(sld, shp.GroupItems(x))
        End If
    Next
End Function
Function iterateNoteShapes(sld                    As Slide)
    Dim shp                                       As Shape
    Dim shpCount                                  As Integer
    shpCount = 1
    
    For Each shp In sld.NotesPage.Shapes
        Debug.Print "    Note Shape " & shpCount & " / " & sld.NotesPage.Shapes.count & " Type: " & shp.Type & " (" & shp.Name & ")"
        debugDetails = debugDetails & "    Note Shape " & shpCount & " / " & sld.NotesPage.Shapes.count & " Type: " & shp.Type & " (" & shp.Name & ")" & vbCrLf
        shpCount = shpCount + 1
        Call noteShapeActions(sld, shp)
        
    Next shp
    
End Function
Function iterateSlideComments(sld                 As Slide)
    Dim cmt                                       As Comment
    Dim cmtCount                                  As Integer
    cmtCount = 1
    For Each cmt In sld.Comments
        Call slideCommentActions(sld, cmt)
    Next cmt
    
End Function
Function getSectionName(currentPresentation       As Presentation, sld As Slide) As String
    If currentPresentation.SectionProperties.count > 0 Then        'sections exist
    If (sld.SlideNumber = 1) Then        'First slide so output section info
    Debug.Print "Section " & sld.sectionIndex & " of " & currentPresentation.SectionProperties.count & " (" & currentPresentation.SectionProperties.Name(sld.sectionIndex) & ")"
    debugDetails = debugDetails & "Section " & sld.sectionIndex & " of " & currentPresentation.SectionProperties.count & " (" & currentPresentation.SectionProperties.Name(sld.sectionIndex) & ")" & vbCrLf
    Call sectionStartAction(currentPresentation, sld)
ElseIf (sld.sectionIndex <> currentPresentation.Slides(sld.SlideNumber - 1).sectionIndex) Then        'Not the first slide but section index is different than previous slide
Debug.Print "Section " & sld.sectionIndex & " of " & currentPresentation.SectionProperties.count & " (" & currentPresentation.SectionProperties.Name(sld.sectionIndex) & ")"
debugDetails = debugDetails & "Section " & sld.sectionIndex & " of " & currentPresentation.SectionProperties.count & " (" & currentPresentation.SectionProperties.Name(sld.sectionIndex) & ")" & vbCrLf
Call sectionStartAction(currentPresentation, sld)
ElseIf (sld.SlideNumber = currentPresentation.Slides.count) Then        'Last slide of the presentation so the last slide in a section
Call sectionEndAction(currentPresentation, sld)
ElseIf (sld.sectionIndex <> currentPresentation.Slides(sld.SlideNumber + 1).sectionIndex) Then 
    Call sectionEndAction(currentPresentation, sld)
End If        'End of first slide or differing sections IF
Else        'There are no sections in current PPT
    If sld.SlideNumber = 1 Then        'Only display the following, once
    Debug.Print "No Sections Present in PPT"
    debugDetails = debugDetails & "No Sections Present in PPT" & vbCrLf
End If        'End of display no sections once
End If        'End of IF there are sections
End Function
Function writeFile(Comment                        As String)
    Dim n                                         As Integer
    n = FreeFile()
    Open Environ("USERPROFILE") & "\Desktop\ppt_report_" & Format(Now(), "yymmdd hhmm") & ".txt" For Output As #n
    'Debug.Print Comment ' write to immediate
    Print #n, Comment        ' write to file
    Close #n
End Function

Dim counter                                       As Integer
Option Explicit
Function presentationActionsStart(currentPresentation As Presentation)
    counter = 0        'USE THE FOLLOWING IN ANY PLACE TO AUGMENT COUNTER "counter = counter + 1"
    'ENTIRE PRESENTATION - CALLED FIRST
    
End Function
Function presentationActionsEnd(currentPresentation As Presentation)
    'MsgBox "Counter: " & counter     'USE THIS LINE TO COUNT INSTANCES
    'ENTIRE PRESENTATION - CALLED LAST
    
End Function
Function sectionStartAction(currentPresentation   As Presentation, sld As Slide)
    'FIRST SLIDE AT THE BEGINNING A NEW SECTION
    
End Function
Function sectionEndAction(currentPresentation     As Presentation, sld As Slide)
    'LAST SLIDE AT THE END OF SECTION
    
End Function
Function slideActions(sld                         As Slide)
    'EACH SLIDE
    
    If sld.Master.Name = "SlideMaster" Then
        'EACH SLIDE USING DEFAULT MASTER LAYOUT
        
        Select Case sld.Layout
            Case Is = ppLayoutTitle        'TITLE SLIDE
                
            Case Is = ppLayoutObject        'TITLE AND CONTENT
                
            Case Is = ppLayoutSectionHeader        'SECTION HEADER
                
            Case Is = ppLayoutTwoObjects        'TWO CONTENT OBJECTS
                
            Case Is = ppLayoutComparison        'COMPARISON - TWO OBJECTS TWO TEXT BOXES
                
            Case Is = ppLayoutTitleOnly        'JUST TITLE
                
            Case Is = ppLayoutBlank        'NO PLACEHOLDERS
                
            Case Is = ppLayoutContentWithCaption        'TITLE, CONTENT, AND TEXTBOX
                
            Case Is = ppLayoutPictureWithCaption        ' TITLE, PICTURE, AND TEXTBOX
                
        End Select
    Else
        'EACH SLIDE USING A NON DEFAULT MASTER LAYOUT
        
    End If
    
End Function
Function slideShapeActions(sld                    As Variant, shp As Variant)
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
                
            Case Else        'EACH SHAPE THAT IS A PLACEHOLDER BUT NOT THE ABOVE
                
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
            
        Case Is = msoGroup        'EACH GROUP OF SHAPES OR GROUP OF GROUPS
            
    End Select
End If
End Function
Function slideCommentActions(sld                  As Slide, cmt As Comment)
    'EACH COMMENT OF A SLIDE
    
End Function
Function noteShapeActions(sld                     As Slide, shp As Shape)
    'EACH SHAPE OF SLIDE NOTES
    
    If shp.Type = msoPlaceholder Then
        'EACH NOTES PlACEHOLDER SHAPE
        
        Select Case shp.PlaceholderFormat.Type
            Case Is = ppPlaceholderTitle        'EACH SLIDE IMAGE PLACEHOLDER
                
            Case Is = ppPlaceholderDate        'EACH DATE PLACEHOLDER
                
            Case Is = ppPlaceholderFooter        'EACH FOOTER PLACEHOLDER
                
            Case Is = ppPlaceholderSlideNumber        'EACH SLIDE NUMBER PLACEHOLDER
                
            Case Is = ppPlaceholderBody        'EACH PRESENTER NOTES PLACEHOLDER
                
            Case Is = ppPlaceholderHeader        'EACH HEADER PLACEHOLDER
                
            Case Else        'EACH SHAPE THAT IS A PLACEHOLDER BUT NOT THE ABOVE
                
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
                
            Case Is = msoGroup        'EACH GROUP OF SHAPES
                
            Case Else        'EACH SHAPE THAT IS NOT A PLACEHOLDER AND NOT THE ABOVE
                
        End Select
    End If
End Function