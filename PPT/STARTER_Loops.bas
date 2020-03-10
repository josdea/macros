Dim debugDetails                                  As String
Option Explicit
Sub Starter()
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
Function iterateSlides(currentPresentation As Presentation)
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
        Debug.Print " Slide " & sld.SlideNumber & " / " & ActivePresentation.Slides.count & " (Layout: " & sld.layout & ")"
        debugDetails = debugDetails & " Slide " & sld.SlideNumber & " / " & ActivePresentation.Slides.count & " (Layout: " & sld.layout & ")" & vbCrLf
        Call slideActions(sld)
        Call iterateSlideComments(sld)
        Call iterateSlideShapes(sld)
        Call iterateNoteShapes(sld)
        
    Next        'Slide
End Function
Function iterateSlideShapes(sld As Slide)
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
                If shp.Name = "LearnerNotes" Then
    
    shp.Delete
    
    
    End If
    Next shp
End Function
Function iterateGroupedSlideShapes(sld, shp)        'TODO iterate group shapes on notes pages
    Dim x                                         As Integer
    Dim shp2                                      As Shape
    For x = 1 To shp.GroupItems.count
        If shp.GroupItems(x).Type = msoGroup Then
            Call iterateGroupedSlideShapes(sld, shp.GroupItems(x))
        Else
            Debug.Print "   Grouped Slide Shape " & x & " / " & shp.GroupItems.count & " Type: " & shp.Type & " (" & shp.GroupItems(x).Name & ")"
            debugDetails = debugDetails & "   Grouped Slide Shape " & x & " / " & shp.GroupItems.count & " Type: " & shp.GroupItems(x).Type & " (" & shp.GroupItems(x).Name & ")" & vbCrLf
            Call slideShapeActions(sld, shp.GroupItems(x))
        End If
    Next
End Function
Function iterateNoteShapes(sld As Slide)
    Dim shp                                       As Shape
    Dim shpCount                                  As Integer
    shpCount = 1
    
    For Each shp In sld.NotesPage.Shapes
        Debug.Print "    Note Shape " & shpCount & " / " & sld.NotesPage.Shapes.count & " Type: " & shp.Type & " (" & shp.Name & ")"
        debugDetails = debugDetails & "    Note Shape " & shpCount & " / " & sld.NotesPage.Shapes.count & " Type: " & shp.Type & " (" & shp.Name & ")" & vbCrLf
        shpCount = shpCount + 1
        Call noteShapeActions(sld, shp)
        If shp.Type = msoGroup Then
            Call iterateGroupedNoteShapes(sld, shp)
        End If
    Next shp
    
End Function
Function iterateGroupedNoteShapes(sld, shp)
    Dim x                                         As Integer
    Dim shp2                                      As Shape
    For x = 1 To shp.GroupItems.count
        If shp.GroupItems(x).Type = msoGroup Then
            Call iterateGroupedSlideShapes(sld, shp.GroupItems(x))
        Else
            Debug.Print "     Grouped Note Shape " & x & " / " & shp.GroupItems.count & " Type: " & shp.Type & " (" & shp.GroupItems(x).Name & ")"
            debugDetails = debugDetails & "     Grouped Note Shape " & x & " / " & shp.GroupItems.count & " Type: " & shp.GroupItems(x).Type & " (" & shp.GroupItems(x).Name & ")" & vbCrLf
            Call slideShapeActions(sld, shp.GroupItems(x))
        End If
    Next
End Function

Function iterateSlideComments(sld As Slide)
    Dim cmt                                       As Comment
    Dim reply  As Comment
    Dim intCommentCount                                  As Integer: intCommentCount = 0
    Dim intReplyCount As Integer: intReplyCount = 0
    
    For Each cmt In sld.Comments
        intCommentCount = intCommentCount + 1
        Call slideCommentActions(sld, cmt)
        
        If cmt.Replies.count > 0 Then
            For Each reply In cmt.Replies
                intReplyCount = intReplyCount + 1
                Call slideCommentActions(sld, reply)
            Next reply
        End If
        
    Next cmt
    If intCommentCount > 0 Or intReplyCount > 0 Then
        
        Debug.Print " Slide Comments: " & intCommentCount & " Replies: " & intReplyCount
        debugDetails = debugDetails & " Slide Comments: " & intCommentCount & " Replies: " & intReplyCount & vbCrLf
        
    End If
    
End Function
Function getSectionName(currentPresentation As Presentation, sld As Slide) As String
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
Function writeFile(Comment As String)
    Dim n                                         As Integer
    n = FreeFile()
    Open Environ("USERPROFILE") & "\Desktop\ppt_report_" & Format(Now(), "yymmdd hhmm") & ".txt" For Output As #n
    'Debug.Print Comment ' write to immediate
    Print #n, Comment        ' write to file
    Close #n
End Function
