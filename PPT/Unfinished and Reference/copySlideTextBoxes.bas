Option Explicit

Sub copySlideTextBoxes()
    
    Debug.Print "Start of copy slide textboxes to notes";
 '   Call createNoteTextBoxesAllSlides.createNoteTextBoxesAllSlides
    Dim sld                                       As Slide        ' declare slide object
    
    For Each sld In ActivePresentation.Slides        ' iterate slides
        
        getSectionName sld
        Debug.Print "Slide " & sld.SlideNumber & " of " & ActivePresentation.Slides.Count
        
       ' ActivePresentation.Slides(sld.SlideNumber).Select
        iterateSlideShapes sld        ' Call to iterate shapes on slide
        iterateNoteShapes sld        ' Call to iterate shapes on notes
        
    Next sld        ' end of iterate slides
    
    
       
    Debug.Print "All done copying slide textboxes to notes textboxes";
     If Not CommandBars.GetPressedMso("SelectionPane") Then CommandBars.ExecuteMso ("SelectionPane")
    MsgBox "All done creating slide textboxes"
End Sub

Sub iterateSlideShapes(sld As Slide)
    
    Dim shp                                       As Shape        ' declare shape object
    Dim shapeCount                                As Integer
    shapeCount = sld.Shapes.Count
    ' Dim textPlaceholderCount                      As Integer        ' TODO ASSUMPTION objective is first, minutes, and then learner notes
    'textPlaceholderCount = 0
    Dim learnerNotesText                              As String
    Dim objectiveText                                 As String
    Dim minutesText                                   As String
    Dim moduleTitleText                               As String
    Dim hasLearnerNotes                           As Boolean
    Dim hasObjective                              As Boolean
    Dim hasMinutes                                As Boolean
    
    hasLearnerNotes = False
    hasObjective = False
    hasMinutes = False
    
    ' Debug.Print shapeCount & " Shapes on Slide"
    
    For Each shp In sld.Shapes
        
        Select Case shp.name
            Case Is = "Objective"        ' objective
                hasObjective = True
                If shp.TextFrame.HasText Then
                    objectiveText = shp.TextFrame.TextRange
                End If
            Case Is = "Minutes"
                hasMinutes = True
                If shp.TextFrame.HasText Then
                    minutesText = shp.TextFrame.TextRange
                End If
            Case Is = "LearnerNotes"
                hasLearnerNotes = True
                If shp.TextFrame.HasText Then
                    learnerNotesText = shp.TextFrame.TextRange
                End If
            Case Else
                ' Debug.Print "NOTE: should not reach this point unless there is another text placeholder on the slide TODO"
        End Select        ' end of case
        
        '    End If        ' end of textframe has text
        
    Next shp
    
  
    
    setNoteShapes sld, learnerNotesText, objectiveText, minutesText, getSectionName(sld)        ' Call to move the text from the slide to the notes page
    
End Sub

Sub setNoteShapes(sld As Slide, learnerNotes As String, objective As String, minutes As String, moduleTitle As String)
    
    Dim shp                                       As Shape
    Dim fontSize                                  As Integer
    fontSize = 11
    
    Debug.Print ("Looping through shapes on slide " & sld.SlideNumber & " to find and set objective, minutes, and learnernotes if they exist")
    For Each shp In sld.NotesPage.Shapes
        
        If InStr(shp.name, "Notes Placeholder") Then
            Debug.Print "Shape that follows has notes placeholder in the shape name and is the presenter notes"
            updateShapeFont shp, fontSize, "+mn-lt", False, False, False, RGB(0, 0, 0)
        End If
        
        Select Case shp.name
            Case Is = "Objective"
                shp.TextFrame.TextRange.Text = "Covering Objective: " & objective
                updateShapeFont shp, 11, "+mn-lt", False, True, False, RGB(0, 0, 0)
                Debug.Print "Updating Objective in Slide " & sld.SlideNumber
            Case Is = "Minutes"
                shp.TextFrame.TextRange.Text = "Minutes: " & minutes
                updateShapeFont shp, 11, "+mn-lt", False, False, False, RGB(0, 0, 0)
                Debug.Print "Updating Minutes in Slide " & sld.SlideNumber
            Case Is = "LearnerNotes"
                shp.TextFrame.TextRange.Text = learnerNotes
                updateShapeFont shp, 11, "+mn-lt", False, False, False, RGB(0, 0, 0)
                Debug.Print "Updating LearnerNotes in Slide " & sld.SlideNumber
            Case Is = "ModuleTitle"
                shp.TextFrame.TextRange.Text = moduleTitle
                updateShapeFont shp, 11, "+mn-lt", False, False, False, RGB(0, 0, 0)
                Debug.Print "Updating Module Title in Slide " & sld.SlideNumber
            Case Else
                
        End Select
        
    Next shp
    
End Sub

Sub updateShapeFont(shp As Shape, size As Integer, name As String, bold As Boolean, italic As Boolean, underline As Boolean, fontColor As String)
    
    With shp.TextFrame.TextRange.Font
        .size = size        'font size
        .name = name        'font name - use "+mn-lt" to inherit main selected body font from ppt theme
        .bold = bold
        .italic = italic
        .underline = underline
        .color = fontColor
    End With
    
End Sub

Sub createSlideShapeTextbox(sld As Slide, shpLeftPosition As Long, shpTopPosition As Long, shpWidth As Long, shpHeight As Long, shpName As String, shpText As String, shpFontSize As Integer)
    
    Dim shp                                       As Shape
    Set shp = sld.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=shpLeftPosition, Top:=shpTopPosition, Width:=shpWidth, Height:=shpWidth)
    With shp
        .TextFrame.TextRange.Text = shpText
        .name = shpName
        .TextFrame2.AutoSize = msoAutoSizeTextToFitShape
        .Height = shpHeight
        .Width = shpWidth
        .TextFrame.TextRange.Font.size = shpFontSize
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
    
End Sub

Function getSectionName(sld As Slide) As String
    Dim sectionCount                              As Long        ' total sections in ppt
    Dim sectionIndex                              As Long        ' section id of current slide
    
    ' Output Section Info
    sectionCount = ActivePresentation.SectionProperties.Count
    If sectionCount > 0 Then        ' If Sections exist
    If sectionIndex <> sld.sectionIndex Then        ' If this section is different than the last or its the first one
    sectionIndex = sld.sectionIndex
    getSectionName = ActivePresentation.SectionProperties.name(sectionIndex)
    Debug.Print "Section " & sectionIndex & " of " & sectionCount & " | " & ActivePresentation.SectionProperties.name(sectionIndex)
End If        ' end display section title only once
Else        ' sectionCount > 0
    If sld.SlideNumber < 2 Then        ' If there no sections then only display this message once
    Debug.Print "No Sections Present in PPT"
End If
End If        ' end of section count > 0 if

End Function

