Option Explicit


Sub createSlideTextBoxesAllSlides()
     Debug.Print "Start of Create all TextBoxes on all slides";
     
    Dim sld                                       As Slide        ' declare slide object
    
    For Each sld In ActivePresentation.Slides        ' iterate slides
        
       ' getSectionName sld
        Debug.Print "Slide " & sld.SlideNumber & " of " & ActivePresentation.Slides.Count
        
       ' ActivePresentation.Slides(sld.SlideNumber).Select
        iterateSlideShapes sld        ' Call to iterate shapes on slide
      
        
    Next sld        ' end of iterate slides
    
       
    Debug.Print "All done creating slide textboxes";
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
    
    If hasLearnerNotes = False Then
        createSlideShapeTextbox sld, -4 * 72, 0, 3.5 * 72, 7.5 * 72, "LearnerNotes", "", 12
        
    End If
    
    If hasMinutes = False Then
        createSlideShapeTextbox sld, 0, -1 * 72, 2 * 72, 0.5 * 72, "Minutes", "", 12
    End If
    
    If hasObjective = False Then
        createSlideShapeTextbox sld, 2.5 * 72, -1 * 72, 3 * 72, 0.5 * 72, "Objective", "", 12
    End If
    
   
    
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




