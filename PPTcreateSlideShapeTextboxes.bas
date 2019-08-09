Option Explicit

Sub createNoteTextBoxesAllSlides()

 Debug.Print "Start of Create all TextBoxes on all slides";
     
    Dim sld                                       As Slide        ' declare slide object
    
    ActiveWindow.ViewType = ppViewNotesPage

    For Each sld In ActivePresentation.Slides        ' iterate slides
       
        Debug.Print "Slide " & sld.SlideNumber & " of " & ActivePresentation.Slides.Count
        
       ' ActivePresentation.Slides(sld.SlideNumber).Select
        iterateNotesPageShapes sld        ' Call to iterate shapes on slide
        
    Next sld        ' end of iterate slides
       
    Debug.Print "All Done";
     If Not CommandBars.GetPressedMso("SelectionPane") Then CommandBars.ExecuteMso ("SelectionPane")
    MsgBox "All Done"
    
End Sub

Sub iterateNotesPageShapes(sld As Slide)
    
    Dim shp                                       As Shape        ' declare shape object
    Dim hasModuleTitle As Boolean
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
                 updateSizePosition sld, 5.5, 9.38, 2, 0.3, "Objective", "", 11
                
            Case Is = "Minutes"
                hasMinutes = True
                 updateSizePosition sld, 0, 9.37, 2, 0.3, "Minutes", "", 11
                
            Case Is = "LearnerNotes"
                hasLearnerNotes = True
                updateSizePosition sld, 0, 3, 4.75, 6.25, "LearnerNotes", "", 11

           Case Is = "ModuleTitle"
                hasModuleTitle = True
                updateSizePosition shp, 0.75, 0, 6, 0.3, "ModuleTitle2", "", 11
            Case Else
                ' Debug.Print "NOTE: should not reach this point unless there is another text placeholder on the slide TODO"
        End Select        ' end of case
        
        '    End If        ' end of textframe has text
        
    Next shp
    
    If hasModuleTitle = False Then
        createSlideShapeTextbox sld, 0.75, 0, 6, 0.3, "ModuleTitle", "", 11
    End If
    

    If hasLearnerNotes = False Then
        createSlideShapeTextbox sld, 0, 3, 4.75, 6.25, "LearnerNotes", "", 11
    End If
    
    If hasObjective = False Then
        createSlideShapeTextbox sld, 5.5, 9.38, 2, 0.3, "Objective", "", 11
    End If

    If hasMinutes = False Then
        createSlideShapeTextbox sld, 0, 9.37, 2, 0.3, "Minutes", "", 11
    End If
    
End Sub


Sub createSlideShapeTextbox(sld As Slide, shpLeftPosition As Long, shpTopPosition As Long, shpWidth As Long, shpHeight As Long, shpName As String, shpText As String, shpFontSize As Integer)
    
    Dim shp                                       As Shape
    Set shp = sld.NotesPage.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=shpLeftPosition * 72, Top:=shpTopPosition * 72, Width:=shpWidth * 72, Height:=shpWidth * 72)
    With shp
        .TextFrame.TextRange.Text = shpText
        .name = shpName
        .TextFrame2.AutoSize = msoAutoSizeTextToFitShape
        .Height = shpHeight * 72
        .Width = shpWidth * 72
        .TextFrame.TextRange.Font.size = shpFontSize
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
    
End Sub

Sub updateSizePosition(shp As Shape, shpLeftPosition As Long, shpTopPosition As Long, shpWidth As Long, shpHeight As Long, shpName As String, shpText As String, shpFontSize As Integer)

   With shp
        .Left = shpLeftPosition * 72
        .Top = shpTopPosition * 72
        .Width = shpWidth * 72
        .Height = shpHeight * 72
        .name = shpName
        .TextFrame2.AutoSize = msoAutoSizeTextToFitShape
        .TextFrame.TextRange.Font.size = shpFontSize
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With


End Sub







