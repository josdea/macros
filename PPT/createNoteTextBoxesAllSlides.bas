Option Explicit

Sub createNoteTextBoxesAllSlides()
    ActiveWindow.ViewType = ppViewNotesPage
    Dim sld                                       As Slide        ' declare slide object
    
    For Each sld In ActivePresentation.Slides        ' iterate slides
        Debug.Print "Slide " & sld.SlideNumber & " of " & ActivePresentation.Slides.Count
        iterateNotesPageShapes sld        ' Call to iterate shapes on slide
    Next sld        ' end of iterate slides
    If Not CommandBars.GetPressedMso("SelectionPane") Then CommandBars.ExecuteMso ("SelectionPane")
    MsgBox "All done creating textboxes On notes"
    
End Sub

Sub iterateNotesPageShapes(sld As Slide)
    Dim shp                                       As Shape
    Dim hasModuleTitle                            As Boolean: hasModuleTitle = False
    Dim hasLearnerNotes                           As Boolean: hasLearnerNotes = False
    Dim hasObjective                              As Boolean: hasObjective = False
    Dim hasMinutes                                As Boolean: hasMinutes = False
    
    For Each shp In sld.NotesPage.Shapes
        Select Case shp.name
            Case Is = "Objective"        ' objective
                hasObjective = True
                Call updateObjective(sld, shp)
            Case Is = "Minutes"
                hasMinutes = True
                Call updateMinutes(sld, shp)
            Case Is = "ModuleTitle"
                hasModuleTitle = True
                Call updateModuleTitle(sld, shp)
            Case Is = "LearnerNotes"
                hasLearnerNotes = True
                Call updateLearnerNotes(sld, shp)
            Case Else
                If shp.Type = msoPlaceholder Then        ' check if its a placeholder
                Select Case shp.PlaceholderFormat.Type        ' type of placeholder
                    Case Is = ppPlaceholderFooter        ' footer shape
                        Call updateFooter(sld, shp)
                    Case Is = ppPlaceholderSlideNumber        ' Slide Number
                        Call updateSlideNumber(sld, shp)
                    Case Is = ppPlaceholderBody        ' Presenter notes
                        Call updatePresenterNotes(sld, shp)
                    Case Is = ppPlaceholderTitle        ' Slide Image Placeholder
                        Call updateSlideImagePlaceholder(sld, shp)
                    Case Else
                End Select
            End If
    End Select        ' end of case
Next shp

If hasObjective = False Then
    Set shp = sld.NotesPage.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=0, Top:=0, Width:=1, Height:=1)
    Call updateObjective(sld, shp)
End If
If hasMinutes = False Then
    Set shp = sld.NotesPage.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=0, Top:=0, Width:=1, Height:=1)
    Call updateMinutes(sld, shp)
End If
If hasModuleTitle = False Then
    Set shp = sld.NotesPage.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=0, Top:=0, Width:=1, Height:=1)
    Call updateModuleTitle(sld, shp)
End If
If hasLearnerNotes = False Then
    Set shp = sld.NotesPage.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=0, Top:=0, Width:=1, Height:=1)
    Call updateLearnerNotes(sld, shp)
End If
End Sub

Function updateObjective(sld As Slide, shp As Shape)
    Dim name                                      As String: name = "Objective"
    Dim shpHeight                                 As Double: shpHeight = 0.3 * 72
    Dim shpWidth                                  As Double: shpWidth = sld.NotesPage.Master.Width * 0.25
    Dim shpHorizontal                             As Double: shpHorizontal = sld.NotesPage.Master.Width - shpWidth
    Dim shpVertical                               As Double: shpVertical = sld.NotesPage.Master.Height - (shpHeight * 2)        ' (2) because its above slide number
    Dim shpFontSize                               As Double: shpFontSize = 10
    Dim shpWordWrap                               As Integer: shpWordWrap = msoFalse
    Dim shpUnderline                              As Boolean: shpUnderline = False
    Dim shpAlign                                  As Integer: shpAlign = msoTextEffectAlignmentRight
    Dim shpFontBold                               As Integer: shpFontBold = msoFalse
    Dim shpFontItalic                             As Integer: shpFontItalic = msoTrue
    Dim shpFontColor                              As Double: shpFontColor = RGB(0, 0, 0)
    Call updateShape(shp, name, shpHeight, shpWidth, shpHorizontal, shpVertical, shpFontSize, shpWordWrap, shpUnderline, shpAlign, shpFontBold, shpFontItalic, shpFontColor)
    If shp.TextFrame.HasText = False Then
        shp.TextFrame.TextRange.Text = "Covering Objective: "
    End If
End Function

Function updateSlideNumber(sld As Slide, shp As Shape)
    Dim name                                      As String: name = "Slide Number Placeholder"
    Dim shpHeight                                 As Double: shpHeight = 0.3 * 72
    Dim shpWidth                                  As Double: shpWidth = sld.NotesPage.Master.Width * 0.25
    Dim shpHorizontal                             As Double: shpHorizontal = sld.NotesPage.Master.Width - shpWidth
    Dim shpVertical                               As Double: shpVertical = sld.NotesPage.Master.Height - shpHeight        ' (2) because its above slide number
    Dim shpFontSize                               As Double: shpFontSize = 10
    Dim shpWordWrap                               As Integer: shpWordWrap = msoFalse
    Dim shpUnderline                              As Boolean: shpUnderline = False
    Dim shpAlign                                  As Integer: shpAlign = msoTextEffectAlignmentRight
    Dim shpFontBold                               As Integer: shpFontBold = msoFalse
    Dim shpFontItalic                             As Integer: shpFontItalic = msoFalse
    Dim shpFontColor                              As Double: shpFontColor = RGB(0, 0, 0)
    Call updateShape(shp, name, shpHeight, shpWidth, shpHorizontal, shpVertical, shpFontSize, shpWordWrap, shpUnderline, shpAlign, shpFontBold, shpFontItalic, shpFontColor)
End Function

Function updateModuleTitle(sld As Slide, shp As Shape)
    Dim name                                      As String: name = "ModuleTitle"
    Dim shpHeight                                 As Double: shpHeight = 0.3 * 72
    Dim shpWidth                                  As Double: shpWidth = sld.NotesPage.Master.Width
    Dim shpHorizontal                             As Double: shpHorizontal = 0
    Dim shpVertical                               As Double: shpVertical = 0
    Dim shpFontSize                               As Double: shpFontSize = 10
    Dim shpWordWrap                               As Integer: shpWordWrap = msoFalse
    Dim shpUnderline                              As Boolean: shpUnderline = False
    Dim shpAlign                                  As Integer: shpAlign = msoTextEffectAlignmentCentered
    Dim shpFontBold                               As Integer: shpFontBold = msoFalse
    Dim shpFontItalic                             As Integer: shpFontItalic = msoFalse
    Dim shpFontColor                              As Double: shpFontColor = RGB(0, 0, 0)
    Call updateShape(shp, name, shpHeight, shpWidth, shpHorizontal, shpVertical, shpFontSize, shpWordWrap, shpUnderline, shpAlign, shpFontBold, shpFontItalic, shpFontColor)
End Function

Function updateMinutes(sld As Slide, shp As Shape)
    Dim name                                      As String: name = "Minutes"
    Dim shpHeight                                 As Double: shpHeight = 0.3 * 72
    Dim shpWidth                                  As Double: shpWidth = sld.NotesPage.Master.Width * 0.75
    Dim shpHorizontal                             As Double: shpHorizontal = 0
    Dim shpVertical                               As Double: shpVertical = sld.NotesPage.Master.Height - (shpHeight * 2)
    Dim shpFontSize                               As Double: shpFontSize = 10
    Dim shpWordWrap                               As Integer: shpWordWrap = msoFalse
    Dim shpUnderline                              As Boolean: shpUnderline = False
    Dim shpAlign                                  As Integer: shpAlign = msoTextEffectAlignmentLeft
    Dim shpFontBold                               As Integer: shpFontBold = msoFalse
    Dim shpFontItalic                             As Integer: shpFontItalic = msoFalse
    Dim shpFontColor                              As Double: shpFontColor = RGB(0, 0, 0)
    Call updateShape(shp, name, shpHeight, shpWidth, shpHorizontal, shpVertical, shpFontSize, shpWordWrap, shpUnderline, shpAlign, shpFontBold, shpFontItalic, shpFontColor)
    If shp.TextFrame.HasText = False Then
        shp.TextFrame.TextRange.Text = "Minutes: "
    End If
End Function

Function updateFooter(sld As Slide, shp As Shape)
    Dim name                                      As String: name = "Footer Placeholder"
    Dim shpHeight                                 As Double: shpHeight = 0.3 * 72
    Dim shpWidth                                  As Double: shpWidth = sld.NotesPage.Master.Width * 0.75
    Dim shpHorizontal                             As Double: shpHorizontal = 0
    Dim shpVertical                               As Double: shpVertical = sld.NotesPage.Master.Height - shpHeight
    Dim shpFontSize                               As Double: shpFontSize = 10
    Dim shpWordWrap                               As Integer: shpWordWrap = msoFalse
    Dim shpUnderline                              As Boolean: shpUnderline = False
    Dim shpAlign                                  As Integer: shpAlign = msoTextEffectAlignmentLeft
    Dim shpFontBold                               As Integer: shpFontBold = msoFalse
    Dim shpFontItalic                             As Integer: shpFontItalic = msoFalse
    Dim shpFontColor                              As Double: shpFontColor = RGB(0, 0, 0)
    Call updateShape(shp, name, shpHeight, shpWidth, shpHorizontal, shpVertical, shpFontSize, shpWordWrap, shpUnderline, shpAlign, shpFontBold, shpFontItalic, shpFontColor)
End Function

Function updatePresenterNotes(sld As Slide, shp As Shape)
    Dim name                                      As String: name = "Notes Placeholder"
    Dim shpHeight                                 As Double: shpHeight = sld.NotesPage.Master.Height - (0.3 * 4 * 72)
    Dim shpWidth                                  As Double: shpWidth = sld.NotesPage.Master.Width * 0.4
    Dim shpHorizontal                             As Double: shpHorizontal = sld.NotesPage.Master.Width - sld.NotesPage.Master.Width * 0.4
    Dim shpVertical                               As Double: shpVertical = 0.3 * 2 * 72
    Dim shpFontSize                               As Double: shpFontSize = 10
    Dim shpWordWrap                               As Integer: shpWordWrap = msoTrue
    Dim shpUnderline                              As Boolean: shpUnderline = False
    Dim shpAlign                                  As Integer: shpAlign = msoTextEffectAlignmentLeft
    Dim shpFontBold                               As Integer: shpFontBold = msoFalse
    Dim shpFontItalic                             As Integer: shpFontItalic = msoFalse
    Dim shpFontColor                              As Double: shpFontColor = RGB(0, 0, 0)
    Call updateShape(shp, name, shpHeight, shpWidth, shpHorizontal, shpVertical, shpFontSize, shpWordWrap, shpUnderline, shpAlign, shpFontBold, shpFontItalic, shpFontColor)
End Function

Function updateLearnerNotes(sld As Slide, shp As Shape)
    Dim name                                      As String: name = "LearnerNotes"
    Dim shpHeight                                 As Double: shpHeight = sld.NotesPage.Master.Height - (0.3 * 4 * 72) - (2.45 * 72)
    Dim shpWidth                                  As Double: shpWidth = sld.NotesPage.Master.Width * 0.6
    Dim shpHorizontal                             As Double: shpHorizontal = 0
    Dim shpVertical                               As Double: shpVertical = ((0.3 * 2) + 2.45) * 72
    Dim shpFontSize                               As Double: shpFontSize = 10
    Dim shpWordWrap                               As Integer: shpWordWrap = msoTrue
    Dim shpUnderline                              As Boolean: shpUnderline = False
    Dim shpAlign                                  As Integer: shpAlign = msoTextEffectAlignmentLeft
    Dim shpFontBold                               As Integer: shpFontBold = msoFalse
    Dim shpFontItalic                             As Integer: shpFontItalic = msoFalse
    Dim shpFontColor                              As Double: shpFontColor = RGB(0, 0, 0)
    Call updateShape(shp, name, shpHeight, shpWidth, shpHorizontal, shpVertical, shpFontSize, shpWordWrap, shpUnderline, shpAlign, shpFontBold, shpFontItalic, shpFontColor)
    With shp.TextFrame.TextRange.ParagraphFormat
        .SpaceAfter = 0
        .SpaceBefore = 0
        .SpaceWithin = 1.5
    End With
End Function

Function updateSlideImagePlaceholder(sld As Slide, shp As Shape)
    Dim name                                      As String: name = "Slide Image Placeholder"
    Dim shpHeight                                 As Double: shpHeight = 2.25 * 72
    Dim shpWidth                                  As Double: shpWidth = 4 * 72
    Dim shpHorizontal                             As Double: shpHorizontal = ((sld.NotesPage.Master.Width * 0.6) - shpWidth) / 2
    Dim shpVertical                               As Double: shpVertical = 0.3 * 2 * 72
    
    With shp
        .name = name
        .Width = shpWidth        'shape height
        .Height = shpHeight        'shape width
        .Left = shpHorizontal        'shape position from left
        .Top = shpVertical        'shape position from top
       ' .ZOrder (msoBringToFront)
    End With
    
End Function

Function updateShape(shp As Shape, name As String, shpHeight As Double, shpWidth As Double, shpHorizontal As Double, shpVertical As Double, shpFontSize As Double, shpWordWrap As Integer, shpUnderline As Boolean, shpAlign As Integer, shpFontBold As Integer, shpFontItalic As Integer, shpFontColor As Double)
    
    With shp
        .name = name
        .Width = shpWidth        'shape height
        .Height = shpHeight        'shape width
        .Left = shpHorizontal        'shape position from left
        .Top = shpVertical        'shape position from top
        With .TextFrame2
            .AutoSize = msoAutoSizeTextToFitShape        'resize font size to fit if too big
            .WordWrap = shpWordWrap        'word wrap for textbox
        End With
        With .TextFrame.TextRange.Font
            .size = shpFontSize        'font size
            .name = "+mn-lt"        'set font to theme
            .underline = shpUnderline        'remove underline
            .color.RGB = shpFontColor
        End With
        With .TextEffect
            .Alignment = shpAlign        'text align
            .FontBold = shpFontBold        'bold text
            .FontItalic = shpFontItalic        'italic text
        End With
        With .Fill
            .ForeColor.RGB = RGB(255, 255, 255)
            .BackColor.RGB = RGB(255, 255, 255)
        End With
    End With
    
End Function
