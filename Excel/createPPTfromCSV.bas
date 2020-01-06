Option Explicit

Public Sub Create_PPT_From_CSV()
    Dim oPresentation As Object
    Set oPresentation = createNewPPT()
    Call LoopCells(oPresentation)
    oPresentation.slides(1).Select
    Set oPresentation = Nothing
    MsgBox ("All Done")
End Sub

Private Function LoopCells(oPresentation As Object)
    
    Dim c      As Range
    Dim myRange As Range
    Dim cellValue As String
    Dim cellAddress As String
    Dim cellColumn As Integer
    Dim cellRow As Integer
    
    Dim strCourseTitle As String
    Dim strClient As String
    Dim strCourseDuration As String
    
    Dim strModule As String
    Dim strSubtitle As String
    Dim strDescription As String
    Dim strInstructor As String
    Dim strModuleDuration As String
    Dim strTopicTitle As String
    Dim strObjective As String
    Dim strSlideText As String
    Dim strPGNotes As String
    Dim strIGNotes As String
    Dim strHasExercise As String
    Dim strExerciseTitle As String
    Dim strExerciseDescription As String
    Dim strMediaRequired As String
    Dim strMediaDetails As String
    Dim intPreviousRow As Integer
    intPreviousRow = 7
    Dim strModulePrevious
    Dim intModuleNumber As Integer
    intModuleNumber = 1
    
    Set myRange = ActiveSheet.usedRange
    For Each c In myRange
        
        ' c.Select
        cellValue = c.Value
        cellAddress = c.Address
        cellColumn = c.Column
        cellRow = c.Row
        
        If cellColumn = 2 And cellRow = 1 Then        ' course title
        strCourseTitle = cleanupString(cellValue)
    End If        'end if course title
    If cellColumn = 2 And cellRow = 2 Then        'course client
    strClient = cleanupString(cellValue)
    Call createTitleSlide(strCourseTitle, strClient, "", "", "", oPresentation)
End If        ' end if course client

If cellRow > 6 Then        ' get content
Select Case cellColumn
    Case 1
        strModule = cleanupString(cellValue)
        If intPreviousRow <> cellRow Then        'a new row and not the first so create slide from assembled stuff
        Call createSlide(strModule, strSubtitle, strDescription, strInstructor, strModuleDuration, strTopicTitle, strObjective, strSlideText, strPGNotes, strIGNotes, strHasExercise, strExerciseTitle, strExerciseDescription, strMediaRequired, strMediaDetails, oPresentation)
        intPreviousRow = cellRow
    End If
Case 2
    strSubtitle = cleanupString(cellValue)
Case 3
    strDescription = cleanupString(cellValue)
Case 4
    strInstructor = cleanupString(cellValue)
Case 5
    strModuleDuration = cleanupString(cellValue)
    If StrComp(strModulePrevious, cleanupString(strModule)) <> 0 Then        'module title is different or first so create section
    Call createSectionTitleSlide("Module " & intModuleNumber & ": " & strModule, strSubtitle, strDescription, strInstructor, strModuleDuration, oPresentation)
    intModuleNumber = intModuleNumber + 1
    strModulePrevious = cleanupString(strModule)
End If
Case 6
    strTopicTitle = cleanupString(cellValue)
Case 7
    strObjective = cleanupString(cellValue)
Case 8
    strSlideText = cleanupString(cellValue)
Case 9
    strPGNotes = cleanupString(cellValue)
Case 10
    strIGNotes = cleanupString(cellValue)
Case 11
    strHasExercise = cleanupString(cellValue)
Case 12
    strExerciseTitle = cleanupString(cellValue)
Case 13
    strExerciseDescription = cleanupString(cellValue)
Case 14
    strMediaRequired = cleanupString(cellValue)
Case 15
    strMediaDetails = cleanupString(cellValue)
End Select
End If        'end if cellrow > 6
Next c

End Function
Private Function createSlide(strModule As String, strSubtitle As String, strDescription As String, strInstructor As String, strModuleDuration As String, strTopicTitle As String, strObjective As String, strSlideText As String, strPGNotes As String, strIGNotes As String, strHasExercise As String, strExerciseTitle As String, strExerciseDescription As String, strMediaRequired As String, strMediaDetails As String, oPresentation As Object)
    Dim sld    As Object
    Dim strPresenterNotes As String
    
    If StrComp(strMediaRequired, "") <> 0 Then        'some form of graphic so create different layout
    Set sld = oPresentation.slides.Add(oPresentation.slides.Count + 1, 29)
    sld.Comments.Add 12, 12, "TODO", "jmd", strMediaRequired & ": " & strMediaDetails
Else
    Set sld = oPresentation.slides.Add(oPresentation.slides.Count + 1, 2)
End If
sld.Select

sld.Shapes(1).TextFrame2.TextRange.Text = strTopicTitle
sld.Shapes(2).TextFrame2.TextRange.Text = strSlideText
sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Objective: " & strObjective & vbCrLf & vbCrLf & "Participant Notes: " & strPGNotes & vbCrLf & "##" & vbCrLf & "Presenter Notes: " & strIGNotes & vbCrLf & "##"

If StrComp(strHasExercise, "True") = 0 Then        'there is an exercise so create a placeholder slide
Set sld = oPresentation.slides.Add(oPresentation.slides.Count + 1, 33)
'sld.Select
sld.Shapes(1).TextFrame2.TextRange.Text = strExerciseTitle
sld.Shapes(2).TextFrame2.TextRange.Text = "Exercise"
sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "##" & vbCrLf & "Exercise Description: " & strExerciseDescription & vbCrLf & "##"
End If
'oPresentation.SectionProperties.AddBeforeSlide sld.slideindex, strModule

End Function

Private Function cleanupString(strText As String) As String
    strText = Replace(strText, "%0A", vbCrLf)
    strText = Replace(strText, "%2C", ",")
    strText = Replace(strText, "%2F", "/")
    
    cleanupString = strText
    
End Function

Private Function createTitleSlide(strModule As String, strSubtitle As String, strDescription As String, strInstructor As String, strModuleDuration As String, oPresentation As Object)
    Dim sld    As Object
    
    Set sld = oPresentation.slides.Add(oPresentation.slides.Count + 1, 1)
    sld.Select
    sld.Shapes(1).TextFrame2.TextRange.Text = strModule
    sld.Shapes(2).TextFrame2.TextRange.Text = strSubtitle
    oPresentation.SectionProperties.AddBeforeSlide sld.slideindex, strModule
    
End Function

Private Function createSectionTitleSlide(strModule As String, strSubtitle As String, strDescription As String, strInstructor As String, strModuleDuration As String, oPresentation As Object)
    Dim sld    As Object
    'strDescription = cleanupString(strDescription)
    'strModule = cleanupString(strModule)
    'strSubtitle = cleanupString(strSubtitle)
    
    Set sld = oPresentation.slides.Add(oPresentation.slides.Count + 1, 1)
    sld.Shapes(1).TextFrame2.TextRange.Text = strModule
    sld.Shapes(2).TextFrame2.TextRange.Text = strSubtitle
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Module Description: " & strDescription & vbCrLf & strModuleDuration & " Minutes" & vbCrLf
    oPresentation.SectionProperties.AddBeforeSlide sld.slideindex, strModule
    
End Function

Private Function createNewPPT() As Object
    
    'Step 1: Declare your variables
    Dim objPPT As Object
    Dim oPresentation As Object
    Dim sld    As Object
    Dim SlideTitle As String
    'Step 2: Open PowerPoint and create new presentation
    Set objPPT = CreateObject("PowerPoint.Application")
    Set oPresentation = objPPT.Presentations.Add
    objPPT.Visible = True
    Set createNewPPT = oPresentation
    'Step 7: Memory Cleanup
    objPPT.Activate
    ' Set sld = Nothing
    ' Set oPresentation = Nothing
    ' Set objPPT = Nothing
    
End Function
