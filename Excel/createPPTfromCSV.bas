Option Explicit

Public Sub Create_PPT_From_CSV()
    Dim oPresentation As Object
    Set oPresentation = createNewPPT()
    Call LoopCells(oPresentation)
    oPresentation.slides(1).Select
    Set oPresentation = Nothing
    'MsgBox ("All Done")
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
    Dim strCourseObjectives As String
    
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
    Dim strModuleObjectives As String
    Dim strMediaDetails As String
    Dim strFileName As String
    Dim strListedTopics As String
    
    Dim intFirstRowOfData As Integer
    intFirstRowOfData = 8        ' update this number if more course details are added later
    Dim intPreviousRow As Integer
    intPreviousRow = intFirstRowOfData        'set to first row of data
    Dim strModulePrevious
    Dim intModuleNumber As Integer
    intModuleNumber = 1
    
    Set myRange = ActiveSheet.usedRange 'range with data in int
    For Each c In myRange 'each cell
        
        ' c.Select
        cellValue = c.Value 'value of cell
        cellAddress = c.Address 'address
        cellColumn = c.Column 'column number
        cellRow = c.Row 'row number
        
        If cellColumn = 2 And cellRow = 1 Then        ' course title
        strCourseTitle = cleanupString(cellValue)
        oPresentation.BuiltinDocumentProperties(1).Value = strCourseTitle
    End If        'end if course title
    If cellColumn = 2 And cellRow = 2 Then        'course client
    strClient = cleanupString(cellValue)
    Call createSlide2(oPresentation, 1, strCourseTitle, strClient, , , , , strCourseTitle) 'title slide for course title
    
End If        ' end if course client

If cellColumn = 2 And cellRow = 4 Then        ' course objectives
strCourseObjectives = cleanupString(cellValue)
Call createObjectivesSlide(2, "Course Objectives", strCourseObjectives, oPresentation)

End If

If cellRow >= intFirstRowOfData Then        ' get content
Select Case cellColumn
    Case 1
        strModule = cleanupString(cellValue)
        If intPreviousRow <> cellRow Then        'a new row and not the first so create slide from assembled stuff
        Call regularSlide(strModule, strSubtitle, strDescription, strInstructor, strModuleDuration, strTopicTitle, strObjective, strSlideText, strPGNotes, strIGNotes, strHasExercise, strExerciseTitle, strExerciseDescription, strMediaRequired, strMediaDetails, strFileName, oPresentation)
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
    
Case 6
    strTopicTitle = cleanupString(cellValue)
    strListedTopics = strListedTopics & strTopicTitle & ", "
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
Case 16
    
    If StrComp(strModulePrevious, strModule) <> 0 Then        'module title is different or first so create review slide, section and title slide
    Call createSlide2(oPresentation, 33, "Review", "Questions?", "Module Objectives: " & vbCrLf & strModuleObjectives & _
    vbCrLf & "Topics Covered: " & vbCrLf & strListedTopics) 'creating review slide
    strListedTopics = ""
    
    strModuleObjectives = cleanupString(cellValue) 'Set module obectives after creating review slide above
    
    Call createSlide2(oPresentation, 33, "Module " & intModuleNumber & ": " & strModule, strSubtitle, "Module Description: " & _
    strDescription & vbCrLf & "Module Duration: " & strModuleDuration & " Minutes" & vbCrLf & "Module Objectives: " & vbCrLf & _
    strModuleObjectives, , , , "Module " & intModuleNumber & ": " & strModule) ' create section and title slide TODO if next module title includes previous reivew then it would happen here
    
    intModuleNumber = intModuleNumber + 1
    strModulePrevious = cleanupString(strModule)
    
End If

Case 17
    strFileName = cellValue
End Select
End If        'end if cellrow > intfirstrowof data
Next c

End Function

Private Function createObjectivesSlide(intLayout As Integer, strSlideTitle As String, strSlideText As String, oPresentation As Object)
    Dim sld    As Object
    
    Set sld = createSlide2(oPresentation, 2, "Course Objectives", strSlideText)
    
    sld.Shapes(2).TextFrame.TextRange.ParagraphFormat.Bullet.Type = 2        'sets to numbered list
End Function
Private Function regularSlide(strModule As String, strSubtitle As String, strDescription As String, strInstructor As String, strModuleDuration As String, strTopicTitle As String, strObjective As String, strSlideText As String, strPGNotes As String, strIGNotes As String, strHasExercise As String, strExerciseTitle As String, strExerciseDescription As String, strMediaRequired As String, strMediaDetails As String, strFileName As String, oPresentation As Object)
    Dim sld    As Object
    Dim strPresenterNotes As String
    Dim shpInserted As Object
    Dim strNotes As String
    strNotes = "Objective: " & strObjective & vbCrLf & _
        "##" & vbCrLf & "Presenter Notes: " & strIGNotes & vbCrLf & "##" & vbCrLf & _
        "Participant Notes: " & vbCrLf & strPGNotes & vbCrLf
    
    If StrComp(strMediaRequired, "") <> 0 Then        'some form of graphic so create different layout
    Set sld = createSlide2(oPresentation, 29, strTopicTitle, strSlideText, strNotes, "TODO", "jmd", strMediaRequired & ": " & strMediaDetails)
     
     If strFileName <> "" Then 'there is a file
     Set shpInserted = addImage(strFileName, sld)
     End If
     
Else
    Set sld = createSlide2(oPresentation, 2, strTopicTitle, strSlideText, strNotes)
   
End If

If StrComp(strHasExercise, "True") = 0 Then        'there is an exercise so create a placeholder slide
'Set sld = createSlide2(oPresentation, 33, strExerciseTitle, "Exercise", "Objective: " & strObjective & vbCrLf & "##" & vbCrLf & "Exercise Description: " & strExerciseDescription & vbCrLf & "##")

Set sld = createSlide2(oPresentation, 35, strExerciseTitle, strExerciseDescription, "Objective: " & _
strObjective)
sld.Shapes(3).TextFrame2.TextRange.Text = "Exercise"
End If

End Function

Private Function cleanupString(strText As String) As String
    strText = Replace(strText, "%0A", vbCrLf)
    strText = Replace(strText, "%2C", ",")
    strText = Replace(strText, "%2F", "/")
    cleanupString = strText
    
End Function

Function createSlide2(oPresentation As Object, Optional intLayout As Integer = 2, Optional strSlideTitle As String = "", Optional strSlideText As String = "", _
         Optional strPresenterNotes As String = "", Optional strCommentTitle As String = "", Optional strCommentInitials As String = "", _
         Optional strCommentText As String = "", Optional strSectionTitle As String) As Object
    
    Dim sld    As Object
    
    Set sld = oPresentation.slides.Add(oPresentation.slides.Count + 1, intLayout)
    sld.Select
    
    If strCommentText <> "" Then        'create comment
    sld.Comments.Add 12, 12, strCommentTitle, strCommentInitials, strCommentText
End If

If sld.Shapes.Count > 0 And strSlideTitle <> "" Then        'hopefully title
sld.Shapes(1).TextFrame2.TextRange.Text = strSlideTitle
End If

If sld.Shapes.Count > 1 And strSlideText <> "" Then        'hopefully title
sld.Shapes(2).TextFrame2.TextRange.Text = strSlideText
End If

If strSectionTitle <> "" Then        ' create section
oPresentation.SectionProperties.AddBeforeSlide sld.slideindex, strSectionTitle
End If

If strPresenterNotes <> "" Then        'create notes
sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = strPresenterNotes
End If

Set createSlide2 = sld        'return this created slide in case

End Function

Private Function createNewPPT() As Object
    
    Dim objPPT As Object
    Dim oPresentation As Object
    Dim sld    As Object
    Dim SlideTitle As String
    Set objPPT = CreateObject("PowerPoint.Application")
    Set oPresentation = objPPT.Presentations.Add
    'slideheight 540 slidewidth 960 default (slidesize: 7 ppSlideSizeCustom) but changed to 4:3 the width is now 720 SlideSize: ppSlideSizeOnScreen (1)
    'With oPresentation.PageSetup
    If (MsgBox("Use 4:3 ratio? Otherwise it will be 16:9?", (vbYesNo + vbQuestion), "Slide Size?") = vbYes) Then
        oPresentation.PageSetup.SlideSize = 1
    End If
    
    'End With
    objPPT.Visible = True
    Set createNewPPT = oPresentation
    objPPT.Activate
    Debug.Print "here: "; oPresentation.PageSetup.SlideSize
    
End Function

Private Function addImage(strFileName As String, sld As Object) As Object
Dim strPath As String
Dim shpInserted As Object

  strPath = ActiveWorkbook.Path & "\" & strFileName
  'strPath = Environ("USERPROFILE") & "\Downloads\" & strFileName
  Debug.Print strPath
  On Error GoTo Errorhandler
    Set shpInserted = sld.Shapes.AddPicture(Filename:=strPath, LinkToFile:=False, SaveWithDocument:=msoTrue, Left:=0, Top:=0)
    
    Exit Function
Errorhandler:
    sld.Comments.Add 12, 12, "TODO", "jmd", "File Not Found, Image not added: " & strFileName
    Resume Next

End Function
