Option Explicit
Public Sub Create_PPT_From_CSV()
    Dim oPresentation As Object
    Set oPresentation = createNewPPT()
    Call LoopCells(oPresentation)
    oPresentation.slides(1).Select
    Set oPresentation = Nothing
End Sub
Private Function LoopCells(oPresentation As Object)
    Dim headersColl As New Collection
    Dim rowColl As New Collection
    Dim strReviewQuestions As String
    Dim currentCell      As Range
    Dim strListedTopics As String
    Dim intPreviousModuleNumber As Integer:   intPreviousModuleNumber = 0
    For Each currentCell In ActiveSheet.usedRange
        'currentCell.Select                        'TODO debug and delete
        
        
        If currentCell.Column = 1 And currentCell.Row > 2 Then ' START OF A NEW ROW AND NOT HEADER ROW OR BEFORE FIRST ROW OF DATA
            If currentCell.Row = 3 Then           ' IS END OF FIRST ROW OF DATA CREATE TITLE SLIDE AFTER FIRST ROW OF DATA
                Call createCourseTitleSlide(rowColl, oPresentation)
                oPresentation.BuiltinDocumentProperties(1).Value = rowColl("Course Title") 'create course title slide
                Call createCourseObjectivesSlide(rowColl, oPresentation)
            End If
            
            If intPreviousModuleNumber <> rowColl("Module Number") Then ' CREATE NEW SECTION AND MODULE TITLE SLIDE, CREATE REVIEW OF PREVIOUS IF NOT THE FIRST
                If intPreviousModuleNumber > 0 Then ' IS NOT THE FIRST SO CREATE REVIEW SLIDE
                    Call createModuleReviewSlide(rowColl, strListedTopics, strReviewQuestions, oPresentation)
                    strListedTopics = ""          ' reset topic list which appears on review slide
                    strReviewQuestions = ""       ' reset review questions
                End If
                
                Call createModuleSectionAndTitleSlide(rowColl, oPresentation)
                intPreviousModuleNumber = rowColl("Module Number")
            End If
            Call createRegularSlides(rowColl, oPresentation)
            Set rowColl = New Collection
            oPresentation.slides(oPresentation.slides.Count).Select
        End If
        
        If currentCell.Row > 1 Then               ' get content for all non header content
            Select Case currentCell.Column
                Case headersColl("Course Title")
                    rowColl.Add cleanupString(currentCell.Value), "Course Title"
                Case headersColl("Course Client")
                    rowColl.Add cleanupString(currentCell.Value), "Course Client"
                Case headersColl("Course Title")
                    rowColl.Add cleanupString(currentCell.Value), "Course Title"
                Case headersColl("Course Client")
                    rowColl.Add cleanupString(currentCell.Value), "Course Client"
                Case headersColl("Course Duration (days)")
                    rowColl.Add cleanupString(currentCell.Value), "Course Duration (days)"
                Case headersColl("Course Objectives")
                    rowColl.Add cleanupString(currentCell.Value), "Course Objectives"
                Case headersColl("Module Number")
                    rowColl.Add cleanupString(currentCell.Value), "Module Number"
                Case headersColl("Module Title")
                    rowColl.Add cleanupString(currentCell.Value), "Module Title"
                Case headersColl("Module Subtitle")
                    rowColl.Add cleanupString(currentCell.Value), "Module Subtitle"
                Case headersColl("Module Description")
                    rowColl.Add cleanupString(currentCell.Value), "Module Description"
                Case headersColl("Module Instructor")
                    rowColl.Add cleanupString(currentCell.Value), "Module Instructor"
                Case headersColl("Module Duration")
                    rowColl.Add cleanupString(currentCell.Value), "Module Duration"
                Case headersColl("Module Objectives")
                    rowColl.Add cleanupString(currentCell.Value), "Module Objectives"
                Case headersColl("Topic Number")
                    rowColl.Add cleanupString(currentCell.Value), "Topic Number"
                Case headersColl("Topic Title")
                    rowColl.Add cleanupString(currentCell.Value), "Topic Title"
                    If strListedTopics <> "" Then
                        strListedTopics = strListedTopics & ", "
                    End If
                    strListedTopics = strListedTopics & rowColl("Topic Title")
                Case headersColl("Topic Objective")
                    rowColl.Add cleanupString(currentCell.Value), "Topic Objective"
                Case headersColl("Topic Slide Text")
                    rowColl.Add cleanupString(currentCell.Value), "Topic Slide Text"
                Case headersColl("Topic PG Notes")
                    rowColl.Add cleanupString(currentCell.Value), "Topic PG Notes"
                Case headersColl("Topic IG Notes")
                    rowColl.Add cleanupString(currentCell.Value), "Topic IG Notes"
                Case headersColl("Topic Review Questions")
                    rowColl.Add cleanupString(currentCell.Value), "Topic Review Questions"
                    If strReviewQuestions <> "" Then
                        strReviewQuestions = strReviewQuestions & vbCrLf
                    End If
                    strReviewQuestions = strReviewQuestions & rowColl("Topic Review Questions")
                Case headersColl("Topic Has Exercise")
                    rowColl.Add cleanupString(currentCell.Value), "Topic Has Exercise"
                Case headersColl("Topic Exercise Title")
                    rowColl.Add cleanupString(currentCell.Value), "Topic Exercise Title"
                Case headersColl("Topic Exercise Description")
                    rowColl.Add cleanupString(currentCell.Value), "Topic Exercise Description"
                Case headersColl("Topic Media Required")
                    rowColl.Add cleanupString(currentCell.Value), "Topic Media Required"
                Case headersColl("Topic Media Details")
                    rowColl.Add cleanupString(currentCell.Value), "Topic Media Details"
                Case headersColl("Topic Media File Name")
                    rowColl.Add cleanupString(currentCell.Value), "Topic Media File Name"
                Case headersColl("Topic Media Alt Text")
                    rowColl.Add cleanupString(currentCell.Value), "Topic Media Alt Text"
            End Select
        ElseIf currentCell.Row = 1 Then           ' GET HEADER INFO FOR EVERY CELL OF THIS ROW
            headersColl.Add currentCell.Column, currentCell.Value
        End If
    Next currentCell
    'TODO assembled course objectives would go here
End Function

Private Function createCourseTitleSlide(rowColl As Collection, oPresentation As Object)
    Dim sld    As Object
    Set sld = oPresentation.slides.Add(oPresentation.slides.Count + 1, 1)
    sld.Shapes(1).TextFrame2.TextRange.Text = rowColl("Course Title")
    sld.Shapes(2).TextFrame2.TextRange.Text = rowColl("Course Client")
    oPresentation.SectionProperties.AddBeforeSlide sld.slideindex, rowColl("Course Title")
End Function

Private Function createCourseObjectivesSlide(rowColl As Collection, oPresentation As Object)
    Dim sld    As Object
    Set sld = oPresentation.slides.Add(oPresentation.slides.Count + 1, 2)
    sld.Shapes(1).TextFrame2.TextRange.Text = "Course Objectives"
    sld.Shapes(2).TextFrame2.TextRange.Text = rowColl("Course Objectives")
    sld.Shapes(2).TextFrame.TextRange.ParagraphFormat.Bullet.Type = 2 'sets to numbered list
End Function

Private Function createModuleSectionAndTitleSlide(rowColl As Collection, oPresentation As Object)
    Dim sld    As Object
    Set sld = oPresentation.slides.Add(oPresentation.slides.Count + 1, 33)
    sld.Shapes(1).TextFrame2.TextRange.Text = "Module " & rowColl("Module Number") & ": " & rowColl("Module Title")
    sld.Shapes(2).TextFrame2.TextRange.Text = rowColl("Module Subtitle")
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Module Description: " & rowColl("Module Description") & vbCrLf & "Module Duration: " & _
                                                        rowColl("Module Duration") & " Minutes" & vbCrLf & "Module Objectives: " & vbCrLf & rowColl("Module Objectives")
    oPresentation.SectionProperties.AddBeforeSlide sld.slideindex, "Module " & rowColl("Module Number") & ": " & rowColl("Module Title")
End Function

Private Function createModuleReviewSlide(rowColl As Collection, strListedTopics As String, strReviewQuestions As String, oPresentation As Object)
    Dim sld    As Object
    Set sld = oPresentation.slides.Add(oPresentation.slides.Count + 1, 33)
    sld.Shapes(1).TextFrame2.TextRange.Text = "Review" 'Title
    sld.Shapes(2).TextFrame2.TextRange.Text = rowColl("Module Objectives") 'Subtitle
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Review Questions: " & vbCrLf & strReviewQuestions & vbCrLf & "Module Objectives: " & vbCrLf _
    & rowColl("Module Objectives") & vbCrLf & "Topics Covered: " & vbCrLf & strListedTopics
End Function

Private Function createRegularSlides(rowColl As Collection, oPresentation As Object)
    Dim sld    As Object
    Dim strNotes As String
    strNotes = "Objective " & rowColl("Topic Objective") & vbCrLf & "##" & vbCrLf & "Presenter Notes: " & rowColl("Topic IG Notes") & vbCrLf & "##" & vbCrLf & _
               "Participant Notes: " & vbCrLf & rowColl("Topic PG Notes") & vbCrLf
    
    If StrComp(rowColl("Topic Media Required"), "") <> 0 Then 'some form of graphic so create different layout
        If InStr(rowColl("Topic Media Required"), "/") > 0 Then 'Multiple graphics
            Call createSlideWithMultipleGraphics(rowColl, strNotes, oPresentation)
        ElseIf InStr(rowColl("Topic Media Required"), "graphic") > 0 Then ' just graphic
            Call createSlideWithGraphic(rowColl, strNotes, oPresentation)
        ElseIf InStr(rowColl("Topic Media Required"), "animation") > 0 Then '
            Call createSlideWithAnimation(rowColl, strNotes, oPresentation)
        Else
            Call createSlideWithVideo(rowColl, strNotes, oPresentation)
        End If
    Else
        Call createTextSlideWithoutGraphic(rowColl, strNotes, oPresentation)
    End If
    If StrComp(rowColl("Topic Has Exercise"), "True") = 0 Then 'there is an exercise so create a placeholder slide
        Call createExerciseSlide(rowColl, oPresentation)
    End If
    
End Function

Private Function createSlideWithMultipleGraphics(rowColl As Collection, strNotes As String, oPresentation As Object)
    Dim shpInserted As Object
    Dim sld    As Object
    Set sld = oPresentation.slides.Add(oPresentation.slides.Count + 1, 22)
    sld.Shapes(1).TextFrame2.TextRange.Text = rowColl("Topic Title") 'Title
    sld.Shapes(4).TextFrame2.TextRange.Text = rowColl("Topic Slide Text") 'Content
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = strNotes
    sld.Comments.Add 12, 12, "TODO", "jmd", rowColl("Topic Media Required") & ": " & rowColl("Topic Media Details")
    If rowColl("Topic Media File Name") <> "" Then 'there is a file
        Set shpInserted = addImage(rowColl("Topic Media File Name"), rowColl("Topic Media Alt Text"), sld)
    End If
    
End Function

Private Function createSlideWithGraphic(rowColl As Collection, strNotes As String, oPresentation As Object)
    Dim shpInserted As Object
    Dim sld    As Object
    Set sld = oPresentation.slides.Add(oPresentation.slides.Count + 1, 36)
    sld.Shapes(1).TextFrame2.TextRange.Text = rowColl("Topic Title") 'Title
    sld.Shapes(3).TextFrame2.TextRange.Text = rowColl("Topic Slide Text") 'Content
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = strNotes
    sld.Comments.Add 12, 12, "TODO", "jmd", rowColl("Topic Media Required") & ": " & rowColl("Topic Media Details")
    If rowColl("Topic Media File Name") <> "" Then 'there is a file
        Set shpInserted = addImage(rowColl("Topic Media File Name"), rowColl("Topic Media Alt Text"), sld)
    End If
    
End Function

Private Function createSlideWithVideo(rowColl As Collection, strNotes As String, oPresentation As Object)
    Dim shpInserted As Object
    Dim sld    As Object
    Set sld = oPresentation.slides.Add(oPresentation.slides.Count + 1, 18)
    sld.Shapes(1).TextFrame2.TextRange.Text = rowColl("Topic Title") 'Title
    sld.Shapes(3).TextFrame2.TextRange.Text = rowColl("Topic Slide Text") 'Content
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = strNotes
    sld.Comments.Add 12, 12, "TODO", "jmd", rowColl("Topic Media Required") & ": " & rowColl("Topic Media Details")
    
End Function

Private Function createSlideWithAnimation(rowColl As Collection, strNotes As String, oPresentation As Object)
    Dim shpInserted As Object
    Dim sld    As Object
    Set sld = oPresentation.slides.Add(oPresentation.slides.Count + 1, 17)
    sld.Shapes(1).TextFrame2.TextRange.Text = rowColl("Topic Title") 'Title
    sld.Shapes(2).TextFrame2.TextRange.Text = rowColl("Topic Slide Text") 'Content
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = strNotes
    sld.Comments.Add 12, 12, "TODO", "jmd", rowColl("Topic Media Required") & ": " & rowColl("Topic Media Details")
    
End Function

Private Function createTextSlideWithoutGraphic(rowColl As Collection, strNotes As String, oPresentation As Object)
    Dim sld    As Object
    Set sld = oPresentation.slides.Add(oPresentation.slides.Count + 1, 2)
    sld.Shapes(1).TextFrame2.TextRange.Text = rowColl("Topic Title") 'Title
    sld.Shapes(2).TextFrame2.TextRange.Text = rowColl("Topic Slide Text") 'Content
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = strNotes
End Function

Private Function createExerciseSlide(rowColl As Collection, oPresentation As Object)
    Dim sld    As Object
    Set sld = oPresentation.slides.Add(oPresentation.slides.Count + 1, 35)
    sld.Shapes(1).TextFrame2.TextRange.Text = rowColl("Topic Exercise Title") 'Title
    sld.Shapes(2).TextFrame2.TextRange.Text = rowColl("Topic Exercise Description") 'Content
    sld.Shapes(3).TextFrame2.TextRange.Text = "Exercise"
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Objective: " & rowColl("Topic Objective")
End Function

Private Function createNewPPT() As Object
    Dim objPPT As Object
    Dim oPresentation As Object
    Dim sld    As Object
    Dim SlideTitle As String
    Set objPPT = CreateObject("PowerPoint.Application")
    Set oPresentation = objPPT.Presentations.Add
    'If (MsgBox("Use 4:3 ratio? Otherwise it will be 16:9?", (vbYesNo + vbQuestion), "Slide Size?") = vbYes) Then
    '    oPresentation.PageSetup.SlideSize = 1
    'End If
    objPPT.Visible = True
    Set createNewPPT = oPresentation
    objPPT.Activate
End Function

Private Function cleanupString(strText As String) As String
    strText = Replace(strText, "%0A", vbCrLf)
    strText = Replace(strText, "%2C", ",")
    strText = Replace(strText, "%2F", "/")
    cleanupString = strText
End Function

Private Function addImage(strFileName As String, strAltText As String, sld As Object) As Object
    Dim strPath As String
    Dim shpInserted As Object
    strPath = ActiveWorkbook.Path & "\" & strFileName
    On Error GoTo Errorhandler
    Set shpInserted = sld.Shapes.AddPicture(Filename:=strPath, LinkToFile:=False, SaveWithDocument:=msoTrue, Left:=0, Top:=0)
    If strAltText <> "" Then
        shpInserted.AlternativeText = strAltText
    End If
    Exit Function
Errorhandler:
    sld.Comments.Add 12, 12, "TODO", "jmd", "File Not Found, Image Not added: " & strFileName
    Resume Next
End Function
