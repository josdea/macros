Option Explicit

Sub Text_Go_To_Small_Text()                       ' checked 2/25/20
    'Go to the next slide that has text smaller than specified
    Dim sld                                       As Slide
    Dim shp                                       As Shape
    Dim currentSlideNumber                        As Integer
    Dim startingSlideNumber                       As Integer
    Dim fontSize As Integer
    fontSize = InputBox("Input font size to find text smaller than", "Font Size Smaller Than", "14")
    startingSlideNumber = Application.ActiveWindow.View.Slide.SlideNumber
    For currentSlideNumber = startingSlideNumber To ActivePresentation.Slides.count ' iterate slides
        Set sld = Application.ActivePresentation.Slides(currentSlideNumber)
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.TextRange.Font.Size < fontSize And shp.TextFrame.HasText = msoTrue Then
                    ActiveWindow.View.goToSlide sld.SlideIndex
                    shp.Select
                    If MsgBox("Small text found On slide " & currentSlideNumber & " Shape: " & shp.Name & ". Fix and set to " & fontSize & "?", (vbYesNo + vbQuestion), "Set to size " & fontSize & "?") = vbYes Then
                        shp.TextFrame.TextRange.Font.Size = fontSize
                        If shp.TextFrame.AutoSize <> ppAutoSizeShapeToFitText Or MsgBox("Shape does not autosize, do you want shape to auto scale?", (vbYesNo + vbQuestion), "Auto Size Shape?") = vbYes Then
                            shp.TextFrame.AutoSize = ppAutoSizeShapeToFitText
                        End If
                    End If
                    Exit Sub
                End If
            End If
        Next shp                                  ' end of iterate shapes
    Next                                          ' end of iterate slides
    MsgBox "No More Videos Found. Move To slide 1 To search again."
End Sub

Sub Text_Remove_Empty_Lines()
    Dim currentPresentation As Presentation: Set currentPresentation = ActivePresentation
    Dim sld    As Slide
    Dim shp    As Shape
    Dim para As TextRange
    Dim ln As TextRange
    
    For Each sld In currentPresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
            For Each para In shp.TextFrame.TextRange.Paragraphs
            Debug.Print para
            
            For Each ln In para.Lines
            Debug.Print ln
            Next ln
            Next para
             '   With shp.TextFrame.TextRange
                'MsgBox .Text
                   ' .Text = removeMultiBlank(.Text)
             '   End With
            End If
        Next shp
        If sld.HasNotesPage Then
            For Each shp In sld.NotesPage.Shapes
                If shp.HasTextFrame Then
                    With shp.TextFrame.TextRange
                   ' MsgBox .Text
                        .Text = removeMultiBlank(.Text)
                    End With
                End If
            Next shp
        End If
    Next sld
    
    
End Sub

Function removeMultiBlank(s As String) As String
    With CreateObject("VBScript.RegExp")
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = "^\s"
        
        removeMultiBlank = .Replace(s, "")
    End With
End Function

Sub Shapes_Delete_Empty_TextBoxes()

    Dim sld    As Slide
    Dim shp    As Shape
    Dim ShapeIndex As Integer

    For Each sld In ActivePresentation.Slides
        
        For ShapeIndex = sld.Shapes.count To 1 Step -1
            
            If sld.Shapes(ShapeIndex).Type = msoTextBox And Not sld.Shapes(ShapeIndex).TextFrame.HasText Then
                sld.Shapes(ShapeIndex).Delete
            End If
        Next
    Next sld
End Sub

Sub Create_Progress_Bar()
    Dim intSlideNumber As Integer
    Dim s      As Shape
    Dim intLineHeight As Integer
    intLineHeight = 3
    Dim oPres  As Presentation
    Set oPres = ActivePresentation
    On Error Resume Next
    With oPres
        For intSlideNumber = 2 To .Slides.count
            .Slides(intSlideNumber).Shapes("Progress_Bar").Delete
            Set s = .Slides(intSlideNumber).Shapes.AddLine(Beginx:=0, BeginY:=.PageSetup.SlideHeight - (intLineHeight / 2), _
                Endx:=intSlideNumber * .PageSetup.SlideWidth / .Slides.count, EndY:=.PageSetup.SlideHeight - (intLineHeight / 2))
            s.line.Weight = intLineHeight
            s.line.BackColor.ObjectThemeColor = msoThemeColorAccent1
            s.line.ForeColor.ObjectThemeColor = msoThemeColorAccent1
            s.Name = "Progress_Bar"
        Next intSlideNumber:
    End With
    MsgBox "All Done. Created a progress bar on " & oPres.Slides.count - 1 & " slides. Theme color accent 1 was used."
End Sub

Sub Video_Convert_Linked_To_Embedded()
    'Converts all linked videos to embedded
    MsgBox "This can take a few minutes,        't worry. Select ok to continue"
    Dim shp    As Shape
    Dim sld    As Slide
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            On Error Resume Next
            shp.LinkFormat.BreakLink
            On Error GoTo 0
        Next shp
    Next sld
    MsgBox "All Done"
End Sub

Sub Text_Language_Toggle_US_UK()                  ' checked 2/17/20
    'Toggles all text in all shapes on all slides and master for spell checking
    Dim currentPresentation As Presentation: Set currentPresentation = ActivePresentation
    Dim currentLanguage As Integer
    Dim totalShapeCount As Integer: totalShapeCount = 0

    Dim strLangSelect As String
    strLangSelect = ""
    
    Dim langList As New Collection
    Dim collMsoLang As New Collection
    langList.Add "English US", "1"
    collMsoLang.Add "1033", "1"
    langList.Add "English UK", "2"
    collMsoLang.Add "2057", "2"
    langList.Add "Arabic", "3"
    collMsoLang.Add "1025", "3"
    langList.Add "Spanish (General)", "4"
    collMsoLang.Add "1034", "4"
    langList.Add "French", "5"
    collMsoLang.Add "1036", "5"
    langList.Add "Russian", "6"
    collMsoLang.Add "1049", "6"
    langList.Add "Polish", "7"
    collMsoLang.Add "1045", "7"
    langList.Add "Romanian", "8"
    collMsoLang.Add "1048", "8"
    
    Dim lang   As Variant
    Dim i      As Integer
    i = 0
    For Each lang In langList
        i = i + 1
        strLangSelect = strLangSelect & ""
        strLangSelect = strLangSelect & i & ". " & lang & vbCrLf
    Next lang
    currentLanguage = collMsoLang(getNumberInput(strLangSelect, "Select from Language", 1, 1, i))
    
    Dim boolUpdateNotesLang As Boolean
    If MsgBox("Do you want to update the notes language as well?", vbYesNo) = vbYes Then
        boolUpdateNotesLang = True
    Else
        boolUpdateNotesLang = False
    End If
    
    Dim sld    As Slide
    Dim shp    As Shape
    For Each sld In currentPresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                shp.TextFrame.TextRange.LanguageID = currentLanguage
                totalShapeCount = totalShapeCount + 1
            End If
        Next shp
        If sld.HasNotesPage And boolUpdateNotesLang = True Then
            For Each shp In sld.NotesPage.Shapes
                If shp.HasTextFrame Then
                    shp.TextFrame.TextRange.LanguageID = currentLanguage
                    totalShapeCount = totalShapeCount + 1
                End If
            Next shp
        End If
    Next sld
    For Each shp In currentPresentation.SlideMaster.Shapes
        If shp.HasTextFrame Then
            shp.TextFrame.TextRange.LanguageID = currentLanguage
            totalShapeCount = totalShapeCount + 1
        End If
    Next shp
    Dim layCustom As CustomLayout
    
    For Each layCustom In currentPresentation.SlideMaster.CustomLayouts
        For Each shp In layCustom.Shapes
            If shp.HasTextFrame Then
                shp.TextFrame.TextRange.LanguageID = currentLanguage
                totalShapeCount = totalShapeCount + 1
            End If
        Next shp
    Next layCustom
    
    MsgBox ("All Done. Total Shapes Set: " & totalShapeCount & ". Press F7 To rerun spellcheck.")
End Sub

Function getNumberInput(strMessage As String, strBoxTitle As String, strDefaultValue As String, intMinNumber As Integer, intMaxNumber As Integer) As Double
    'This function is needed for the above
    Dim strInputField As String
    strMessage = strMessage & vbCrLf & vbCrLf & "Please enter a number between " & intMinNumber & " and " & intMaxNumber & ":"
    
    Do
        'Retrieve an answer from the user
        strInputField = InputBox(strMessage, strBoxTitle, strDefaultValue)
        If StrComp(strInputField, "x", 1) = 0 Then
            End
        ElseIf TypeName(strInputField) = "Boolean" Then 'Check if user selected cancel button
            getNumberInput = -1
        ElseIf Not IsNumeric(strInputField) Then  'input wasnt numeric
            getNumberInput = -1
        Else
            getNumberInput = strInputField        ' Number is numeric
        End If
    Loop While getNumberInput < intMinNumber Or getNumberInput > intMaxNumber Or getNumberInput < 0 'Keep prompting while out of range
    
End Function

Sub Sections_Bulk_Create()                        ' checked 1/17/20
    'Bulk creates sections with optional prefix and or suffix
    Dim sectionsDesired As Integer
    Dim sectionPrefix As String
    Dim sectioSuffix As String
    Dim inputText As String
    Dim inputAnswer As Variant
    Dim strPrefix As String
    Dim strSuffix As String
    inputText = "Please enter number of desired sections To create?"
    Do
        inputAnswer = InputBox(inputText)
        If TypeName(inputAnswer) = "Boolean" Then Exit Sub
    Loop While inputAnswer <= 0
    If inputAnswer > 0 Then
        sectionsDesired = inputAnswer
        strPrefix = InputBox("Enter Optional prefix." & vbNewLine & "(remember To add a space If you need one)", , "Module ")
        strSuffix = InputBox("Enter an Optional suffix." & vbNewLine & "(remember To add a space If you need one)")
        Dim i  As Integer
        For i = 1 To sectionsDesired
            ActivePresentation.SectionProperties.AddBeforeSlide 1, strPrefix & i & strSuffix
        Next i
    End If
End Sub

Sub File_PPTX_Combine_All_In_Folder()
    'Combines all PPTX files in the same folder as current file
    Dim vArray() As String
    Dim x      As Long
    Dim slideCountbeforeInsert As Integer
    slideCountbeforeInsert = ActivePresentation.Slides.count
    EnumerateFiles ActivePresentation.Path & "\", "*.PPTX", vArray
    If MsgBox("Are you sure you want To combine all " & UBound(vArray) & " PPTX files in the current folder? You should be in the first file And the order of the files will be alphabetical. This may require renaming them To 01, 02, 03 etc. If using numbers", (vbYesNo + vbQuestion), "Combine all?") = vbYes Then
        ActivePresentation.SectionProperties.AddBeforeSlide 1, "Module 1"
        With ActivePresentation
            For x = 1 To UBound(vArray)
                If Len(vArray(x)) > 0 Then
                    .Slides.InsertFromFile vArray(x), .Slides.count
                    ActivePresentation.SectionProperties.AddBeforeSlide slideCountbeforeInsert + 1, "Module " & x
                    slideCountbeforeInsert = ActivePresentation.Slides.count
                End If
            Next
            MsgBox "The " & UBound(vArray) & " files have been combined. There are now " & ActivePresentation.Slides.count & " total slides And " & ActivePresentation.SectionProperties.count & " total sections."
        End With
    Else
        MsgBox ("Action canceled.")
    End If
End Sub

Private Function EnumerateFiles(ByVal sDirectory As String, _
        ByVal sFileSpec As String, _
        ByRef vArray As Variant)
    Dim sTemp  As String
    ReDim vArray(1 To 1)
    sTemp = Dir$(sDirectory & sFileSpec)
    Do While Len(sTemp) > 0
        If sTemp <> ActivePresentation.Name Then
            ReDim Preserve vArray(1 To UBound(vArray) + 1)
            vArray(UBound(vArray)) = sDirectory & sTemp
        End If
        sTemp = Dir$
    Loop
End Function
Sub Shapes_Count_By_Name()                        ' checked 1/17/20
    'Count all shapes that have specified name
    Dim sld    As Slide
    Dim shp    As Shape
    Dim shapeNameToCount As String
    Dim shapesCounted As Integer
    shapeNameToCount = InputBox("What Is the shape name To count On all slides (case sensitive And notes Not counted")
    For Each sld In ActivePresentation.Slides     ' iterate slides
        For Each shp In sld.Shapes
            If shp.Name = shapeNameToCount Then
                shapesCounted = shapesCounted + 1
            End If
        Next shp
    Next sld                                      ' end of iterate slides
    MsgBox shapesCounted & " shapes match the name: " & shapeNameToCount
End Sub

Sub Presenter_Notes_Remove_All()                  ' checked 1/17/20
    'Deletes all presenter notes on all slides
    Dim oSl    As Slide
    Dim oSh    As Shape
    If MsgBox("Are you sure you want To delete all presenter/instructor notes?", (vbYesNo + vbQuestion), "Delete all Notes?") = vbYes Then
        For Each oSl In ActivePresentation.Slides
            For Each oSh In oSl.NotesPage.Shapes
                If oSh.Type = msoPlaceholder Then
                    If oSh.PlaceholderFormat.Type = ppPlaceholderBody Then
                        If oSh.HasTextFrame Then
                            oSh.TextFrame.TextRange.Text = ""
                        End If
                    End If
                End If
            Next oSh
        Next oSl
    Else
        MsgBox ("Action canceled.")
    End If
End Sub

Sub Shapes_Delete_By_Name()                       ' checked 1/17/20
    'Deletes all shapes that have specified name
    Dim sld    As Slide
    Dim shp    As Shape
    Dim shapeNameToDelete As String
    Dim shapesDeleted As Integer
    shapeNameToDelete = InputBox("What Is the shape name To delete On all slides?")
    For Each sld In ActivePresentation.Slides     ' iterate slides
        For Each shp In sld.Shapes
            If shp.Name = shapeNameToDelete Then
                shapesDeleted = shapesDeleted + 1
                shp.Delete
            End If
        Next shp
    Next sld                                      ' end of iterate slides
    MsgBox shapesDeleted & " shapes have been deleted"
End Sub

Sub Comments_Search_And_Export()                  ' checked 1/17/20
    'Exports all comments or based on search
    Dim replyCount As Integer
    Dim commentCount As Integer
    Dim sld    As Slide
    Dim myComment As Comment
    Dim Comment As String
    Dim commentSearch As String
    Dim reply  As String
    Dim replyIndex As Integer
    Dim tempComment As String
    commentSearch = InputBox("Enter Search term within comments Or leave blank For all")
    For Each sld In ActivePresentation.Slides
        For Each myComment In sld.Comments
            tempComment = "Slide " & sld.SlideNumber & " of " & ActivePresentation.Slides.count & vbCrLf
            tempComment = tempComment & "  " & myComment.Text & " (" & myComment.DateTime & ")" & vbCrLf
            replyCount = 0
            If myComment.Replies.count > 0 Then
                For replyIndex = 1 To myComment.Replies.count
                    tempComment = tempComment & "    \-" & myComment.Replies(replyIndex).Text & " (" & myComment.Replies(replyIndex).DateTime & ")" & vbCrLf
                    replyCount = replyCount + 1
                Next                              'next reply
            End If                                'there are replies
            If (InStr(UCase(tempComment), UCase(commentSearch))) Or commentSearch = "" Then
                Comment = Comment & tempComment
                Comment = Comment & vbCrLf
                commentCount = commentCount + 1 + replyCount
            End If                                'search matches or is blank
        Next myComment
    Next sld
    
    If commentCount > 0 Then                      ' there were comments
        Dim n      As Integer
        n = FreeFile()
        Open Environ("USERPROFILE") & "\Desktop\ppt_comments_" & Format(Now(), "yymmdd hhmm") & "-" & commentSearch & ".txt" For Output As #n
        Print #n, Comment                         ' write to file
        Close #n
        
        MsgBox commentCount & " comments and/or replies have been written To a Text file (ppt_comments) On your desktop"
    Else
        MsgBox "There were no comments found"
    End If
End Sub

Sub Image_Go_To_Next()                            ' checked 1/17/20
    'Go to the next slide that has an image
    Dim sld                                       As Slide
    Dim shp                                       As Shape
    Dim currentSlideNumber                        As Integer
    Dim startingSlideNumber                       As Integer
    Dim boolImageFound As Boolean
    boolImageFound = False
    
    
    startingSlideNumber = Application.ActiveWindow.View.Slide.SlideNumber + 1
    If startingSlideNumber > ActivePresentation.Slides.count Then
        startingSlideNumber = 1
    End If
    
    For currentSlideNumber = startingSlideNumber To ActivePresentation.Slides.count ' iterate slides
        Set sld = Application.ActivePresentation.Slides(currentSlideNumber)
        For Each shp In sld.Shapes
            If shp.Type = msoPicture Then
                boolImageFound = True
            End If
            
            If shp.Type = msoPlaceholder Then
                If shp.PlaceholderFormat.ContainedType = msoPicture Then
                    boolImageFound = True
                End If
                
            End If
            
            If boolImageFound = True Then
                boolImageFound = False
                ActiveWindow.View.goToSlide sld.SlideIndex
                shp.Select
                MsgBox "Slide Number: " & currentSlideNumber & vbCrLf & "Shape Name: " & shp.Name, vbOKOnly, "Image Found" ' outputs slide number and shape name
                If Not CommandBars.GetPressedMso("SelectionPane") Then CommandBars.ExecuteMso ("SelectionPane") ' show selection pane
                Exit Sub
            End If
            
        Next shp                                  ' end of iterate shapes
    Next                                          ' end of iterate slides
    MsgBox "No More Images Found."
End Sub

Sub Video_Go_To_Next()                            ' checked 1/17/20
    'Go to the next slide that has a video

    Dim sld                                       As Slide
    Dim shp                                       As Shape
    Dim currentSlideNumber                        As Integer
    Dim startingSlideNumber                       As Integer
    startingSlideNumber = Application.ActiveWindow.View.Slide.SlideNumber
    For currentSlideNumber = startingSlideNumber To ActivePresentation.Slides.count ' iterate slides
        Set sld = Application.ActivePresentation.Slides(currentSlideNumber)
        For Each shp In sld.Shapes
            If shp.Type = msoMedia Then
                If shp.MediaType = ppMediaTypeMovie Then
                    ActiveWindow.View.goToSlide sld.SlideIndex
                    MsgBox "Video found On slide " & currentSlideNumber & " Shape: " & shp.Name & ". Move To the Next slide And run again To continue searching" ' outputs slide number and shape name
                    If Not CommandBars.GetPressedMso("SelectionPane") Then CommandBars.ExecuteMso ("SelectionPane") ' show selection pane
                    Exit Sub                      ' end program for user to do things
                End If
            End If                                ' end of if type is msomedia
        Next shp                                  ' end of iterate shapes
    Next                                          ' end of iterate slides
    MsgBox "No More Videos Found. Move To slide 1 To search again."
End Sub

Sub Slide_Go_To()                                 ' checked 1/17/20
    'Go to a slide by number
    Dim slide_num As Integer
    Dim total_slides As Integer
    total_slides = ActivePresentation.Slides.count
    slide_num = InputBox("Enter slide number between 1 And " & total_slides, "Go To Slide")
    If ((slide_num <= 0) Or (slide_num > total_slides)) Then
        Slide_Go_To
    ElseIf (slide_num <= total_slides) Then
        ActiveWindow.View.goToSlide slide_num
    End If
End Sub

Sub PresenterNotes_Remove_Text_In_Hashtags()      ' checked 1/17/20
    'Removes all text wrapped in ##double hashtags## in presenter notes
    Dim stringBwDels As String, originalString As String, firstDelPos As Integer, secondDelPos As Integer, stringToReplace As String, replacedString As String
    Dim sld                                       As Slide ' declare slide object
    Dim shp                                       As Shape ' declare shape object
    For Each sld In ActivePresentation.Slides     ' iterate slides
    For Each shp In sld.NotesPage.Shapes          ' iterate note shapes
    If shp.Type = msoPlaceholder Then             ' check if its a placeholder
    If shp.PlaceholderFormat.Type = ppPlaceholderBody Then ' its presenter notes
    originalString = shp.TextFrame2.TextRange.Text
    stringToReplace = ""
    firstDelPos = InStr(originalString, "##") - 1 ' position of start delimiter
    secondDelPos = InStr(firstDelPos + 2, originalString, "##") ' position of end delimiter
    If secondDelPos <> 0 Then
        stringBwDels = Mid(originalString, firstDelPos + 1, secondDelPos - firstDelPos + 2) 'extract the string between two delimiters
    Else
        stringBwDels = 0
    End If
    replacedString = Replace(originalString, stringBwDels, stringToReplace)
    shp.TextFrame2.TextRange.Text = replacedString
End If
End If
Next shp                                          ' end of iterate shapes
Next sld                                          ' end of iterate slides
MsgBox "All Done"
End Sub

Sub PresenterNotes_Toggle_Visibility()            ' checked 1/17/20
    'Toggles visibility of all presenter notes on notes pages
    Dim toggleOn                                  As Boolean
    toggleOn = msoTrue
    If MsgBox("Do you want To hide all presenter note shapes On the notes pages. If you answer `no` then they will all be made visible.", (vbYesNo + vbQuestion), "Toggle Presenter Notes?") = vbYes Then
        toggleOn = msoFalse
    End If
    Dim sld                                       As Slide ' declare slide object
    Dim shp                                       As Shape ' declare shape object
    For Each sld In ActivePresentation.Slides     ' iterate slides
        For Each shp In sld.NotesPage.Shapes      ' iterate note shapes
            If shp.Type = msoPlaceholder Then     ' check if its a placeholder
                If shp.PlaceholderFormat.Type = ppPlaceholderBody Then ' its presenter notes
                    shp.Visible = toggleOn
                End If
            End If
        Next shp                                  ' end of iterate shapes
    Next sld                                      ' end of iterate slides
    If Not CommandBars.GetPressedMso("SelectionPane") Then CommandBars.ExecuteMso ("SelectionPane")
    MsgBox "All done toggling presenter notes"
    ActiveWindow.ViewType = ppViewNotesPage
End Sub

Sub Text_Font_Reset_To_Master()                   ' checked 1/17/20
    'Resets all text in all shapes and notes to master theme font
    Dim sld                                       As Slide
    Dim shp                                       As Shape
    Dim shapesAffected As Integer
    shapesAffected = shapesAffected + 1
    If MsgBox("Are you sure you want To reset all titles, text, And notes To the master font theme", (vbYesNo + vbQuestion), "Reset Font?") = vbYes Then
        For Each sld In ActivePresentation.Slides ' iterate slides
            For Each shp In sld.Shapes
                If shp.HasTextFrame Then
                    If shp.Type = msoPlaceholder Then
                        If shp.PlaceholderFormat.Type = ppPlaceholderTitle Or shp.PlaceholderFormat.Type = ppPlaceholderCenterTitle Then
                            shp.TextFrame.TextRange.Font.Name = "+mj-lt"
                            shapesAffected = shapesAffected + 1
                        Else
                            shp.TextFrame.TextRange.Font.Name = "+mn-lt"
                            shapesAffected = shapesAffected + 1
                        End If
                    Else
                        shp.TextFrame.TextRange.Font.Name = "+mn-lt"
                        shapesAffected = shapesAffected + 1
                    End If
                End If
            Next shp                              ' end of iterate shapes
            For Each shp In sld.NotesPage.Shapes
                If shp.HasTextFrame Then
                    If shp.Type = msoPlaceholder Then
                        If shp.PlaceholderFormat.Type = ppPlaceholderTitle Or shp.PlaceholderFormat.Type = ppPlaceholderCenterTitle Then
                            shp.TextFrame.TextRange.Font.Name = "+mj-lt"
                            shapesAffected = shapesAffected + 1
                        Else
                            shp.TextFrame.TextRange.Font.Name = "+mn-lt"
                            shapesAffected = shapesAffected + 1
                        End If
                    Else
                        shp.TextFrame.TextRange.Font.Name = "+mn-lt"
                        shapesAffected = shapesAffected + 1
                    End If
                End If
            Next shp
        Next sld                                  ' end of iterate slides
        MsgBox "All Done. " & shapesAffected & " shapes have been searched And reset"
    Else
        MsgBox ("Action canceled.")
    End If
End Sub

Sub Text_Remove_Double_Spaces()                   ' checked 1/17/20
    'Removes all text which has double spacebars and replaces with one
    Dim spacesRemoved As Integer
    Dim sld    As Slide
    Dim shp    As Shape
    Dim shpText As String
    spacesRemoved = 0
    Dim shapeCount As Integer
    shapeCount = 0
    If MsgBox("Do you want To replace all instances of multiple spaces With one space", (vbYesNo + vbQuestion), "Remove extra Spaces?") = vbYes Then
        For Each sld In ActivePresentation.Slides
            For Each shp In sld.Shapes
                shapeCount = shapeCount + 1
                If shp.HasTextFrame Then
                    shpText = shp.TextFrame.TextRange.Text 'Get the shape's text
                    Do While InStr(shpText, "  ") > 0
                        shpText = Trim(Replace(shpText, "  ", " "))
                        spacesRemoved = spacesRemoved + 1
                    Loop
                    shp.TextFrame.TextRange.Text = shpText 'Put the new text in the shape
                Else
                    shpText = vbNullString
                End If
            Next shp
            For Each shp In sld.NotesPage.Shapes
                shapeCount = shapeCount + 1
                If shp.HasTextFrame Then
                    shpText = shp.TextFrame.TextRange.Text 'Get the shape's text
                    Do While InStr(shpText, "  ") > 0
                        shpText = Trim(Replace(shpText, "  ", " "))
                        spacesRemoved = spacesRemoved + 1
                    Loop
                    shp.TextFrame.TextRange.Text = shpText 'Put the new text in the shape
                Else
                    shpText = vbNullString
                End If
            Next shp
        Next sld
    End If
    MsgBox spacesRemoved & " places where extra spacing was removed in " & shapeCount & " shapes."
End Sub

Sub Layout_Create_All_Types()                     ' checked 1/17/20
    'Creates an example slide of each default layout
    Dim sld    As Slide
    Dim layout As Integer

    'For layout = 1 To 36
    '  Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=layout)
    '    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: " & sld.layout
    'Next layout

    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutTitle)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutTitle (1) - Title"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutText (2) - Text"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutTwoColumnText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutTwoColumnText (3) - Two-column text"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutTable)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutTable (4) - Table"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutTextAndChart)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutTextAndChart (5) - Text and chart"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutChartAndText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutChartAndText (6) - Chart and text"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutOrgchart)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutOrgchart (7) - Organization chart"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutChart)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutChart (8) - Chart"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutTextAndClipart)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutTextAndClipArt (9) - Text and ClipArt"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutClipartAndText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutClipArtAndText (10) - ClipArt and text"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutTitleOnly)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutTitleOnly (11) - Title only"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutBlank)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutBlank (12) - Blank"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutTextAndObject)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutTextAndObject (13) - Text and object"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutObjectAndText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutObjectAndText (14) - Object and text"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutLargeObject)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutLargeObject (15) - Large object"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutObject)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutObject (16) - Object"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutTextAndMediaClip)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutTextAndMediaClip (17) - Text and MediaClip"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutMediaClipAndText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutMediaClipAndText (18) - MediaClip and text"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutObjectOverText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutObjectOverText (19) - Object over text"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutTextOverObject)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutTextOverObject (20) - Text over object"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutTextAndTwoObjects)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutTextAndTwoObjects (21) - Text and two objects"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutTwoObjectsAndText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutTwoObjectsAndText (22) - Two objects and text"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutTwoObjectsOverText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutTwoObjectsOverText (23) - Two objects over text"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutFourObjects)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutFourObjects (24) - Four objects"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutVerticalText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutVerticalText (25) - Vertical text"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutClipArtAndVerticalText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutClipArtAndVerticalText (26) - ClipArt and vertical text"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutVerticalTitleAndText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutVerticalTitleAndText (27) - Vertical title and text"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutVerticalTitleAndTextOverChart)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutVerticalTitleAndTextOverChart (28) - Vertical title and text over chart"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutTwoObjects)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutTwoObjects (29) - Two objects"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutObjectAndTwoObjects)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutObjectAndTwoObjects (30) - Object and two objects"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutTwoObjectsAndObject)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutTwoObjectsAndObject (31) - Two objects and object"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutCustom)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutCustom (32) - Custom"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutSectionHeader)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutSectionHeader (33) - Section header"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutComparison)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutComparison (34) - Comparison"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutContentWithCaption)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutContentWithCaption (35) - Content with caption"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutPictureWithCaption)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutPictureWithCaption (36) - Picture with caption"
    MsgBox "36 Slides have been created"
End Sub

Sub Image_Add_Comment_For_Missing_AltText()       ' checked 1/17/20
    'Add a comment for all slides with images having missing alternate text
    Dim sld    As Slide                           ' declare slide object
    Dim shp                                       As Shape ' declare shape object
    Dim intCommentCount As Integer
    intCommentCount = 0
    Dim boolImageFound As Boolean
    boolImageFound = False
    For Each sld In ActivePresentation.Slides     ' iterate slides
        For Each shp In sld.Shapes                ' iterate note shapes
            If shp.Type = msoPicture Then         ' check if its a placeholder
                boolImageFound = True
            End If
            
            If shp.Type = msoPlaceholder Then
                If shp.PlaceholderFormat.ContainedType = msoPicture Then
                    boolImageFound = True
                End If
            End If
            
            If boolImageFound = True Then
                boolImageFound = False
                
                If shp.AlternativeText = "" Or InStr(shp.AlternativeText, "generated") Then
                    sld.Comments.Add 12, 12, "Auto", "JMD", "TODO: Need Image ID Or source"
                    intCommentCount = intCommentCount + 1
                End If
            End If
            
        Next shp                                  ' end of iterate shapes
    Next sld                                      ' end of iterate slides
    MsgBox "All Done " & intCommentCount & " comments were added"
End Sub

Sub PresenterNotes_Add_Comment_If_Empty()         ' checked 1/17/20
    'Adds a comment on every slide with empty presenter notes
    Dim sld    As Slide                           ' declare slide object
    Dim shp                                       As Shape ' declare shape object
    Dim intCommentCount As Integer
    intCommentCount = 0
    For Each sld In ActivePresentation.Slides     ' iterate slides
        For Each shp In sld.NotesPage.Shapes      ' iterate note shapes
            If shp.Type = msoPlaceholder Then     ' check if its a placeholder
                If shp.PlaceholderFormat.Type = ppPlaceholderBody Then ' its presenter notes
                    If shp.TextFrame.HasText = False Then
                        sld.Comments.Add 12, 12, "Auto", "JMD", "TODO: Need Presenter Notes"
                        intCommentCount = intCommentCount + 1
                    ElseIf shp.TextFrame.TextRange.Text = "" Then
                        sld.Comments.Add 12, 12, "Auto", "JMD", "TODO: Need Presenter Notes"
                        intCommentCount = intCommentCount + 1
                    End If
                End If
            End If
        Next shp                                  ' end of iterate shapes
    Next sld                                      ' end of iterate slides
    MsgBox "All Done " & intCommentCount & " comments were added"
End Sub

Sub Shapes_Go_To_Next_Non_Placeholder()
    'Go to the next slide which has a non placeholder shape
    Dim sld                                       As Slide
    Dim shp                                       As Shape
    Dim currentSlideNumber                        As Integer
    Dim startingSlideNumber                       As Integer
    startingSlideNumber = Application.ActiveWindow.View.Slide.SlideNumber + 1
    For currentSlideNumber = startingSlideNumber To ActivePresentation.Slides.count ' iterate slides
        Set sld = Application.ActivePresentation.Slides(currentSlideNumber)
        
        For Each shp In sld.Shapes
            If shp.Type <> msoPlaceholder Then
                ActiveWindow.View.goToSlide sld.SlideIndex
                
                If currentSlideNumber = Application.ActiveWindow.View.Slide.SlideNumber Then
                    shp.Select
                End If
                
                MsgBox "Non-Placeholder found On slide " & currentSlideNumber & " Shape: " & shp.Name & ". Move To the Next slide And run again To continue searching" ' outputs slide number and shape name
                If Not CommandBars.GetPressedMso("SelectionPane") Then CommandBars.ExecuteMso ("SelectionPane") ' show selection pane
                Exit Sub                          ' end program for user to do things
            End If                                ' end of if type is msomedia
        Next shp                                  ' end of iterate shapes
    Next                                          ' end of iterate slides
    MsgBox "No More Images Found. Move To slide 1 To search again."
End Sub

Sub Shape_Display_Type_And_Details()
    'Provides additional details about currently selected shape or shapes
    Dim currentPresentation As Presentation
    Set currentPresentation = ActivePresentation
    Dim shp    As Shape
    Dim intShapeCount As Integer
    intShapeCount = 0
    Dim strBoxText As String
    
    Dim colBool As New Collection
    colBool.Add "True", "-1"
    colBool.Add "False", "0"
    
    Dim colType As New Collection
    colType.Add "3D model", "30"
    colType.Add "AutoShape", "1"
    colType.Add "Callout", "2"
    colType.Add "Canvas", "20"
    colType.Add "Chart", "3"
    colType.Add "Comment", "4"
    colType.Add "Content Office Add-in", "27"
    colType.Add "Diagram", "21"
    colType.Add "Embedded OLE object", "7"
    colType.Add "Form control", "8"
    colType.Add "Freeform", "5"
    colType.Add "Graphic", "28"
    colType.Add "Group", "6"
    colType.Add "SmartArt graphic", "24"
    colType.Add "Ink", "22"
    colType.Add "Ink comment", "23"
    colType.Add "Line", "9"
    colType.Add "Linked 3D model", "31"
    colType.Add "Linked graphic", "29"
    colType.Add "Linked OLE object", "10"
    colType.Add "Linked picture", "11"
    colType.Add "Media", "16"
    colType.Add "OLE control object", "12"
    colType.Add "Picture", "13"
    colType.Add "Placeholder", "14"
    colType.Add "Script anchor", "18"
    colType.Add "Mixed shape type", "-2"
    colType.Add "Table", "19"
    colType.Add "Text box", "17"
    colType.Add "Text effect", "15"
    colType.Add "Web video", "26"
    
    Dim colPlaceholderType As New Collection
    colPlaceholderType.Add "Bitmap", "9"
    colPlaceholderType.Add "Body", "2"
    colPlaceholderType.Add "Center Title", "3"
    colPlaceholderType.Add "Chart", "8"
    colPlaceholderType.Add "Date", "16"
    colPlaceholderType.Add "Footer", "15"
    colPlaceholderType.Add "Header", "14"
    colPlaceholderType.Add "Media Clip", "10"
    colPlaceholderType.Add "Mixed", "-2"
    colPlaceholderType.Add "Object", "7"
    colPlaceholderType.Add "Organization Chart", "11"
    colPlaceholderType.Add "Picture", "18"
    colPlaceholderType.Add "Slide Number", "13"
    colPlaceholderType.Add "Subtitle", "4"
    colPlaceholderType.Add "Table", "12"
    colPlaceholderType.Add "Title", "1"
    colPlaceholderType.Add "Vertical Body", "6"
    colPlaceholderType.Add "Vertical Object", "17"
    colPlaceholderType.Add "Vertical Title", "5"
    
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        intShapeCount = intShapeCount + 1
        With shp
            
            strBoxText = "Shape is Placeholder: " & colBool(CStr(msoPlaceholder)) & vbCrLf
            
            'TODO here convert above to string before calling mso placeholder was title
            If shp.Type = msoPlaceholder Then
                strBoxText = strBoxText & "PlaceholderFormat.Type: " & colPlaceholderType(.PlaceholderFormat.Type) & vbCrLf _
                           & "PlaceholderFormat.ContainedType: " & colType(.PlaceholderFormat.ContainedType) & vbCrLf
            Else
                strBoxText = strBoxText & "PlaceholderFormat.Type: NA" & vbCrLf _
                           & "PlaceholderFormat.ContainedType: NA" & vbCrLf
            End If
            
            
            
            
            
            
            If shp.HasTextFrame Then
                strBoxText = strBoxText & "Autosize: " & .TextFrame.AutoSize & "(0:no autofit, -2: shrink text, 1: resize shape)" & vbCrLf
            End If
        End With
        
        
        'MsgBox "Shape " & intShapeCount & " of " & ActiveWindow.Selection.ShapeRange.count & vbCrLf
        
        MsgBox strBoxText, vbOKOnly, shp.Name & " (" & intShapeCount & " of " & ActiveWindow.Selection.ShapeRange.count & ")"
        
        
        strBoxText = ""
        
    Next shp
End Sub

