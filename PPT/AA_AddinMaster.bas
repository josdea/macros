Option Explicit
Sub Video_Convert_Embedded_To_Linked()
    Dim oSl    As Slide
    Dim oSh    As Shape
    Dim x      As Long
    Dim sPath  As String
    Dim oNewVid As Shape
    Dim lZOrder As Long
    Dim vidfileName As String
    sPath = ActivePresentation.Path
    For Each oSl In ActivePresentation.Slides
        oSl.Select
        For x = oSl.Shapes.count To 1 Step -1
            Set oSh = oSl.Shapes(x)
            If oSh.Type = msoMedia Then
                If oSh.MediaType = ppMediaTypeMovie Then
                    If oSh.MediaFormat.IsEmbedded Then
                        Set oSl = oSh.Parent
                        lZOrder = oSh.ZOrderPosition
                        vidfileName = oSh.Name
                        Set oNewVid = oSl.Shapes.AddMediaObject2(sPath & "\" & oSh.Name, _
                            msoTrue, msoFalse, _
                            oSh.Left, oSh.Top, _
                            oSh.Width, oSh.Height)
                        Do Until oNewVid.ZOrderPosition = lZOrder
                            oNewVid.ZOrder (msoSendBackward)
                        Loop
                        oSh.Delete
                    End If
                End If
            End If
        Next        ' Shape
    Next        ' Slide
End Sub

Sub Video_Convert_Linked_To_Embedded()
    MsgBox "This can take a few minutes,        't worry"
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

Sub Text_Language_Toggle_US_UK()
    Dim currentSlide As Integer        'current slide number
    Dim currentShape As Integer        ' current shape on current slide or notes
    Dim slideCount As Integer
    Dim shapeCount As Integer        'Find out how many slides there are in the presentation
    Dim noteShape As Shape
    Dim currentLanguage As Integer
    Dim totalShapeCount As Integer
    totalShapeCount = 0
    If MsgBox("Do you want UK English Spelling For Slides And notes? Otherwise US English will be applied", vbYesNo) = vbYes Then
        currentLanguage = msoLanguageIDEnglishUK        'language set to UK
    Else
        currentLanguage = msoLanguageIDEnglishUS        'language set to US
    End If
    slideCount = ActivePresentation.Slides.count        'Get slide count
    MsgBox ("There are " & slideCount & " slides")
    For currentSlide = 1 To slideCount        'Find out how many shapes there are so identify all the text boxes
        shapeCount = ActivePresentation.Slides(currentSlide).Shapes.count        'Loop through all the shapes on that slide changing the language option
        For currentShape = 1 To shapeCount
            If ActivePresentation.Slides(currentSlide).Shapes(currentShape).HasTextFrame Then
                ActivePresentation.Slides(currentSlide).Shapes(currentShape) _
                                                                             .TextFrame.TextRange.LanguageID = currentLanguage
                totalShapeCount = totalShapeCount + 1
            End If
        Next currentShape
        If ActivePresentation.Slides(currentSlide).HasNotesPage Then
            For Each noteShape In ActivePresentation.Slides(currentSlide).NotesPage.Shapes
                If noteShape.HasTextFrame Then
                    noteShape.TextFrame _
                                        .TextRange.LanguageID = currentLanguage
                    totalShapeCount = totalShapeCount + 1
                End If
            Next noteShape
        End If
    Next currentSlide
    MsgBox ("All Done. Total Shapes Set: " & totalShapeCount)
End Sub

Sub Sections_Bulk_Create()
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
Sub Shapes_Count_By_Name()
    Dim sld    As Slide
    Dim shp    As Shape
    Dim shapeNameToCount As String
    Dim shapesCounted As Integer
    shapeNameToCount = InputBox("What Is the shape name To count On all slides (case sensitive?")
    For Each sld In ActivePresentation.Slides        ' iterate slides
        For Each shp In sld.Shapes
            If shp.Name = shapeNameToCount Then
                shapesCounted = shapesCounted + 1
            End If
        Next shp
    Next sld        ' end of iterate slides
    MsgBox shapesCounted & " shapes match the name: " & shapeNameToCount
End Sub

Sub Presenter_Notes_Remove_All()
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

Sub Shapes_Delete_By_Name()
    Dim sld    As Slide
    Dim shp    As Shape
    Dim shapeNameToDelete As String
    Dim shapesDeleted As Integer
    shapeNameToDelete = InputBox("What Is the shape name To delete On all slides?")
    For Each sld In ActivePresentation.Slides        ' iterate slides
        For Each shp In sld.Shapes
            If shp.Name = shapeNameToDelete Then
                shapesDeleted = shapesDeleted + 1
                shp.Delete
            End If
        Next shp
    Next sld        ' end of iterate slides
    MsgBox shapesDeleted & " shapes have been deleted"
End Sub

Sub Comments_Export_To_Disk()
    Dim n      As Integer
    Dim sld    As Slide
    Dim myComment As Comment
    Dim Comment As String
    For Each sld In ActivePresentation.Slides
        For Each myComment In sld.Comments
            Comment = Comment & vbCrLf
            Comment = Comment & "Slide " & sld.SlideNumber & " of " & ActivePresentation.Slides.count & " (" & myComment.DateTime & ")" & vbCrLf
            Comment = Comment & "" & myComment.Text & vbCrLf
        Next myComment
    Next sld
    n = FreeFile()
    Open Environ("USERPROFILE") & "\Desktop\ppt_comments_" & Format(Now(), "yymmddhhmm") & ".txt" For Output As #n
    Print #n, Comment        ' write to file
    Close #n
    MsgBox "All Comments have been written To a Text file (ppt_comments) On your desktop"
End Sub

Sub Comments_Search_And_Export()
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
                Next        'next reply
            End If        'there are replies
            If (InStr(UCase(tempComment), UCase(commentSearch))) Or commentSearch = "" Then
                Comment = Comment & tempComment
                Comment = Comment & vbCrLf
                commentCount = commentCount + 1 + replyCount
            End If        'search matches or is blank
        Next myComment
    Next sld
    Dim n      As Integer
    n = FreeFile()
    Open Environ("USERPROFILE") & "\Desktop\ppt_comments_" & Format(Now(), "yymmdd hhmm") & "-" & commentSearch & ".txt" For Output As #n
    Print #n, Comment        ' write to file
    Close #n
    MsgBox commentCount & " comments and/or replies have been written To a Text file (ppt_comments) On your desktop"
End Sub

Sub Image_Go_To_Next()
    Dim sld                                       As Slide
    Dim shp                                       As Shape
    Dim currentSlideNumber                        As Integer
    Dim startingSlideNumber                       As Integer
    startingSlideNumber = Application.ActiveWindow.View.Slide.SlideNumber + 1
    For currentSlideNumber = startingSlideNumber To ActivePresentation.Slides.count        ' iterate slides
        Set sld = Application.ActivePresentation.Slides(currentSlideNumber)
        For Each shp In sld.Shapes
            If shp.Type = msoPicture Then
                sld.Select        'select current slide
                shp.Select
                MsgBox "Image found On slide " & currentSlideNumber & " Shape: " & shp.Name & ". Move To the Next slide And run again To continue searching"        ' outputs slide number and shape name
                If Not CommandBars.GetPressedMso("SelectionPane") Then CommandBars.ExecuteMso ("SelectionPane")        ' show selection pane
                Exit Sub        ' end program for user to do things
            End If        ' end of if type is msomedia
        Next shp        ' end of iterate shapes
    Next        ' end of iterate slides
    MsgBox "No More Images Found. Move To slide 1 To search again."
End Sub

Sub Video_Go_To_Next()
    Dim sld                                       As Slide
    Dim shp                                       As Shape
    Dim currentSlideNumber                        As Integer
    Dim startingSlideNumber                       As Integer
    startingSlideNumber = Application.ActiveWindow.View.Slide.SlideNumber
    For currentSlideNumber = startingSlideNumber To ActivePresentation.Slides.count        ' iterate slides
        Set sld = Application.ActivePresentation.Slides(currentSlideNumber)
        For Each shp In sld.Shapes
            If shp.Type = msoMedia Then
                If shp.MediaType = ppMediaTypeMovie Then
                    sld.Select        'select current slide
                    MsgBox "Video found On slide " & currentSlideNumber & " Shape: " & shp.Name & ". Move To the Next slide And run again To continue searching"        ' outputs slide number and shape name
                    If Not CommandBars.GetPressedMso("SelectionPane") Then CommandBars.ExecuteMso ("SelectionPane")        ' show selection pane
                    Exit Sub        ' end program for user to do things
                End If
            End If        ' end of if type is msomedia
        Next shp        ' end of iterate shapes
    Next        ' end of iterate slides
    MsgBox "No More Videos Found. Move To slide 1 To search again."
End Sub

Sub Slide_Go_To()
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

Sub PresenterNotes_Remove_Text_In_Hashtags()
    Dim stringBwDels As String, originalString As String, firstDelPos As Integer, secondDelPos As Integer, stringToReplace As String, replacedString As String
    Dim sld                                       As Slide        ' declare slide object
    Dim shp                                       As Shape        ' declare shape object
    For Each sld In ActivePresentation.Slides        ' iterate slides
        For Each shp In sld.NotesPage.Shapes        ' iterate note shapes
            If shp.Type = msoPlaceholder Then        ' check if its a placeholder
            If shp.PlaceholderFormat.Type = ppPlaceholderBody Then        ' its presenter notes
            originalString = shp.TextFrame2.TextRange.Text
            stringToReplace = ""
            firstDelPos = InStr(originalString, "##") - 1        ' position of start delimiter
            secondDelPos = InStr(firstDelPos + 2, originalString, "##")        ' position of end delimiter
            If secondDelPos <> 0 Then
                stringBwDels = Mid(originalString, firstDelPos + 1, secondDelPos - firstDelPos + 1)        'extract the string between two delimiters
            Else
                stringBwDels = 0
            End If
            replacedString = Replace(originalString, stringBwDels, stringToReplace)
            shp.TextFrame2.TextRange.Text = replacedString
        End If
    End If
Next shp        ' end of iterate shapes
Next sld        ' end of iterate slides
End Sub

Sub PresenterNotes_Toggle_Visibility()
    Dim toggleOn                                  As Boolean
    toggleOn = msoTrue
    If MsgBox("Do you want To hide all presenter note shapes On the notes pages. If you answer        'no' then they will all be made visible.", (vbYesNo + vbQuestion), "Toggle Presenter Notes?") = vbYes Then
        toggleOn = msoFalse
    End If
    Dim sld                                       As Slide        ' declare slide object
    Dim shp                                       As Shape        ' declare shape object
    For Each sld In ActivePresentation.Slides        ' iterate slides
        For Each shp In sld.NotesPage.Shapes        ' iterate note shapes
            If shp.Type = msoPlaceholder Then        ' check if its a placeholder
            If shp.PlaceholderFormat.Type = ppPlaceholderBody Then        ' its presenter notes
            shp.Visible = toggleOn
        End If
    End If
Next shp        ' end of iterate shapes
Next sld        ' end of iterate slides
If Not CommandBars.GetPressedMso("SelectionPane") Then CommandBars.ExecuteMso ("SelectionPane")
MsgBox "All done toggling presenter notes"
ActiveWindow.ViewType = ppViewNotesPage
End Sub

Sub Text_Font_Reset_To_Master()
    Dim sld                                       As Slide
    Dim shp                                       As Shape
    Dim shapesAffected As Integer
    shapesAffected = shapesAffected + 1
    If MsgBox("Are you sure you want To reset all titles, text, And notes To the master font theme", (vbYesNo + vbQuestion), "Reset Font?") = vbYes Then
        For Each sld In ActivePresentation.Slides        ' iterate slides
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
            Next shp        ' end of iterate shapes
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
        Next sld        ' end of iterate slides
        MsgBox "All Done. " & shapesAffected & " shapes have been search And reset"
    Else
        MsgBox ("Action canceled.")
    End If
End Sub

Sub Text_Remove_Double_Spaces()
    Dim spacesRemoved As Integer
    Dim sld    As Slide
    Dim shp    As Shape
    Dim shpText As String
    spacesRemoved = 0
    If MsgBox("Do you want To replace all instances of multiple spaces With one space", (vbYesNo + vbQuestion), "Remove extra Spaces?") = vbYes Then
        For Each sld In ActivePresentation.Slides
            For Each shp In sld.Shapes
                If shp.HasTextFrame Then
                    shpText = shp.TextFrame.TextRange.Text        'Get the shape's text
                    Do While InStr(shpText, "  ") > 0
                        shpText = Trim(Replace(shpText, "  ", " "))
                        spacesRemoved = spacesRemoved + 1
                    Loop
                    shp.TextFrame.TextRange.Text = shpText        'Put the new text in the shape
                Else
                    shpText = vbNullString
                End If
            Next shp
            For Each shp In sld.NotesPage.Shapes
                If shp.HasTextFrame Then
                    shpText = shp.TextFrame.TextRange.Text        'Get the shape's text
                    Do While InStr(shpText, "  ") > 0
                        shpText = Trim(Replace(shpText, "  ", " "))
                        spacesRemoved = spacesRemoved + 1
                    Loop
                    shp.TextFrame.TextRange.Text = shpText        'Put the new text in the shape
                Else
                    shpText = vbNullString
                End If
            Next shp
        Next sld
    End If
    MsgBox spacesRemoved & " places where extra spacing was removed"
End Sub

Sub Layout_Create_All_Types()
    Dim sld    As Slide
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutBlank)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutBlank (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutChart)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutChart (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutChartAndText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutChartAndText (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutClipartAndText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutClipartAndText (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutClipArtAndVerticalText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutClipArtAndVerticalText (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutComparison)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutComparison (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutContentWithCaption)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutContentWithCaption (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutCustom)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutCustom (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutFourObjects)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutFourObjects (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutLargeObject)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutLargeObject (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutMediaClipAndText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutMediaClipAndText (" & sld.Layout & ")"
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutMixed (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutObject)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutObject (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutObjectAndText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutObjectAndText (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutObjectAndTwoObjects)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutObjectAndTwoObjects (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutObjectOverText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutObjectOverText (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutOrgchart)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutOrgchart (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutPictureWithCaption)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutPictureWithCaption (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutSectionHeader)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutSectionHeader (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutTable)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutTable (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutText (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutTextAndChart)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutTextAndChart (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutTextAndClipart)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutTextAndClipart (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutTextAndMediaClip)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutTextAndMediaClip (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutTextAndObject)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutTextAndObject (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutTextAndTwoObjects)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutTextAndTwoObjects (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutTextOverObject)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutTextOverObject (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutTitle)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutTitle (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutTitleOnly)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutTitleOnly (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutTwoColumnText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutTwoColumnText (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutTwoObjects)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutTwoObjects (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutTwoObjectsAndObject)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutTwoObjectsAndObject (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutTwoObjectsAndText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutTwoObjectsAndText (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutTwoObjectsOverText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutTwoObjectsOverText (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutVerticalText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutVerticalText (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutVerticalTitleAndText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutVerticalTitleAndText (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutVerticalTitleAndTextOverChart)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutVerticalTitleAndTextOverChart (" & sld.Layout & ")"
End Sub

Sub Image_Add_Comment_For_Missing_AltText()
    Dim sld    As Slide        ' declare slide object
    Dim shp                                       As Shape        ' declare shape object
    For Each sld In ActivePresentation.Slides        ' iterate slides
        For Each shp In sld.Shapes        ' iterate note shapes
            If shp.Type = msoPicture Then        ' check if its a placeholder
            If shp.AlternativeText = "" Or InStr(shp.AlternativeText, "generated") Then
                sld.Comments.Add 12, 12, "Auto", "JMD", "TODO: Need Image ID Or source"
            End If
        End If
    Next shp        ' end of iterate shapes
Next sld        ' end of iterate slides
End Sub

Sub PresenterNotes_Add_Comment_If_Empty()
    Dim sld    As Slide        ' declare slide object
    Dim shp                                       As Shape        ' declare shape object
    For Each sld In ActivePresentation.Slides        ' iterate slides
        For Each shp In sld.NotesPage.Shapes        ' iterate note shapes
            If shp.Type = msoPlaceholder Then        ' check if its a placeholder
            If shp.PlaceholderFormat.Type = ppPlaceholderBody Then        ' its presenter notes
            If shp.TextFrame.HasText = FALSE Then
                sld.Comments.Add 12, 12, "Auto", "JMD", "TODO: Need Presenter Notes"
            Else If shp.textframe.textrange.text = "" then
                sld.Comments.Add 12, 12, "Auto", "JMD", "TODO: Need Presenter Notes"
            End If
            replacedString = Replace(originalString, stringBwDels, stringToReplace)
            shp.TextFrame2.TextRange.Text = replacedString
        End If
    End If
Next shp        ' end of iterate shapes
Next sld        ' end of iterate slides
End Sub