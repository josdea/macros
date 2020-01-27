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
        Next                                      ' Shape
    Next                                          ' Slide
    MsgBox "All Done"
End Sub

Sub Video_Convert_Linked_To_Embedded()
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

Sub Text_Language_Toggle_US_UK()                  ' checked 1/17/20
    Dim currentSlide As Integer                   'current slide number
    Dim currentShape As Integer                   ' current shape on current slide or notes
    Dim slideCount As Integer
    Dim shapeCount As Integer                     'Find out how many slides there are in the presentation
    Dim noteShape As Shape
    Dim currentLanguage As Integer
    Dim totalShapeCount As Integer
    totalShapeCount = 0
    If MsgBox("Do you want UK English Spelling For Slides And notes? Otherwise US English will be applied", vbYesNo) = vbYes Then
        currentLanguage = msoLanguageIDEnglishUK  'language set to UK
    Else
        currentLanguage = msoLanguageIDEnglishUS  'language set to US
    End If
    slideCount = ActivePresentation.Slides.count  'Get slide count
    For currentSlide = 1 To slideCount            'Find out how many shapes there are so identify all the text boxes
        shapeCount = ActivePresentation.Slides(currentSlide).Shapes.count 'Loop through all the shapes on that slide changing the language option
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
    MsgBox ("All Done. Total Shapes Set: " & totalShapeCount & ". Press F7 To rerun spellcheck.")
End Sub

Sub Sections_Bulk_Create()                        ' checked 1/17/20
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
Sub Shapes_Count_By_Name()                        ' checked 1/17/20
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
    Dim n      As Integer
    n = FreeFile()
    Open Environ("USERPROFILE") & "\Desktop\ppt_comments_" & Format(Now(), "yymmdd hhmm") & "-" & commentSearch & ".txt" For Output As #n
    Print #n, Comment                             ' write to file
    Close #n
    MsgBox commentCount & " comments and/or replies have been written To a Text file (ppt_comments) On your desktop"
End Sub

Sub Image_Go_To_Next()                            ' checked 1/17/20
    Dim sld                                       As Slide
    Dim shp                                       As Shape
    Dim currentSlideNumber                        As Integer
    Dim startingSlideNumber                       As Integer
    startingSlideNumber = Application.ActiveWindow.View.Slide.SlideNumber + 1
    For currentSlideNumber = startingSlideNumber To ActivePresentation.Slides.count ' iterate slides
        Set sld = Application.ActivePresentation.Slides(currentSlideNumber)
        For Each shp In sld.Shapes
            If shp.Type = msoPicture Or shp.PlaceholderFormat.ContainedType = msoPicture Then
                sld.Select                        'select current slide
                shp.Select
                MsgBox "Image found On slide " & currentSlideNumber & " Shape: " & shp.Name & ". Move To the Next slide And run again To continue searching" ' outputs slide number and shape name
                If Not CommandBars.GetPressedMso("SelectionPane") Then CommandBars.ExecuteMso ("SelectionPane") ' show selection pane
                Exit Sub                          ' end program for user to do things
            End If                                ' end of if type is msomedia
        Next shp                                  ' end of iterate shapes
    Next                                          ' end of iterate slides
    MsgBox "No More Images Found. Move To slide 1 To search again."
End Sub

Sub Video_Go_To_Next()                            ' checked 1/17/20
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
                    sld.Select                    'select current slide
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
    Dim toggleOn                                  As Boolean
    toggleOn = msoTrue
    If MsgBox("Do you want To hide all presenter note shapes On the notes pages. If you answer        'no' then they will all be made visible.", (vbYesNo + vbQuestion), "Toggle Presenter Notes?") = vbYes Then
    toggleOn = msoFalse
End If
Dim sld                                       As Slide ' declare slide object
Dim shp                                       As Shape ' declare shape object
For Each sld In ActivePresentation.Slides         ' iterate slides
    For Each shp In sld.NotesPage.Shapes          ' iterate note shapes
        If shp.Type = msoPlaceholder Then         ' check if its a placeholder
            If shp.PlaceholderFormat.Type = ppPlaceholderBody Then ' its presenter notes
                shp.Visible = toggleOn
            End If
        End If
    Next shp                                      ' end of iterate shapes
Next sld                                          ' end of iterate slides
If Not CommandBars.GetPressedMso("SelectionPane") Then CommandBars.ExecuteMso ("SelectionPane")
MsgBox "All done toggling presenter notes"
ActiveWindow.ViewType = ppViewNotesPage
End Sub

Sub Text_Font_Reset_To_Master()                   ' checked 1/17/20
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
End Sub

Sub Image_Add_Comment_For_Missing_AltText()       ' checked 1/17/20
    Dim sld    As Slide                           ' declare slide object
    Dim shp                                       As Shape ' declare shape object
    Dim intCommentCount As Integer
    intCommentCount = 0
    For Each sld In ActivePresentation.Slides     ' iterate slides
        For Each shp In sld.Shapes                ' iterate note shapes
            If shp.Type = msoPicture Or shp.PlaceholderFormat.ContainedType = msoPicture Then ' check if its a placeholder
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
    Dim sld                                       As Slide
    Dim shp                                       As Shape
    Dim currentSlideNumber                        As Integer
    Dim startingSlideNumber                       As Integer
    startingSlideNumber = Application.ActiveWindow.View.Slide.SlideNumber + 1
    For currentSlideNumber = startingSlideNumber To ActivePresentation.Slides.count ' iterate slides
        Set sld = Application.ActivePresentation.Slides(currentSlideNumber)
        For Each shp In sld.Shapes
            If shp.Type <> msoPlaceholder Then
                sld.Select                        'select current slide
                shp.Select
                MsgBox "Non-Placeholder found On slide " & currentSlideNumber & " Shape: " & shp.Name & ". Move To the Next slide And run again To continue searching" ' outputs slide number and shape name
                If Not CommandBars.GetPressedMso("SelectionPane") Then CommandBars.ExecuteMso ("SelectionPane") ' show selection pane
                Exit Sub                          ' end program for user to do things
            End If                                ' end of if type is msomedia
        Next shp                                  ' end of iterate shapes
    Next                                          ' end of iterate slides
    MsgBox "No More Images Found. Move To slide 1 To search again."
End Sub

Sub Shape_Display_Type_And_Details()
    Dim currentPresentation As Presentation
    Set currentPresentation = ActivePresentation
    Dim shp    As Shape
    Dim intShapeCount As Integer
    intShapeCount = 0
    For Each shp In ActiveWindow.Selection.ShapeRange
        intShapeCount = intShapeCount + 1
        'MsgBox "Shape " & intShapeCount & " of " & ActiveWindow.Selection.ShapeRange.count & vbCrLf
        With shp
            MsgBox "(-1: true, 0:false)" & vbCrLf _
                 & "Shape: " & .Name & " (" & intShapeCount & " of " & ActiveWindow.Selection.ShapeRange.count & ")" & vbCrLf _
                 & "Has Text Frame: " & .HasTextFrame & vbCrLf _
                 & "Has Text: " & .TextFrame.HasText & vbCrLf _
                 & "Height: " & .Height & " points Or " & .Height / 72 & " inches" & vbCrLf _
                 & "Width: " & .Width & " points Or " & .Width / 72 & " inches" & vbCrLf _
                 & "LockAspectRatio: " & .LockAspectRatio & vbCrLf _
                 & "Left: " & .Left & " points Or " & .Left / 72 & " inches" & vbCrLf _
                 & "Top: " & .Top & " points Or " & .Top / 72 & " inches" & vbCrLf _
                 & "Type: " & .Type & vbCrLf _
                 & "PlaceholderFormat.Type: " & .PlaceholderFormat.Type & vbCrLf _
                 & "PlaceholderFormat.ContainedType: " & .PlaceholderFormat.ContainedType & vbCrLf _
                 & "Autosize: " & .TextFrame.AutoSize & "(0:no autofit, -2: shrink text, 1: resize shape)" & vbCrLf
            
        End With
    Next shp
End Sub

