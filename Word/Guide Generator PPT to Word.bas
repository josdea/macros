Option Explicit
Sub createGuides()
    statusOutput "**********STARTING GUIDE**********"
    
    Dim docTempTarget                             As Document
    Set docTempTarget = wordDocumentSelection()
    
    Call createStyles(docTempTarget)
    Call createInstructorGuide(docTempTarget)
    Call removeNonBreakingSpaces(docTempTarget)
    Call removeDoubleTabs(docTempTarget)
    Call removeDoubleParagraphs(docTempTarget)
    
    If (MsgBox("All Done. Do you want to save at this time?", (vbYesNo + vbQuestion), "Save?") = vbYes) Then
        docTempTarget.Save
    End If
    
    If (MsgBox("Do you want to run it again?", (vbYesNo + vbQuestion), "Again?") = vbYes) Then
        Call createGuides
    End If
    
End Sub

Function wordDocumentSelection() As Document
    Dim docs                                      As Documents
    Set docs = Documents
    Dim doc    As Document
    Dim intDocSelect As Integer
    Dim docCount As Integer:     docCount = 0
    Dim strDocSelect As String
    
    strDocSelect = "0. Create New Word Document" & vbCrLf
    For Each doc In docs
        docCount = docCount + 1
        strDocSelect = strDocSelect & docCount & ". " & Left(doc.Name, 40) & "..." & vbCrLf
    Next doc
    intDocSelect = getNumberInput(strDocSelect, "Select from Open Documents", 0, 0, docs.Count)
    
    If intDocSelect > 0 Then
        Set wordDocumentSelection = Documents(intDocSelect)
    Else
        Set wordDocumentSelection = Application.Documents.Add
    End If
    
End Function

Private Function promptPowerpointFile() As Object
    Dim objPPT                                    As Object
    Set objPPT = CreateObject("PowerPoint.Application")
    Dim oPresentations As Object
    Set oPresentations = objPPT.presentations
    Dim pptDoc                                       As Object
    
    Dim intDocSelect As Integer
    Dim docCount As Integer:     docCount = 0
    Dim strDocSelect As String
    
    '  If oPresentations.Count = 0 Then
    '  Set promptPowerpointFile = readPowerpointFile()
    ' Else
    
    strDocSelect = "0. Open PPT from disk" & vbCrLf
    For Each pptDoc In oPresentations
        docCount = docCount + 1
        strDocSelect = strDocSelect & docCount & ". " & Left(pptDoc.Name, 40) & "..." & vbCrLf
    Next pptDoc
    intDocSelect = getNumberInput(strDocSelect, "Select from Open Presentations", "0", 0, oPresentations.Count)
    
    If intDocSelect > 0 Then
        Set promptPowerpointFile = oPresentations(intDocSelect)
    Else
        Set promptPowerpointFile = readPowerpointFile()
    End If
    '  End If
    
End Function
Private Function readPowerpointFile() As Object
    Dim objPPT                                    As Object
    Set objPPT = CreateObject("PowerPoint.Application") ' Create and initialize the PowerPoint application object.

    With objPPT
        .Activate                                 ' Activate the PPT application object.
        .Visible = True                           ' Make it visible.
        On Error GoTo failCleanly                 ' Trap any errors and abort gracefully if the presentation file fails to open ... and remain open.
        Dim dlgOpen                               As FileDialog ' Show the Open File dialog.
        Set dlgOpen = .FileDialog(Type:=msoFileDialogOpen)
        With dlgOpen
            .Title = "Select an Input Course File"
            .InitialFileName = Environ("USERPROFILE") & "\Downloads\" ' Set the default directory path.
            .Show
            .Execute
        End With
        Set readPowerpointFile = .ActivePresentation ' Set the active presentation to be the source file.
    End With                                      ' End initializing the PowerPoint object.
    statusOutput "PowerPoint Opening Completed"
    Exit Function
    ' Error trap for subroutine's On Error statement above.
failCleanly:
    MsgBox "We had some trouble reading the PowerPoint file. To continue, try re-running the macro. If that still does not work try closing Word completely And restarting it before rerunning the macro again.", Buttons:=vbExclamation, Title:="PowerPoint Had Trouble"
    End
    With objPPT
        .Activate                                 ' Activate PowerPoint
        .Quit                                     ' Exit PowerPoint
    End With
    
    
    Exit Function
End Function
Function createInstructorGuide(docTempTarget As Document)
    Dim tblGuide                                  As Table

    Dim strTempImgDir                             As String: strTempImgDir = ""
    Dim sld                                       As Object
    Dim strExportFormat                           As String: strExportFormat = "GIF"
    Dim strThisImgPath                            As String: strThisImgPath = ""
    Dim shpThisImg                                As Word.Shape
    Dim shp                                       As Object
    Dim boolIncludePresenterNotes                 As Boolean: boolIncludePresenterNotes = True
    Dim boolIncludeHiddenSlides                   As Boolean: boolIncludeHiddenSlides = False
    Dim dblImgWidth                               As Double
    Dim strImgAlign                               As String: strImgAlign = wdShapeRight
    Dim objSrcFile                                    As Object 'PPT File
    Dim strModuleTitle                            As String: strModuleTitle = ""
    Dim intSlideNumber As Integer: intSlideNumber = 0
    Dim boolSectionNumbering As Boolean: boolSectionNumbering = False
    'Dim strWordForSlide As String: strWordForSlide = "Slide "
    Dim strWordForSlide As String: strWordForSlide = readDocumentVariable(docTempTarget, "WordForSlide", "Slide ")
    Dim intModuleNumber As Integer: intModuleNumber = 1
    Dim dblRowHeight As Double


    Set objSrcFile = promptPowerpointFile()       'Open Powerpoint PPT file

    docTempTarget.Activate
    
    
    'PROMPT FOR SPECS
    If (MsgBox("Include presenter notes?", (vbYesNo + vbQuestion), "Instructor Guide?") = vbNo) Then
        boolIncludePresenterNotes = False
    End If
    If (MsgBox("Include hidden slides?", (vbYesNo + vbQuestion), "Hidden Slides?") = vbYes) Then
        boolIncludeHiddenSlides = True
    End If
    If (MsgBox("Image To the left? (otherwise it will be right aligned", (vbYesNo + vbQuestion), "Image Align?") = vbYes) Then
        strImgAlign = wdShapeLeft
    End If
    strWordForSlide = InputBox("Desired word or translation of 'Slide'?" _
                    & vbCrLf & "This can also be changed with find and replace later." _
                    & vbCrLf & "Examples:" _
                    & vbCrLf & "French - 'Diapositive' | Spanish - 'Diapositiva'" _
    & vbCrLf & "(leave blank for only slide number)", "Slide Translation?", strWordForSlide) ' TODO check what happens if user cancels
    Call updateDocumentVariable(docTempTarget, "WordForSlide", strWordForSlide)
    
    dblRowHeight = getRowHeightFromInput(docTempTarget)
    
    dblImgWidth = getImgWidthFromInput(docTempTarget)
    
    
    'END PROMPT FOR SPECS
    
    
    strTempImgDir = exportSlideImages(objSrcFile, strExportFormat) 'export slide images
    
    If objSrcFile.SectionProperties.Count > 0 Then ' sections exist prompt for number preference
        If (MsgBox("Sections Exist. Reset slide numbering for each section?", (vbYesNo + vbQuestion), "Section Numbering?") = vbYes) Then
            boolSectionNumbering = True
        End If
    End If                                        'sections exist prompt
    
    docTempTarget.Activate
    
    With objSrcFile
        For Each sld In .slides
            intSlideNumber = intSlideNumber + 1
            If sld.slideShowTransition.Hidden = False Or boolIncludeHiddenSlides = True Then
                
                If .SectionProperties.Count > 0 Then 'sections exist
                    
                    If (sld.slideNumber = 1) Then 'First slide so output section info
                        With Selection            'Begin new table
                            .MoveEnd Unit:=wdStory ' Get clear of any content
                            .Start = .End         ' move to the end
                            strModuleTitle = getTitleFromFirstSlide(sld, intModuleNumber) 'get title shape from the slide if there
                            
                            .TypeText (strModuleTitle) 'Module Title
                            
                            .Style = docTempTarget.Styles("Heading 1") 'style module title
                            .TypeText (vbCrLf)    'new line
                            .ClearFormatting
                            '.MoveEnd Unit:=wdStory        ' Get clear of any content
                            '.Start = .End
                            
                            ' Create and format the target table.
                            Set tblGuide = createTable(docTempTarget, dblRowHeight)
                            tblGuide.cell(1, 1).Select
                            .End = .Start
                            .TypeText (strModuleTitle)
                            'Call setTableFormat(tblGuide)
                        End With
                        
                        If (sld.slideNumber <> .slides.Count) Then 'First slide and not the only in the PPT
                            If (sld.sectionIndex <> .slides(sld.slideNumber + 1).sectionIndex) Then 'first slide but also last in a section
                                Call setTableFormat(tblGuide)
                            End If
                        Else                      ' first slide but also the last
                            Call setTableFormat(tblGuide)
                        End If
                        
                    ElseIf (sld.sectionIndex <> .slides(sld.slideNumber - 1).sectionIndex) Then 'Not the first slide but section index is different than previous slide
                        If boolSectionNumbering = True Then
                            intSlideNumber = 1
                        End If
                        intModuleNumber = intModuleNumber + 1 ' new section so increase module number
                        With Selection            'Begin new table
                            .MoveEnd Unit:=wdStory ' Get clear of any content
                            .Start = .End         ' move to the end
                            strModuleTitle = getTitleFromFirstSlide(sld, intModuleNumber) 'get title shape from the slide if there
                            .TypeText (strModuleTitle) 'module title
                            .Style = docTempTarget.Styles("Heading 1")
                            .TypeText (vbCrLf)
                            .ClearFormatting
                            .MoveEnd Unit:=wdStory ' Get clear of any content
                            .Start = .End
                            
                            Set tblGuide = createTable(docTempTarget, dblRowHeight)
                            tblGuide.cell(1, 1).Select
                            .End = .Start
                            .TypeText (strModuleTitle)
                            ' Call setTableFormat(tblGuide)
                        End With
                    ElseIf (sld.slideNumber = .slides.Count) Then 'Last slide of the presentation so the last slide in a section
                        Call setTableFormat(tblGuide)
                        
                    ElseIf (sld.sectionIndex <> .slides(sld.slideNumber + 1).sectionIndex) Then 'not last slide but different than next section
                        Call setTableFormat(tblGuide)
                        
                    End If                        'End of first slide or differing sections IF
                Else                              ' there are no sections in current PPT
                    If sld.slideNumber = 1 Then   'so look to first slide for title
                        
                        With Selection            'Begin new table
                            .MoveEnd Unit:=wdStory ' Get clear of any content
                            .Start = .End         ' move to the end
                            strModuleTitle = getTitleFromFirstSlide(sld, intModuleNumber) 'get title shape from the slide if there
                            'strModuleTitle = getTitleFromFirstSlide(objSrcFile)        'get title shape from first slide of PPT
                            
                            .TypeText (strModuleTitle)
                            
                            .Style = docTempTarget.Styles("Heading 1")
                            .TypeText (vbCrLf)
                            .ClearFormatting
                            .MoveEnd Unit:=wdStory ' Get clear of any content
                            .Start = .End
                            
                            Set tblGuide = createTable(docTempTarget, dblRowHeight) ' Create and format the target table.
                            
                            tblGuide.cell(1, 1).Select
                            .End = .Start
                            .TypeText (strModuleTitle)
                            'Call setTableFormat(tblGuide)
                        End With
                    ElseIf sld.slideNumber = .slides.Count Then 'last slide in presentaiton
                        Call setTableFormat(tblGuide)
                        
                    End If                        'slide 1
                    
                End If                            'sections exist
                
                With docTempTarget
                    .Activate
                    With Selection
                        .MoveRight Unit:=wdCell, Extend:=wdMove ' Move right again to create the next line in the table if it isnt the first slide
                        .TypeText (strWordForSlide & intSlideNumber & ": ") ' Add the slide number.
                        
                        .Style = "Slide Number"   ' Set the style. TODO check if exists
                        .Cells(1).Select          ' Move the cursor to the start of the slide number before importing the image. Otherwise, the anchor will be at the end of the cell contents.
                        .End = .Start
                        strThisImgPath = strTempImgDir & "\Slide" & sld.slideNumber & "." & strExportFormat
                        
                        '.Style = "Slide Image" 'TODO style all images separately somehow
                        Set shpThisImg = .InlineShapes.AddPicture(FileName:=strThisImgPath, LinkToFile:=False, SaveWithDocument:=True).ConvertToShape
                        
                        Call wrapImage(shpThisImg, dblImgWidth, strImgAlign) ' TODO
                        shpThisImg.AlternativeText = "Slide " & sld.slideNumber
                        
                        ' Move right to end of the cell.
                        .Cells(1).Select          'TODO check if the cell is greater than what can fit on a page and change overall height
                        .Start = .End
                        .MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
                        
                        If boolIncludePresenterNotes = True Then
                            
                            For Each shp In sld.NotesPage.Shapes
                                If shp.Type = 14 Then 'is a placeholder
                                    If shp.placeholderformat.Type = 2 Then 'If there's a body to the shape (ppPlaceholderBody = 2).
                                        If shp.HasTextFrame Then ' Check that a text frame exists in the body.
                                            If shp.TextFrame.HasText Then ' Check for text in that text frame.
                                                .TypeText (vbCrLf)
                                                .Style = "Slide Text" 'TODO check if style exists
                                                .TypeText (shp.TextFrame.TextRange.Text) ' direct replace instead of copy TODO verify is working and how it handles font changes
                                                ' NOTE: The following line hangs occasionally. Simply press the Debug button on error, then continue running by pressing F5.
                                                ' shp.TextFrame.TextRange.Copy    ' Copy the Note text to the clipboard. It'll be pasted into the current Slide Notes cell in a few lines, when we return to the table.
                                                '.Paste
                                                .Start = .End 'todo make sure works
                                                .TypeBackspace ' todo make sure this works to remove ending paragraph symbol
                                                Exit For
                                                
                                                ' TODO multiple file support
                                            End If ' End text in the text frame test.
                                        End If    ' End Body text frame test.
                                    End If        ' End Placeholder Body test.
                                End If            'is a placeholder
                            Next shp
                            
                        End If                    'include presenter notes
                        
                    End With                      'selection of temptarget
                    
                End With                          'docTempTarget
            End If                                ' hidden slides
        Next sld
        
    End With                                      'objSrcFile
    
    If (MsgBox("Finished with PPT File. Do you want to close it?", (vbYesNo + vbQuestion), "Close PPT?") = vbYes) Then
        objSrcFile.Close
    End If
    
    Call setTableFormat(tblGuide)                 'TODO move backt o main at top maybe or
    
    'Display nav pane
    With ActiveWindow
        .View.ShowHeading 2
        .DocumentMap = True
    End With
    
End Function

Function getNumberInput(strMessage As String, strBoxTitle As String, strDefaultValue As String, intMinNumber As Integer, intMaxNumber As Integer) As Double
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

Function getImgWidthFromInput(docTempTarget As Document) As Double
    Dim strInputField                             As String
    Dim defaultWidth As String
    defaultWidth = readDocumentVariable(docTempTarget, "ImageWidth", "3.5")
    
    Do
        'Retrieve an answer from the user
        strInputField = InputBox("Desired image width (0.5 - 8 inches).", "Image Width?", defaultWidth)
        If StrComp(strInputField, "x", 1) = 0 Then
            End
        ElseIf TypeName(strInputField) = "Boolean" Then 'Check if user selected cancel button
            getImgWidthFromInput = 0
        ElseIf Not IsNumeric(strInputField) Then  'input wasnt numeric
            getImgWidthFromInput = 0
        Else
            getImgWidthFromInput = strInputField
            Call updateDocumentVariable(docTempTarget, "ImageWidth", strInputField)
        End If
    Loop While getImgWidthFromInput <= 0.5 Or getImgWidthFromInput > 8
    
End Function

Function getRowHeightFromInput(docTempTarget As Document) As Double
    Dim strInputField                             As String
    Dim defaultHeight As String
    defaultHeight = readDocumentVariable(docTempTarget, "RowHeight", "4")
    
    Do
        'Retrieve an answer from the user
        strInputField = InputBox("Desired row height (0.25 - 11 inches). Use '4' to have two slides per page, '8' for one.", "Row height?", defaultHeight)
        If StrComp(strInputField, "x", 1) = 0 Then
            End
        ElseIf TypeName(strInputField) = "Boolean" Then 'Check if user selected cancel button
            getRowHeightFromInput = 0
        ElseIf Not IsNumeric(strInputField) Then  'input wasnt numeric
            getRowHeightFromInput = 0
        Else
            getRowHeightFromInput = strInputField
            Call updateDocumentVariable(docTempTarget, "RowHeight", strInputField)
        End If
    Loop While getRowHeightFromInput <= 0.25 Or getRowHeightFromInput > 11
    
End Function

Function getTitleFromFirstSlide(sld As Object, intModuleNumber As Integer) As String 'looks at first slide of PPT file and return title text if there
    Dim shp                                       As Object
    sld.Select
    For Each shp In sld.Shapes
        If shp.Type = 14 Then                     'shape is a placeholder
            If shp.placeholderformat.Type = 2 Or shp.placeholderformat.Type = 3 Or shp.placeholderformat.Type = 1 Then 'is a slide title
                If shp.HasTextFrame Then          ' Check that a text frame exists in the body.
                    If shp.TextFrame.HasText Then ' Check for text in that text frame.
                        getTitleFromFirstSlide = InputBox("Is this the title of this module", "Module Title?", shp.TextFrame.TextRange.Text)
                        Exit For
                    End If                        ' End text in the text frame test.
                End If                            ' End Body text frame test.
            End If                                ' End Placeholder Body test.
        End If                                    'is a placeholder
    Next shp
    
    If getTitleFromFirstSlide = "" Then           'if there was no title placeholder on the first slide
        getTitleFromFirstSlide = InputBox("Title shape not found. What should the module title be?", "Module Title?", "Module " & intModuleNumber)
        
    End If
    
End Function

Function exportSlideImages(objSrcFile As Object, strExportFormat As String) As String
    statusOutput "Exporting image To file"
    Dim strTempImgDir                             As String: strTempImgDir = objSrcFile.Path & "\pptSlideExport" & Format(Now(), "yymmddhhmm")
    Dim intScaleWidth                             As Integer: intScaleWidth = 700
    
    With objSrcFile
        .Export Path:=strTempImgDir, FilterName:=strExportFormat, ScaleWidth:=intScaleWidth
    End With
    exportSlideImages = strTempImgDir             'returns folder location of all images
    
End Function

Function wrapImage(shp As Word.Shape, dblImgWidth As Double, strImgAlign As String)
    Dim intBorderWidth                            As Double: intBorderWidth = 0.25

    With shp
        .LockAspectRatio = msoTrue                'prevent misformed shapes
        .Width = InchesToPoints(dblImgWidth)      'apply image width from previous prompt
        .RelativeHorizontalPosition = wdRelativeHorizontalPositionColumn
        .Line.Visible = msoTrue                   'make a border for images
        .Line.ForeColor.RGB = RGB(0, 0, 0)        'make border black
        .Line.Weight = intBorderWidth             'assign border width for images
        With .WrapFormat
            .Type = wdWrapSquare
            .Side = wdWrapBoth                    ' Wraps both sides depending on position of .left below
            .DistanceTop = InchesToPoints(0.2)
            .DistanceBottom = InchesToPoints(0.1)
            .DistanceLeft = InchesToPoints(0.1)
            .DistanceRight = InchesToPoints(0.1)
        End With                                  ' Stop wrap formatting.
        .LayoutInCell = True
        .Left = strImgAlign                       'align image based on previous prompt
    End With
    
End Function

Function setTableFormat(tblGuide As Table)

    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = -603930625
    End With
    
    With tblGuide
        .Style = "Table Grid"
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
        .Rows.Borders.InsideLineStyle = wdLineStyleSingle
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = -603930625
        End With
        With .Borders(wdBorderHorizontal)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = -603930625
        End With
        .ApplyStyleHeadingRows = True
        .Rows(1).Height = InchesToPoints(0.3)
        .Rows.HeadingFormat = False               'removes all rows as a heading
        .Rows(1).HeadingFormat = True             'says there is a heading on row 1
        .Rows(1).Range.Paragraphs.Alignment = wdAlignParagraphCenter 'center row 1 text
        .Rows(1).Range.Bold = True                'bold row 1 text
    End With
End Function
Function removeNonBreakingSpaces(docTempTarget As Document) ' removes nonbreaking spaces throughout from PPT export
    docTempTarget.Activate
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^s"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Function
Function removeDoubleParagraphs(docTempTarget As Document) 'removes double paragraphs together
    docTempTarget.Activate
    
    With Selection.Find
        .Text = "^13{2,}"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Function

Function removeDoubleTabs(docTempTarget As Document) 'removes double paragraphs together
    docTempTarget.Activate
    
    With Selection.Find
        .Text = "^t{2,}"
        .Replacement.Text = "^t"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Function

Sub createStyles(docTempTarget)                   ' Create the custom styles in the target document.

    With docTempTarget.Styles                     ' Access the Styles gallery.

    If Not styleExists("Slide Text", docTempTarget) Then ' Create the "Slide Text" style to apply to the slide notes cell.
    .Add Name:="Slide Text"
    With docTempTarget.Styles("Slide Text")
        .AutomaticallyUpdate = False
        With .Font
            .Size = 11
            .Bold = False
            .Italic = False
        End With                                  ' Stop working with the font.
    End With                                      ' Stop working with the Slide Number style.
End If                                            'style not exists

If Not styleExists("Slide Number", docTempTarget) Then ' Create the "Slide Number" style to apply to the slide notes cells.
    .Add Name:="Slide Number"                     ' Create the style.
    With docTempTarget.Styles("Slide Number")     ' Define the style.
        .AutomaticallyUpdate = False
        .NextParagraphStyle = "Slide Text"
        With .Font
            .Name = "+Body"
            .Size = 11
            .Bold = True
            .Italic = True
        End With                                  ' Stop working with the font.
    End With                                      ' Stop working with the Slide Number style.
End If                                            'style not exists

If Not styleExists("Slide Image", docTempTarget) Then ' Create the "Slide Image" style to apply to the slide images. TODO unused in this version
    .Add Name:="Slide Image"
    With docTempTarget.Styles("Slide Image")
        .AutomaticallyUpdate = False
        With .Font
            .Size = 11
            .Bold = False
            .Italic = False
        End With                                  ' Stop working with the font.
    End With                                      ' Stop working with the Slide Number style.
End If                                            'style not exists

End With                                          ' Stop working with docTempTarget.Styles

End Sub
Function styleExists(ByVal styleToTest As String, ByVal docToTest As Word.Document) As Boolean
    Dim testStyle                                 As Word.Style
    On Error Resume Next
    Set testStyle = docToTest.Styles(styleToTest)
    styleExists = Not testStyle Is Nothing
End Function
Sub statusOutput(strMessage As String)
    ' Outputs to debug window, Status Bar, and optionally a message box
    Debug.Print strMessage
    Application.StatusBar = strMessage            ' TODO make sure this outputs to output doc instead of macro doc
    ' MsgBox (strMessage) ' uncomment out this line to also see a message box of the message
    
End Sub

Function createTable(docTempTarget As Document, dblRowHeight As Double) As Table

    Set createTable = docTempTarget.Tables.Add(Range:=Selection.Range, NumRows:=1, NumColumns:=1, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed)
    With createTable
        .Rows.HeightRule = wdRowHeightAtLeast
        .Rows.Height = InchesToPoints(dblRowHeight)
        .PreferredWidthType = wdPreferredWidthPercent
        .PreferredWidth = 100
    End With                                      ' Stop formatting the table, and select it.
    
End Function

Function goEnd(docTempTarget As Document)
    ' Just go to the end of the document, clear of any previous content.
    docTempTarget.Activate
    With Selection
        .MoveEnd Unit:=wdStory                    ' Get clear of any content
        .Start = .End
    End With
End Function



Function updateDocumentVariable(docTempTarget As Document, varName As String, varValue As String)
    Dim aVar   As Variable
    Dim aVarIndex As Integer

    For Each aVar In docTempTarget.Variables
        If aVar.Name = varName Then aVarIndex = aVar.Index
    Next aVar
    If aVarIndex = 0 Then
        docTempTarget.Variables.Add Name:=varName, Value:=varValue
    Else
        docTempTarget.Variables(aVarIndex).Value = varValue
        
    End If
    
End Function

Function readDocumentVariable(docTempTarget As Document, varName As String, varDefaultValue As String) As String
    Dim aVar   As Variable
    Dim aVarIndex As Integer

    For Each aVar In docTempTarget.Variables
        If aVar.Name = varName Then aVarIndex = aVar.Index
    Next aVar
    If aVarIndex = 0 Then
        docTempTarget.Variables.Add Name:=varName, Value:=varDefaultValue
        readDocumentVariable = varDefaultValue
    Else
        readDocumentVariable = docTempTarget.Variables(aVarIndex).Value
        
    End If
    
End Function
