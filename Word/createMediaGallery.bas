Option Explicit

Sub createMediaGallery()
    If (MsgBox("THIS WILL REALLY MESS UP YOUR POWERPOINT. Images will be reformatted And everything will be ungrouped. SAVE IT FIRST And DONT SAVE IT AFTERWARDS. Are you ready To run it now?", (vbYesNo + vbQuestion), "WARNING?") = vbYes) Then
        Dim docTempTarget                             As Document
        Set docTempTarget = wordDocumentSelection() ' Select word doc to output to
        Call createMediaTable(docTempTarget) ' main functionality
        If (MsgBox("All Done. Do you want To save this document at this time?", (vbYesNo + vbQuestion), "Save?") = vbYes) Then
            docTempTarget.Save ' save the document
        End If
        If (MsgBox("Do you want To run it again?", (vbYesNo + vbQuestion), "Again?") = vbYes) Then
            Call updateDocumentVariable(docTempTarget, "RunningAgain", "true")
            Call createMediaGallery ' call this very function if repeating
        End If
    End If
End Sub

Function wordDocumentSelection() As Document
    Dim docs                                      As Documents
    Set docs = Documents
    Dim doc         As Document
    Dim intDocSelect As Integer
    Dim docCount    As Integer:     docCount = 0
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
    Dim docCount    As Integer:     docCount = 0
    Dim strDocSelect As String
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
End Function

Private Function readPowerpointFile() As Object
    Dim objPPT                                    As Object
    Set objPPT = CreateObject("PowerPoint.Application")        ' Create and initialize the PowerPoint application object.
    With objPPT
        .Activate        ' Activate the PPT application object.
        .Visible = True        ' Make it visible.
        On Error GoTo failCleanly        ' Trap any errors and abort gracefully if the presentation file fails to open ... and remain open.
        Dim dlgOpen                               As FileDialog        ' Show the Open File dialog.
        Set dlgOpen = .FileDialog(Type:=msoFileDialogOpen)
        With dlgOpen
            .Title = "Select an Input Course File"
            .InitialFileName = Environ("USERPROFILE") & "\Downloads\"        ' Set the default directory path.
            .Show
            .Execute
        End With
        Set readPowerpointFile = .ActivePresentation        ' Set the active presentation to be the source file.
    End With        ' End initializing the PowerPoint object.
    statusOutput "PowerPoint Opening Completed"
    Exit Function
failCleanly:
    MsgBox "We had some trouble reading the PowerPoint file. To continue, try re-running the macro. If that still does Not work try closing Word completely And restarting it before rerunning the macro again.", Buttons:=vbExclamation, Title:="PowerPoint Had Trouble"
    End
    With objPPT
        .Activate        ' Activate PowerPoint
        .Quit        ' Exit PowerPoint
    End With
    Exit Function
End Function

Function createMediaTable(docTempTarget As Document)
    Dim tblGuide                                  As Table
    Dim strAltText  As String
    Dim sld                                       As Object
    Dim shp                                       As Object
    Dim dblImgWidth                               As Double: dblImgWidth = 1
    Dim dblMaxHeight As Double: dblMaxHeight = 4
    Dim objSrcFile                                    As Object        'PPT File
    'Dim intSlideNumber As Integer: intSlideNumber = 0
    'Dim dblRowHeight As Double
    Set objSrcFile = promptPowerpointFile()        'Open Powerpoint PPT file
    Dim boolShowSldNumb As Boolean: boolShowSldNumb = False 'include slide number in citation column
    Dim boolShowAutoAltText As Boolean: boolShowAutoAltText = False 'use alt-text if it contains the word "generated"
    Dim intCount    As Integer:    intCount = 0 'number of images exported
    Dim intErrors   As Integer: intErrors = 0 'number of times copy and paste failed
    Dim groupsExist As Boolean: groupsExist = True 'used for ungrouping
    Dim intGroupCount As Integer: intGroupCount = 0 'number of groups ungrouped
    Dim intErrorMax As Integer: intErrorMax = 5 'on error try this many times to paste
    If (MsgBox("Do you want To display slide number With source info?", (vbYesNo + vbQuestion), "Slide Number?") = vbYes) Then
        boolShowSldNumb = True 'prompt for dispaying slide number in citation column
    End If
    If (MsgBox("Do you want To show auto-generated Alt-Text (probably No)?", (vbYesNo + vbQuestion), "Slide Number?") = vbYes) Then
        boolShowAutoAltText = True 'prompt to show alt-text if it contains generated
    End If
    docTempTarget.Activate
    With Selection        'Begin new table
                .MoveEnd Unit:=wdStory        ' Get clear of any content
                .Start = .End        ' move to the end
                .TypeText ("[TODO Course Title] Media Gallery") 'create document heading
                .Style = docTempTarget.Styles("Heading 1") 'make previous a heading
                .TypeText (vbCrLf) 'new line
                .ClearFormatting 'make new line not a heading
                .MoveEnd Unit:=wdStory        ' Get clear of any content
                .Start = .End
                Set tblGuide = docTempTarget.Tables.Add(Range:=Selection.Range, NumRows:=1, NumColumns:=2, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed)
                With tblGuide
                    .PreferredWidthType = wdPreferredWidthPercent
                    .PreferredWidth = 100
                    .TopPadding = 0.1 * 72
                    .BottomPadding = 0.1 * 72
                    .Columns(1).PreferredWidth = 80
                    .Columns(2).Cells.VerticalAlignment = wdCellAlignVerticalCenter
                End With
                tblGuide.Cell(1, 1).Select
                .End = .Start
                .TypeText ("Image")
                tblGuide.Cell(1, 2).Select
                .End = .Start
                .TypeText ("Source Information")
            End With
            
    With objSrcFile
        For Each sld In .Slides
            sld.Select
            Do While (groupsExist = True) 'ungroup all shapes on each slide. keep looping to find nested groups
                groupsExist = False
                For Each shp In sld.Shapes
                    If shp.Type = msoGroup Then
                        shp.Ungroup
                        intGroupCount = intGroupCount + 1
                        groupsExist = True
                    End If
                Next shp
            Loop
            groupsExist = True 'reset for next slide
            Application.StatusBar = "Checking images On " & sld.slidenumber & " of " & objSrcFile.Slides.Count
           ' intSlideNumber = intSlideNumber + 1
           
    With docTempTarget
        .Activate
        With Selection
            For Each shp In sld.Shapes
                Debug.Print "Slide: " & sld.slidenumber & " Shape Name: " & shp.Name
                If shp.Type = 13 Then 'image type
                    intCount = intCount + 1
                    .MoveRight Unit:=wdCell, Extend:=wdMove        ' Move right again to create the next line in the table if it isnt the first slide
                    .End = .Start
                    On Error GoTo ErrorHandler        ' Enable error-handling routine. for copy and paste errors
                    shp.LockAspectRatio = True 'lock aspect ratio
                    shp.Width = InchesToPoints(dblImgWidth) 'set to default image width
                    If shp.Height > InchesToPoints(dblMaxHeight) Then        'the shape it too tall and can be less than 1 inch wide
                    shp.Height = InchesToPoints(dblMaxHeight) 'reset image height to max height
                End If
                shp.Rotation = 0 'reset orientation of the image
                shp.Copy
BeforePaste:
                .Paste
                intErrorMax = 5 'reset error count
                .MoveRight Unit:=wdCell, Count:=1 'move to the right column
                If boolShowSldNumb = True Then
                    strAltText = "Slide " & sld.slidenumber & vbCrLf
                End If
                If InStr(shp.AlternativeText, "generated") And boolShowAutoAltText = False Then
                    .TypeText (strAltText)
                Else
                    strAltText = strAltText & shp.AlternativeText
                    .TypeText (strAltText)
                End If
            End If
            strAltText = ""
        Next shp
    End With        'selection of temptarget
End With        'docTempTarget
Next sld
End With        'objSrcFile
'   objSrcFile.Close
'End If
For Each shp In ActiveDocument.InlineShapes 'reset shape borders and boxes
    With shp
        .Reset
        '.LockAspectRatio = msoCTrue        ' Lock the aspect ratio.
    End With
Next shp
With tblGuide
    .Rows.HeadingFormat = False        'removes all rows as a heading
    .Rows(1).HeadingFormat = True        'says there is a heading on row 1
    .Rows(1).Range.Paragraphs.Alignment = wdAlignParagraphCenter        'center row 1 text
    .Rows(1).Range.Bold = True        'bold row 1 text
    .Columns(1).Select 'select column 1 of images
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter 'align center
    .PreferredWidthType = wdPreferredWidthPercent '
    .PreferredWidth = 100
    .TopPadding = 0.1 * 72
    .BottomPadding = 0.1 * 72
    .Columns(1).PreferredWidth = 80
    .Columns(2).Cells.VerticalAlignment = wdCellAlignVerticalCenter
End With
With ActiveWindow
    .View.ShowHeading 2
    .DocumentMap = True
End With
MsgBox intCount & " Images Exported With " & intErrors & " errors Or images skipped, And " & intGroupCount & " groups ungrouped. REMEMBER, DONT SAVE THE PPT FILE."
Exit Function
ErrorHandler:
intErrorMax = intErrorMax - 1
intErrors = intErrors + 1
If intErrorMax > 0 Then
    GoTo BeforePaste
Else
    Resume Next
End If
End Function
