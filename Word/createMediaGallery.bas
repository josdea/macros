Option Explicit
Sub createMediaGallery()
    statusOutput "**********STARTING GUIDE**********"
    
    Dim docTempTarget                             As Document
    Set docTempTarget = wordDocumentSelection()
    
    Call createMediaTable(docTempTarget)
    
    If (MsgBox("All Done. Do you want To save at this time?", (vbYesNo + vbQuestion), "Save?") = vbYes) Then
        docTempTarget.Save
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
        .Visible = TRUE        ' Make it visible.
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
    Dim shpThisImg                                As Word.Shape
    Dim shp                                       As Object
    Dim dblImgWidth                               As Double
    Dim objSrcFile                                    As Object        'PPT File
    Dim intSlideNumber As Integer: intSlideNumber = 0
    Dim dblRowHeight As Double
    Set objSrcFile = promptPowerpointFile()        'Open Powerpoint PPT file
    Dim boolShowSldNumb As Boolean: boolShowSldNumb = FALSE
    Dim boolShowAutoAltText As Boolean: boolShowAutoAltText = FALSE
    
    If (MsgBox("Do you want To display slide number With source info?", (vbYesNo + vbQuestion), "Slide Number?") = vbYes) Then
        boolShowSldNumb = TRUE
    End If
    If (MsgBox("Do you want To show auto-generated Alt-Text?", (vbYesNo + vbQuestion), "Slide Number?") = vbYes) Then
        boolShowAutoAltText = TRUE
    End If
    
    docTempTarget.Activate
    
    dblRowHeight = 0.5
    
    dblImgWidth = 1
    
    With objSrcFile
        For Each sld In .slides
            Application.StatusBar = "Checking images On " & sld.slidenumber & " of " & objSrcFile.slides.Count
            
            Call completionBar(sld.slidenumber, objSrcFile.slides.Count)
            intSlideNumber = intSlideNumber + 1
            
            If sld.slidenumber = 1 Then        'so look to first slide for title
            
            With Selection        'Begin new table
                .MoveEnd Unit:=wdStory        ' Get clear of any content
                .Start = .End        ' move to the end
                
                .TypeText ("[TODO Course Title] Media Gallery")
                
                .Style = docTempTarget.Styles("Heading 1")
                .TypeText (vbCrLf)
                .ClearFormatting
                .MoveEnd Unit:=wdStory        ' Get clear of any content
                .Start = .End
                
                Set tblGuide = docTempTarget.Tables.Add(Range:=Selection.Range, NumRows:=1, NumColumns:=2, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed)
                With tblGuide
                    .PreferredWidthType = wdPreferredWidthPercent
                    .PreferredWidth = 100
                    .TopPadding = 0.1 * 72
                    .BottomPadding = 0.1 * 72
                    .PreferredWidthType = wdPreferredWidthPercent
                    .Columns(1).PreferredWidth = 80
                    .Columns(2).Cells.VerticalAlignment = wdCellAlignVerticalCenter
                    .Columns(1).Select
                    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
                End With
                tblGuide.Cell(1, 1).Select
                .End = .Start
                .TypeText ("Image")
                tblGuide.Cell(1, 2).Select
                .End = .Start
                .TypeText ("Source Information")
            End With
        ElseIf sld.slidenumber = .slides.Count Then        'last slide in presentaiton
        
    End If        'slide 1
    
    With docTempTarget
        .Activate
        With Selection
            
            For Each shp In sld.Shapes
                If shp.Type = 13 Then
                    .MoveRight Unit:=wdCell, Extend:=wdMove        ' Move right again to create the next line in the table if it isnt the first slide
                    .End = .Start
                    shp.Copy
                    .Paste
                    .MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                    With .InlineShapes(1)
                        .LockAspectRatio = msoCTrue        ' Lock the aspect ratio.
                        .Reset
                        .Width = 72
                    End With        ' Stop working with the pasted image shape.
                    
                    .MoveRight Unit:=wdCell, Count:=1
                    
                    If boolShowSldNumb = TRUE Then
                        strAltText = "Slide " & sld.slidenumber & vbCrLf
                        
                    End If
                    
                    If InStr(shp.AlternativeText, "generated") And boolShowAutoAltText = FALSE Then
                        
                        .TypeText (strAltText)
                    Else
                        strAltText = strAltText & shp.AlternativeText
                        .TypeText (strAltText)
                    End If
                    
                End If
            Next shp
            
        End With        'selection of temptarget
        
    End With        'docTempTarget
Next sld
End With        'objSrcFile
With ActiveWindow
    .View.ShowHeading 2
    .DocumentMap = TRUE
End With
End Function