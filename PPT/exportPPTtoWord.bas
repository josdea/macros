Option Explicit

Sub createGuides()
    statusOutput "Starting the export process"
    
    ' declare and create new word doc
    Dim docTempTarget                             As Document
    Set docTempTarget = createNewWordDocument()
    
    ' create word styles in new document TODO have this called from the section that needs it only
    'Call createStyles(docTempTarget)
    
    ' declare and open powerpoint file
    Dim objSrcFile                                As Object
    Set objSrcFile = readPowerpointFile(docTempTarget)
    
    Call createMetaData(docTempTarget, objSrcFile)
    Call createInstructorGuide(docTempTarget, objSrcFile)
    'Call exportSlideImages(docTempTarget, objSrcFile)
    
    MsgBox ("All Done")
End Sub
Function createNewWordDocument() As Document
    statusOutput "Creating New Word Document"
    
    Set createNewWordDocument = Application.Documents.Add
    
    statusOutput createNewWordDocument.Name & " Created Successfully"
End Function

Function readPowerpointFile(docTempTarget As Document) As Object
    statusOutput "Open And read from Powerpoint File"
    Dim objPPT                                    As Object
    Set objPPT = CreateObject("PowerPoint.Application")        ' Create and initialize the PowerPoint application object.
    With objPPT
        .Activate        ' Activate the PPT application object.
        .Visible = True        ' Make it visible.
        
        ' Trap any errors and abort gracefully if the presentation file fails to open ... and remain open.
        On Error GoTo failCleanly
        
        ' Show the Open File dialog.
        Dim dlgOpen                               As FileDialog
        Set dlgOpen = .FileDialog(Type:=msoFileDialogOpen)
        With dlgOpen
            .Title = "Select an Input Course File"
            .InitialFileName = Environ("USERPROFILE") & "\Desktop\"        ' Set the default directory path.
            .Show
            .Execute
        End With
        .Activate
        Set readPowerpointFile = .ActivePresentation        ' Set the active presentation to be the source file.
    End With        ' End initializing the PowerPoint object.
    statusOutput "PowerPoint Opening Completed"
    
    Exit Function
    ' Error trap for subroutine's On Error statement above.
failCleanly:
    ' Trap errors. Alert the user that PPT closed unexpected.
    ' Debug.Print "Error in makeModuleList: " & "-Error(" & Error(ErrorNumber) & "): " & Error(ErrorNumber)
    MsgBox "PowerPoint quit unexpectedly before we could read its file content. To continue, try re-running the macro after we discard the New MS-Word file. If this problem persists, try closing Word completely And restarting it before rerunning the macro again. Press OK To continue.", Buttons:=vbExclamation, Title:="PowerPoint Quit Unexpectedly"
    
    With objPPT
        .Activate        ' Activate PowerPoint
        .Quit        ' Exit PowerPoint
    End With
    docTempTarget.Close SaveChanges:=wdDoNotSaveChanges
    Exit Function
End Function

Function createMetaData(docTempTarget As Document, objSrcFile As Object)
    docTempTarget.Activate
    
    With Selection
        .EndKey Unit:=wdStory
        
        ' Add target doc title
        .TypeText ("Temp Guide Content")
        .Style = "Title"
        .InsertParagraph
        .EndKey Unit:=wdStory
        
        ' Add metadata
        ActiveDocument.Bookmarks.Add Name:="metadata", Range:=.Range        ' Create the metadata bookmark
        .TypeText ("Metadata")        ' Metadata heading
        .Style = "Heading 1"
        .InsertParagraph
        .EndKey Unit:=wdStory
        .TypeText ("Course Title: " & "Present Title TODO")        ' Course Title, from Slide 1
        .InsertParagraph
        .EndKey Unit:=wdStory
        .TypeText ("File Title: " & objSrcFile.BuiltInDocumentProperties.Item(1).Value)        'File Title property
        .InsertParagraph
        .EndKey Unit:=wdStory
        .TypeText ("File Name: " & objSrcFile.Name)        ' Filename property
        .InsertParagraph
        .EndKey Unit:=wdStory
        .TypeText ("Slide Count: " & objSrcFile.Slides.Count)        ' Slidecount property
        .InsertParagraph
        .EndKey Unit:=wdStory
        .TypeText ("Guide Created: " & DateTime.Time & " " & DateTime.Date)        ' Slidecount property
        '.InsertParagraph
        '.EndKey Unit:=wdStory
        
    End With
    
End Function

Function exportSlideImages(docTempTarget As Document, objSrcFile As Object)
    Dim fld                                       As Field
    Dim sld                                       As Object
    Dim objSrcFilePath                            As String
    objSrcFilePath = objSrcFile.Path & "\" & objSrcFile.Name
    objSrcFilePath = Replace(objSrcFilePath, "\", "\\")
    
    For Each sld In objSrcFile.Slides
        docTempTarget.Activate
        
        With Selection
            .InsertParagraph
            .EndKey Unit:=wdStory
            Set fld = docTempTarget.Fields.Add(Range:=Selection.Range, Type:=wdFieldLink, Text:="PowerPoint.Slide.8 " & objSrcFilePath & "!" & sld.slideid)
        End With
        
        With fld.InlineShape
            .ScaleWidth = 70
            .ScaleHeight = 70
            .AlternativeText = sld.slidenumber
            If sld.slidenumber > 10 Then
                End
            End If
            
        End With
        
    Next sld
    
End Function
Function createInstructorGuide(docTempTarget As Document, objSrcFile As Object)
    Dim tblGuide                                  As Table
    Dim intRowHeight                              As Integer: intRowHeight = 4        'INCHES
    
    docTempTarget.Activate
    With Selection
        .InsertParagraph
        .EndKey Unit:=wdStory
        
        ' Add a heading
        docTempTarget.Bookmarks.Add Name:="instructorguide", Range:=.Range        ' Create the metadata bookmark
        
        .TypeText ("Instructor Guide")
        .Style = docTempTarget.Styles("Heading 1")
        .InsertParagraph
        .EndKey Unit:=wdStory
        
        Set tblGuide = docTempTarget.Tables.Add(Range:=.Range, NumRows:=1, NumColumns:=1, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitWindow)
        With tblGuide
            .Style = "Table Grid"
            .ApplyStyleHeadingRows = True
            .ApplyStyleLastRow = False
            .ApplyStyleFirstColumn = True
            .ApplyStyleLastColumn = False
            .ApplyStyleRowBands = True
            .ApplyStyleColumnBands = False
            .Select
            ' Set Row height rules and table width
            .Rows.HeightRule = wdRowHeightAtLeast
            .Rows.Height = InchesToPoints(intRowHeight)
            .PreferredWidthType = wdPreferredWidthPercent
            .PreferredWidth = 100
            .Cell(1, 1).Select
        End With        'tblGuide
        .End = .Start
        
       Dim fld                                       As Field
    Dim sld                                       As Object
    Dim objSrcFilePath                            As String
    objSrcFilePath = objSrcFile.Path & "\" & objSrcFile.Name
    objSrcFilePath = Replace(objSrcFilePath, "\", "\\")
    
    For Each sld In objSrcFile.Slides
        docTempTarget.Activate
        Application.StatusBar = "Adding slide " & sld.slidenumber
        
        'With Selection
            '.InsertParagraph
            '.EndKey Unit:=wdStory
        
          .TypeText ("Slide " & sld.slidenumber & ": ")   ' Add the slide number.
            .InsertParagraph
            .EndKey Unit:=wdStory
            .Cells(1).Select
            .End = .Start
            
            Set fld = docTempTarget.Fields.Add(Range:=Selection.Range, Type:=wdFieldLink, Text:="PowerPoint.Slide.8 " & objSrcFilePath & "!" & sld.slideid)
        'End With
        
        With fld.InlineShape
            .ScaleWidth = 70
            .ScaleHeight = 70
            .AlternativeText = sld.slidenumber
         '   If sld.slidenumber > 10 Then
         '       End
         '   End If
            
        End With
        
           ' Move right to end of the cell.
                            .Cells(1).Select
                            .Start = .End
                            .MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
        
    Next sld
        
    End With        'Selection
    
End Function
