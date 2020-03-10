Option Explicit

' INS Training Guide Preparation and Publication Process
' ver. 2.0
' Developed by Erik Faye, Atkins Global NS (ejfaye@sandia.gov, 505-998-9322)
'
' Change History
'   Date     Version   Reason
'===============================================================================================
'   11/15/2018 2.0     Revised for template update, including recognizing new slide layouts. Add image alt text to media catalog. Resolve minor issues (e.g., table formatting).
'   01/24/2018 1.0     Original release.
'

    '   Declare global variables
    
    ' General variables
    Dim strTempImgDir As String     ' The directory path for the temporary image directory, into which the slide images are exported.
    Dim inErrorState As Boolean     ' Create a boolean to determine if there's an error.
    
    ' Word objects
    Dim docTempTarget As Document   ' The document that will hold the temporary values.  Final documents are then generated from the source content held here.
    Dim tblModList As Table         ' The table that lists the various modules. Iterated backward to determine where to add module section headings in makeModuleSections().
    Dim tblPartGuide As Table       ' The notes table of the Participant Guide.
    Dim tblInstrGuide As Table      ' The notes table of the Instructor Guide.
    Dim tblMediaCat As Table        ' The notes table of the Media catalog.
    Dim tblAtaGlance As Table       ' The Module At-A-Glance table.
    Dim tblModObjectives As Table   ' The Module Objectives table.
    Dim tblAcronyms As Table        ' The Adjectives list table.
    
    ' PowerPoint objects
    Dim objPPT                      ' The PowerPoint application object.
    Dim objSrcFile                  ' The source PowerPoint File.
    Dim strCurrPath As String       ' The current directory of the PPT file. Used for storing the image files, etc.
    Dim sldSlides                   ' The range of slide objects.
    Dim sldSlide                    ' The current slide object
    Dim intSlideCnt As Integer      ' The number of slides in the source presentation.
    Dim intSlideNum As Integer      ' The slide number of the current slide. Redefined/valued within For Each Slide loops.
    Dim strPresentTitle As String   ' The title property of the presentation.
    
    ' Image parameters
    Dim strExportFormat As String   ' The file format of the exported images (e.g., gif, jpg, png)
    Dim strThisImgPath As String    ' The import path of each image.
    Dim intScaleWidth As Integer    ' The pixel width of the exported images.
    Dim intScaleHeight As Integer   ' The pixel width of the exported images.
    Dim intImprtScale As Integer    ' The percentage of the
    Dim intBorderWidth As Integer   ' The border width of the images.
    
    ' String objects. varBadWords() is dim'ed in cleanNoteTables(), but the count has to be a constant dimmed here as a private global.
    Private Const cnstBadWordCnt = 22     ' Create a constant of the size of the array.

Sub guideInit()
' Initialize and coordinate the processes.
' Start everything by running this and it'll call all needed, dependent subroutines.
    
    Debug.Print "Starting the process."
    
    ' Fill global variables.
    strTempImgDir = "\tempPptExport"    ' A temporary holding directory within the presentation's active directory for the exported images.
    strExportFormat = "gif"     ' The file format/filter of the exported images. This can also be png, jpg, bmp, or others, but GIF works for the file size.

' *** COMMENT OUT THE THUMBNAIL IMAGE SIZE YOU DON'T WANT TO USE ***
' Half-wide Images
'    intScaleWidth = 600         ' Width in pixels.
'    intScaleHeight = 450        ' Height in pixels.
'    intImprtScale = 65          ' The percent scale of the image on import. Used in makePartGuideNotes() and makeInstGuideNotes().
' Two-thirds Images
    intScaleWidth = 700         ' Width in pixels.
    intScaleHeight = 526        ' Height in pixels.
    intImprtScale = 55          ' The percent scale of the image on import. Used in makePartGuideNotes() and makeInstGuideNotes().
    
    intBorderWidth = 0.5        ' The width in points of the border of the imported image. Used by the same subroutines.
    inErrorState = False        ' Initialize the macro error state as false. Error traps shift to true to avoid redundant alerts to users.
    
    ' NOTE: The varBadWords(i) list, which includes those strings to be deleted from the imported Notes text, is dimmed in cleanNoteTables().
    
    ' >> UNCOMMENT ON PUBLICATION >> On Error GoTo quitProcess
    
    ' Create the target temporary document.
    Set docTempTarget = Application.Documents.Add
    
    ' Create the custom styles in the temporary document.
    Call createStyles
    
    ' Open PowerPoint and make the metadata section in the target document.
    Call makeMetaData
    
    ' Build the Module Listing table
    Call makeModuleList      ' Assumes that the selection is at the end of the story.
    
    ' Build the Module Objectives table
    Call makeModuleObjs      ' Assumes that the selection is at the end of the story.
    
    ' Build the At-a-Glance table
    Call makeAtaGlance      ' Assumes that the selection is at the end of the story.
    
    ' Export the slide images to the temp image directory.
    Debug.Print "Exporting slide images from " & objSrcFile.Name
    With objSrcFile
        .Export Path:=strTempImgDir, FilterName:=strExportFormat, ScaleWidth:=intScaleWidth, ScaleHeight:=intScaleHeight
    End With
    
    docTempTarget.Activate  '    Reactivatite the temp target file.
    
    ' Build the Module Image Notes tables
    Call makePartGuideNotes      ' Make the Participant Guide table by calling the makePartGuideNotes(). Assumes that the selection is at the end of the story.
    Call makeInstrctGuideNotes   ' Make the Instructor Guide table by calling the makeInstGuideNotes(). Assumes that the selection is at the end of the story.
    
    ' Make the media catalog with the images pasted into the presentation.
    Call makeMediaCatalog
    
    ' Do some housecleaning on the imported Notes tables.
    Call cleanNoteTables            ' Delete defined strings (e.g., "Slide Purpose:")
    Call RemoveBlankParasLoop       ' Delete empty paragraphs, namely those of 1 or fewer characters (e.g., empty or having just a stray space).
    
    ' With the text cleaned up -- including deleting empty paragraphs -- style the note text in the instructor guide.
    Call styleInstrGuide
    
    ' Split the two tables into Module sections by creating headings from each module's number and title.
    Call makeModuleSections
    
    ' Compile the acronyms into the Acronym section, which is between the At-a-Glance and Module sections.
    Call makeAcronyms
    
    Call goStart        ' Move to start of the document
    
    ' Clean up and close out the final output documents.
    Call guideFinish
    
    ' Shut down and quit PowerPoint
    Debug.Print "Closing and Exiting PowerPoint."
    With objPPT
        .Activate               ' Activate PowerPoint
        .Quit                   ' Exit PowerPoint
    End With
    
    ' Clear any remaining variables.
    Set dlgOpen = Nothing       ' Release the Open Presentation dialog object.
    Set objSrcFile = Nothing    ' Clear the presentation object.
    Set objPPT = Nothing        ' Clear the PPT object.
    
    Debug.Print "Done creating the master Guide file."
    Debug.Print " -- #### -- "
    
    Call goStart        ' Move to start of the document
    
  Exit Sub
  
quitProcess:
    
    Debug.Print " -- Prematurely Aborting guideInit() -- "
    Exit Sub

' End guideInit()
End Sub

Sub makeAcronyms()
' Compile the acronyms into a table between the At-a-Glance and Module sections, just before Participant or Instructor Guide module content.
    
    ' NOTE: Target and source tables are dimmed globally as tblInstrGuide (source table) and tblAcronyms (target table).
    
    Dim oRange As Range         ' The range to be searched for acronyms. This is the contents of the instructor guide table.
    Dim n As Long               ' The number of acronyms found, which drives the number of rows in the target table.
    Dim strAcronym As String    ' Each acronym that's found.
    Dim strAllFound As String   ' This becomes a "wrapper" for the found acronym; if the result is just the point sign, no acronym was found.
    
    strAllFound = "#"
    
    Debug.Print "Compiling acronyms."
    
    ' Create a bookmark and the heading
    With Selection
        .GoTo What:=wdGoToBookmark, Name:="partguide"                   ' Go to the start of the Participant Guide.
        .InsertParagraphBefore                                          ' Insert a new paragraph before it.
        docTempTarget.Bookmarks.Add Name:="acronyms", Range:=.Range     ' Create the acronym list bookmark.
        .TypeText ("Acronyms")                      ' Add the heading text.
        .Style = "Heading 1"                        ' Style it.
        .InsertParagraphAfter                       ' Create a new paragraph.
        .MoveDown Unit:=wdLine, Count:=1            ' Move down a line to the new paragraph.
        .Start = .End                               ' Collapse the selection.
        .Style = "Normal"                           ' Style it as Normal.
    End With        ' Stop working with the selection.
    
    ' Create the target table.
    Set tblAcronyms = docTempTarget.Tables.Add(Range:=Selection.Range, NumRows:=2, NumColumns:=2)
    With tblAcronyms    ' Start working with the Acronyms table.
        .Cell(1, 1).Range.Text = "Acronym"          ' Create the column headings.
        .Cell(1, 2).Range.Text = "Definition"
        .Rows(1).HeadingFormat = True
        .Rows(1).Range.Font.Bold = True
        .PreferredWidthType = wdPreferredWidthPercent
        .Columns(1).PreferredWidth = 30             ' Set the column widths.
        .Columns(2).PreferredWidth = 70
    End With            ' Stop working with the Acronyms table.
    
    ' Search the text in the instructor guide
    With tblInstrGuide      ' Start working with the Instructor Guide table, where the relevant Acronyms are to be found.
        Set oRange = .Range
        n = 1
        With oRange.Find
            .Text = "<[A-Z][A-Z\-&/]{1,}>"
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = True
            .MatchWildcards = True
            Do While .Execute
                strAcronym = oRange
                If InStr(1, strAllFound, "#" & strAcronym & "#") = 0 Then
                    If n > 1 Then tblAcronyms.Rows.Add
                    strAllFound = strAllFound & strAcronym & "#"
                    With tblAcronyms
                        .Cell(n + 1, 1).Range.Text = strAcronym
                    End With
                    n = n + 1
                End If
            Loop
        End With    ' Stop working with the search range.
    End With    ' Stop working with the instructor guide table.
    
    ' Sort and format the resulting table
    With tblAcronyms    ' Start working with the Acronyms table.
        .Sort ExcludeHeader:=True, FieldNumber:="Column 1", SortFieldType:=wdSortFieldAlphanumeric, SortOrder:=wdSortOrderAscending
        ' .Rows.Range.Font.Bold = False
        .Rows(1).Range.Font.Bold = True
    End With            ' Stop working with the Acronyms table.
    
    Debug.Print "Finished compiling " & n - 1 & " acronyms to the listing"
    
' End makeAcronyms()
End Sub

Sub makeModuleSections()
'
' Loop backward through the Modules table and split the Instructor and Participant Guides into sections.
' Do this by going to the row of the title slides, insert a new row, convert it to text, then add the styled module name and number.
' Note that Module 1, found in tblModList row 2, needs to act differently as it also includes the course title slide.
'
    
    Debug.Print "Making the module sections."
    
    ' Declare the variables.
    Dim intRowNum As Integer
    Dim strModTitle As String
    
    ' Get the number of rows.
    intRowNum = tblModList.Rows.Count       ' By using the number of rows, we iterate from the end of the table to the start to preserve row positions.
    
    ' Loop the module list's rows, from the last and moving up, but skipping Row 1, which is the column heading row.
    While intRowNum > 1      ' Begin iterating down the rows, but not the first row.
        
        ' Get the last row's Module Title and Slide Number
        strModTitle = Trim(tblModList.Rows(intRowNum).Cells(2).Range.Text)
        intSlideNum = Int(tblModList.Rows(intRowNum).Cells(3).Range.Words.First)
        
        Debug.Print intRowNum & " Slide " & intSlideNum & " " & strModTitle
        
        ' Go to that slide's row that's numbered the same as the slide number. Do in both the Instructor and Participant Guide tables.
        With tblInstrGuide.Rows(intSlideNum)    ' Start with the Instructor Guide table
            ' Insert a new row above, select it and convert to text.
            .Select                         ' Select the slidenumber's row (module title slide)
            With Selection                  ' Work with the selection
                If intSlideNum = 2 Then     ' Move up a row if this is the first module slide, always on row 2
                    tblInstrGuide.Rows(1).Select  ' Select Row 1, instead of Row 2.
                End If
                .InsertRowsAbove NumRows:=1 ' Insert a row, which moves the selection
                .Rows(1).ConvertToText      ' Convert the selected row to text.
                .End = .Start               ' Collapse the insertion point.
                .TypeText (strModTitle)     ' Type the Module title.
                .MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend ' Select the cell debris
                .Delete                     ' Delete the debris.
                .Style = "Heading 1"   ' Style the remaining paragraph as Heading 1.
            End With ' Stop working with the selection.
            ' Type the Module Title, then style the paragraph.
        End With ' Stop working with the slide table row.
        
        With tblPartGuide.Rows(intSlideNum)    ' Then process the Participant Guide table
            ' Insert a new row above, select it and convert to text.
            .Select                         ' Select the slidenumber's row (module title slide)
            With Selection                  ' Work with the selection
                If intSlideNum = 2 Then     ' Move up a row if this is the first module slide, always on row 2
                    tblPartGuide.Rows(1).Select  ' Select Row 1, instead of Row 2.
                End If
                .InsertRowsAbove NumRows:=1 ' Insert a row, which moves the selection
                .Rows(1).ConvertToText      ' Convert the selected row to text.
                .End = .Start               ' Collapse the insertion point.
                .TypeText (strModTitle)     ' Type the Module title.
                .MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend ' Select the cell debris
                .Delete                     ' Delete the debris.
                .Style = "Heading 1"   ' Style the remaining paragraph as Heading 1.
            End With ' Stop working with the selection.
            ' Type the Module Title, then style the paragraph.
        End With ' Stop working with the slide table row.
        
        ' Move to the next tblModList row up (subtract 1 from intRowNum) and repeat
        intRowNum = intRowNum - 1   ' Subtract 1 from intRowNum, so that it's now the number of the next row up.
    Wend    ' Stop this row and move to the next row up.
    
' End makeModuleSections()
End Sub

Sub RemoveBlankParasLoop()
'
' This loops through each paragraph, replacing those that are empty, meaning it has a length of 2 (the paragraph mark and possibly a leading space or other debris).
'
    
    Dim para As Paragraph
    Debug.Print "Removing Blank Paragraphs from " & ActiveDocument.Name
    
    ' Iterate all paragraphs.
    For Each para In ActiveDocument.Paragraphs
        On Error Resume Next
        ' Delete any blank or almost blank paragraphs.
        If Len(para.Range.Text) <= 2 Then
            'Only the paragraph mark, or maybe a stray space, so..
            para.Range.Delete   ' Delete it.
        End If
        
        ' Delete the leading space if there's one.
        If para.Range.Characters(1) = " " Then
            para.Range.Characters(1).Delete
        End If
    Next para
    
' End RemoveBlankParasLoop()
End Sub

Sub cleanNoteTables()
' Remove known unneeded or sloppy content, such as the Slide Purpose labels in the notes cells.
    
    ' Create a counter to iterate the array slots, deleting each from the tempTargetDoc.
    Dim i As Integer
    i = 0
    
    ' Create the list of bad words.
    Dim varBadWords(cnstBadWordCnt) As String    ' The array of strings to remove. Dim'ed in guideInit(). The array size is dimmed as the cnstBadWordCnt globally.
    
    ' Fill the bad word array.
    ' When adding terms, increase the cnstBadWordCnt constant in the module's global Declarations to ensure that enough array slots are available.
    varBadWords(0) = "Slide Purpose:"       ' Note labels
    varBadWords(1) = "Instructor Notes:"
    varBadWords(2) = "Bathrooms, Emergency Procedures, and Breaks:"
    varBadWords(3) = "Course Purpose & Brief Facilitator Introduction:"
    varBadWords(4) = "Student Notes:"
    varBadWords(5) = "No notes"
    varBadWords(6) = "Slide Objective (Main Point):"
    varBadWords(7) = "Summary slide"
    varBadWords(8) = "Image credit:"
    varBadWords(9) = "Images"
    varBadWords(10) = "References:"
    varBadWords(11) = "Miscellaneous: "
    varBadWords(12) = "N/A"
    varBadWords(13) = "How does slide relate to the weekly topics learning objectives?:"
    varBadWords(14) = "References: "
    varBadWords(15) = "Outline slide"
    varBadWords(16) = "Image credits:"
    varBadWords(17) = "Purpose:"
    varBadWords(18) = "TALKING POINTS:"
    varBadWords(19) = "TRANSITION:"
    varBadWords(20) = "Slide Purpose"
    varBadWords(21) = "Instructor Notes"
    
    ' Find all the bad words in the document and replace them with nothing.
    While i < cnstBadWordCnt                    '   iterate the array of bad words and phrases.
        With Selection
            .Find.ClearFormatting
            .Find.Replacement.ClearFormatting
            With .Find
                .Text = varBadWords(i)
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = True
                .MatchWholeWord = True
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            .Find.Execute Replace:=wdReplaceAll
            .Find.ClearFormatting
            .Find.Replacement.ClearFormatting
        End With
        i = i + 1 ' increase the iterator
    Wend    ' Stop iterating varBadWords(i).
    
    ' Remove the last paragraph/pilcrow in the note tables.
    For Each Row In tblInstrGuide.Rows
        For Each Cell In Row.Cells          ' Loop the Cells
            Cell.Select                 ' Select the first cell.
            With Selection              ' Work with the selection.
                .Start = .End           ' collapse to the end of the selection.
                .TypeBackspace          ' Backspace to delete the stray paragraph.
            End With    ' Stop working with the selection.
        Next Cell       ' Move on to the next cell.
    Next Row        ' Move on to the next row.
    
' End cleanNoteTables()
End Sub

Sub makeMediaCatalog()
' Transfer images from the slides to a table to create the media catalog.
'
    
    Dim strAltText As String
    
    Debug.Print "Making the Media catalog."
    
    ' Create the target table for the Media Catalog.
    With Selection
        
        Call goEnd  ' Move to the end and collapse.
        
        ' Add a heading
        docTempTarget.Bookmarks.Add Name:="mediacatalog", Range:=.Range    ' Create the metadata bookmark
        
        .TypeText ("Media Catalog")
        .Style = docTempTarget.Styles("Heading 1")
        .TypeText (vbCrLf)
        .ClearFormatting
        .TypeText ("Note that the following media catalog might be incomplete. Double check it against the processed input PowerPoint file to ensure that all grouped photographs are included here. Each image from such groups should be listed separately.")
        .TypeText (vbCrLf)
        
        Call goEnd  ' Move to the end and collapse
        
        ' Create and format the target table.
        Set tblMediaCat = docTempTarget.Tables.Add(Range:=.Range, NumRows:=1, NumColumns:=3, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed)
        With tblMediaCat
            If .Style <> "Table Grid" Then
                .Style = "Table Grid"
            End If
            .ApplyStyleHeadingRows = True
            .ApplyStyleLastRow = False
            .ApplyStyleFirstColumn = True
            .ApplyStyleLastColumn = False
            .ApplyStyleRowBands = True
            .ApplyStyleColumnBands = False
            .Select
            .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
            .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
            .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
            .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
            .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
            
            ' Make the heading row
            .Cell(1, 1).Range.Text = "Image"          ' Create the column headings.
            .Cell(1, 2).Range.Text = "Slide"
            .Cell(1, 3).Range.Text = "Source"
            .Rows(1).HeadingFormat = True
            .Rows(1).Range.Font.Bold = True
            .PreferredWidthType = wdPreferredWidthPercent
            .Columns(1).PreferredWidth = 25             ' Set the column widths.
            .Columns(2).PreferredWidth = 25
            .Columns(3).PreferredWidth = 50
            
            ' Create and go to the next row to begin receiving images.
            .Cell(1, 3).Select
            Selection.MoveRight Unit:=wdCell, Extend:=wdMove
            Selection.End = Selection.Start
            
        End With    ' Stop formatting and selecting the table.
        
        ' Go to the start of the table and collapse.
        tblMediaCat.Cell(2, 1).Select
        .End = .Start
        
        ' Iterate through the Slides
        With objPPT
            .Activate                       ' Activate the PPT application object.
            
            With objSrcFile
                
                ' Trap any errors and abort gracefully if the presentation file fails to open ... and remain open.
'                 On Error GoTo keepGoing
                
                intSlideNum = 0                     ' Initialize the slide number to zero; it gets incremented at the start of each new slide.
                strThisImgPath = ""
                
                Set sldSlides = .Slides             ' Connect to the slide stack
                
                ' Flip through the slides.
                For Each sldSlide In sldSlides                 ' Iterate each slide in the range presentaiton object.
                
                    ' Get the slide number.
                    intSlideNum = intSlideNum + 1           ' Increment the slide number from the previous value
                    Application.StatusBar = "Checking slide " & intSlideNum & " images"
                    
                    ' Ungroup any groups of objects.
                    For Each objPict In sldSlide.Shapes         ' Loop through the shapes on the slide
                        If objPict.Type = 6 Then                ' Check for grouped objects
                            objPict.Ungroup                     ' Ungroup if so.
                        End If      ' End checking for groups.
                    Next objPict    ' Move to the next shape.
                    
                    ' For each picture that may be found....
                    For Each objPict In sldSlide.Shapes
                        
                        If objPict.Type = 13 Then                ' If this is a valid picture type, copy it to the clipboard (13 is a Object Type).
                            
                            ' Get the image's Alt Text
                            ' strAltText = "Boolah"
                            strAltText = objPict.AlternativeText
                            
                            ' Copy it to the clipboard,
                            objPict.Copy
                            
                            ' Paste it into the target table's cell(n,1)
                            With docTempTarget
                                With Selection
                                    .Paste
                                    
                                    ' Lock aspect ratio and resize and reset image.
                                    .MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                                    With .InlineShapes(1)
                                        .LockAspectRatio = msoCTrue ' Lock the aspect ratio.
                                        .ScaleWidth = 25           ' Scale to 1" width, leaving height as-is.
                                    End With    ' Stop working with the pasted image shape.
                                    
                                    .MoveRight Unit:=wdCell, Count:=1
                                    ' Add slide number to cell(n,2)
                                    .TypeText (intSlideNum)
                                    ' Move to create a next cell.
                                    .MoveRight Unit:=wdCell, Count:=1
                                    ' Paste AltText
                                    .TypeText (strAltText)
                                    ' Move to create a new row.
                                    .MoveRight Unit:=wdCell, Count:=1
                                End With
                            End With
                            
                          ' Else
                            ' Debug.Print "False " & objPict.Type
                        End If ' Stop testing for whether the object is a picture.
                        
                    Next objPict    ' Stop looping images.
                        
                Next sldSlide      ' Move on to the next slide.
                
            End With    ' Stop working with the source PPT file.
        End With    ' Stop working with the PPT application.
        
        ' Remove the last, trailing row of the table; it's a leftover from looping the slides.
        Call trimLastRow        ' Strip the last, trailing row of the table.
        
        ' Format the table
        With tblMediaCat    ' Start working with the Acronyms table.
            .Range.Font.Bold = False
            .Rows(1).Range.Font.Bold = True
        End With            ' Stop working with the Acronyms table.
        
        Call goEnd           ' And collapse the selection to the end.
        
    End With    ' Stop working with the selection in the target document.
    
' End makeMediaCatalog()
End Sub

Sub makeInstrctGuideNotes()
' Capture the module chapters, with images, slide numbers, and notes.
'
' The following variables were DIM'd globally in guideInit().
' strTempImgDir = "\tempPptExport"
' intScaleWidth = 600       ' <- Note that these measures also have alternates pre-written in comments in guideInit().
' intScaleHeight = 450
' intImprtScale = 50
' intBorderWidth = 0.5
' strThisImgPath = null initially
    
    Debug.Print "Making the Instructor Guide table."
    
    strExportFormat = UCase(strExportFormat)    ' Upper case the extention because the export function makes it all-cap.
    
    ' Create the target table for the Slide Notes.
    With Selection
        
        Call goEnd  ' Move to the end and collapse.
        
        ' Add a heading
        docTempTarget.Bookmarks.Add Name:="instrctguide", Range:=.Range    ' Create the metadata bookmark
        
        .TypeText ("Instructor Guide")
        .Style = docTempTarget.Styles("Heading 1")
        .TypeText (vbCrLf)
        .ClearFormatting
        
        Call goEnd  ' Move to the end and collapse
        
        ' Create and format the target table.
        Set tblInstrGuide = docTempTarget.Tables.Add(Range:=.Range, NumRows:=1, NumColumns:=1, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed)
        With tblInstrGuide
            If .Style <> "Table Grid" Then
                .Style = "Table Grid"
            End If
            .ApplyStyleHeadingRows = True
            .ApplyStyleLastRow = False
            .ApplyStyleFirstColumn = True
            .ApplyStyleLastColumn = False
            .ApplyStyleRowBands = True
            .ApplyStyleColumnBands = False
            .Select
            
            ' Set Row height rules and table width
            .Rows.HeightRule = wdRowHeightAtLeast
            .Rows.Height = InchesToPoints(4)
            .PreferredWidthType = wdPreferredWidthPercent
            .PreferredWidth = 100
        End With    ' Stop formatting the table, and select it.
        
        ' Go to the start of the table and collapse.
        tblInstrGuide.Cell(1, 1).Select
        .End = .Start
        
        ' Iterate through the Slides
        With objPPT
            .Activate                       ' Activate the PPT application object.
            
            With objSrcFile
                
                ' Trap any errors and abort gracefully if the presentation file fails to open ... and remain open.
                On Error GoTo keepGoing
                
                intSlideNum = 0                     ' Initialize the slide number to zero; it gets incremented at the start of each new slide.
                strThisImgPath = ""
                
                Set sldSlides = .Slides             ' Connect to the slide stack
                
                ' Flip through the slides.
                For Each sldSlide In sldSlides                 ' Iterate each slide in the range presentaiton object.
                    
                    ' Get slide number.
                    intSlideNum = intSlideNum + 1           ' Increment the slide number from the previous value
                    Application.StatusBar = "Adding slide " & intSlideNum
                    
                    ' Add content to the Temporary Target Document
                    With docTempTarget
                        .Activate
                        With Selection
                            
                            ' Insert and style the Page Number text.
                            .TypeText ("Slide " & intSlideNum & ": ")   ' Add the slide number.
                            .Style = "Slide Number"                     ' Set the style.
                            
                            ' Move the cursor to the start of the slide number before importing the image. Otherwise, the anchor will be at the end of the cell contents.
                            .Cells(1).Select
                            .End = .Start
                            
                            ' Make the image path, insert the image, and format it.
                            strThisImgPath = strTempImgDir & "\Slide" & intSlideNum & "." & strExportFormat   ' Concatenate the path to the slide images.
                            ' Import the slide image.
                            .InlineShapes.AddPicture FileName:=strThisImgPath, LinkToFile:=False, SaveWithDocument:=True ' Images are previously dimensioned to Width:=600, Height:=450
                            ' Scale and border the image.
                            .Cells(1).Select
                            
                            ' Format the image's word wrap.
                            Call wrapImage
                            
                            ' Move right to end of the cell.
                            .Cells(1).Select
                            .Start = .End
                            .MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
                            
                            ' Copy any Notes to the clipboard. These will be pasted into the target document a few lines later.
                            For Each oSh In sldSlides(intSlideNum).NotesPage.Shapes    ' Get the shapes in the Notes page of this particular slide, identified with intSlideNum.
                                ' >> STOPPED WORKING FOR SOME REASON >> If oSh.PlaceholderFormat.Type = ppPlaceholderBody Then 'If there's a body to the shape (ppPlaceholderBody = 2).
                                If oSh.PlaceholderFormat.Type = 2 Then 'If there's a body to the shape (ppPlaceholderBody = 2).
                                    If oSh.HasTextFrame Then    ' Check that a text frame exists in the body.
                                        If oSh.TextFrame.HasText Then   ' Check for text in that text frame.
                                        
                                            ' NOTE: The following line hangs occasionally. Simply press the Debug button on error, then continue running by pressing F5.
                                            oSh.TextFrame.TextRange.Copy    ' Copy the Note text to the clipboard. It'll be pasted into the current Slide Notes cell in a few lines, when we return to the table.
                                        
                                        End If      ' End text in the text frame test.
                                    End If      ' End Body text frame test.
                                End If      ' End Placeholder Body test.
                            Next oSh    ' Move to the next Notes Page shape object.
                            
                            .TypeText (vbCrLf)          ' Add a return to get away from the slide number
                            .Paste                      ' Paste the clipboard's content
                            
                            ' Clear the clipboard by typing a space, selecting it, then cutting it to the clipboard, overriding previous content.
                            .TypeText (" ")
                            .MoveStart Unit:=wdCharacter, Count:=-1
                            .Cut
                            
                            ' Move right to create the next row.
                            .MoveRight Unit:=wdCell, Extend:=wdMove     ' Move right again to create the next line.
                            
                            End With    ' Stop working with the selection.
                        End With     ' Stop working with the selection in the target document.
                Next sldSlide    ' Move to the next slide.
            End With    ' Stop working with the source presentation file.
        End With    ' Stop working with PowerPoint
        
        ' Return focus to the target temp table.
        docTempTarget.Activate  '    Reactivatite the temp target file.
        
        ' Remove the internal vertical border from the table
        .Tables(1).Select ' Select the row's contents.
        
        ' Remove the table's borders. This is pointless here as the bordering of the image also re-borders the cell.
        Call setTableFormat
        
        .Start = .End
        .MoveUp Unit:=wdLine, Extend:=wdExtend
        
        ' Remove the last, trailing row of the table; it's a leftover from looping the slides.
        Call trimLastRow        ' Strip the last, trailing row of the table.
        
        Call goEnd           ' And collapse the selection to the end.
        
    End With ' Stop writing to the selection point.
    
    ' Quit this subroutine because we're done.
    Exit Sub
    
' Error trap for subroutine's On Error statement above.
keepGoing:
    
    ' Trap errors. Alert the user that PPT closed unexpected.
    Debug.Print "makeInstrGuideNotes: " & "-Error(" & Error(ErrorNumber) & ") on Slide " & intSlideNum
    If inErrorState = False Then
        Debug.Print " -#- Entering Error Mode. -#-"
        MsgBox "PowerPoint had trouble reading content from the Slide " & intSlideNum & ". This is usually a temporary and recoverable hiccup. However, check the resulting file when it's complete to ensure that everything worked as expected. Press OK to continue.", Buttons:=vbExclamation, Title:="PowerPoint Stumbled"
    End If
    inErrorState = True
    Resume
    
' End makeInstrctGuideNotes()
End Sub

Sub makePartGuideNotes()
' Capture the module chapters, with images, slide number and notes
    
' The following variables are globally DIM'd in guideInit().
' strTempImgDir = "\tempPptExport"
' intScaleWidth = 600       ' <- Note that these measures also have alternates pre-written in comments in guideInit().
' intScaleHeight = 450
' intImprtScale = 50
' intBorderWidth = 0.5
' strThisImgPath = null initially
    
    Debug.Print "Making the Participant Guide table."
    
    strExportFormat = UCase(strExportFormat)     ' Reset the extention because the export function makes it all-cap.
    
    ' Create the target table for the Slide Notes.
    With Selection
        
        Call goEnd  ' Move to the end of the document and collapse.
        
        ' add a heading
        docTempTarget.Bookmarks.Add Name:="partguide", Range:=.Range    ' Create the metadata bookmark
        
        .TypeText ("Participant Guide")
        .Style = docTempTarget.Styles("Heading 1")
        .TypeText (vbCrLf)
        .ClearFormatting
        
        Call goEnd  ' Move to the end and collapse
        
        ' Create and style the target table
        Set tblPartGuide = docTempTarget.Tables.Add(Range:=.Range, NumRows:=1, NumColumns:=1, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed)
        
        With tblPartGuide
            If .Style <> "Table Grid" Then
                .Style = "Table Grid"
            End If
            .ApplyStyleHeadingRows = True
            .ApplyStyleLastRow = False
            .ApplyStyleFirstColumn = True
            .ApplyStyleLastColumn = False
            .ApplyStyleRowBands = True
            .ApplyStyleColumnBands = False
            .Select
            
            ' Set Row height rules and table width
            .Rows.HeightRule = wdRowHeightAtLeast
            .Rows.Height = InchesToPoints(4)
            .PreferredWidthType = wdPreferredWidthPercent
            .PreferredWidth = 100
        End With    ' Stop formatting the table, and select it.
        
        ' Go to the start of the table and collapse.
        tblPartGuide.Cell(1, 1).Select
        .End = .Start
        
        ' Iterate through the Slides
        With objPPT
            .Activate                       ' Activate the PPT application object.
            
            With objSrcFile
                
                intSlideNum = 0                     ' Initialize the slide number to zero; it gets incremented at the start of each new slide.
                strThisImgPath = ""
                
                Set sldSlides = .Slides             ' Connect to the slide stack
                
                ' Flip through the slides.
                For Each sldSlide In sldSlides              ' Iterate each slide in the range presentation object.
                    
                    ' Get slide number.
                    intSlideNum = intSlideNum + 1           ' Increment the slide number from the previous value
                    Application.StatusBar = "Adding slide " & intSlideNum
                    ' Add content to the Temporary Target Document
                    With docTempTarget
                        .Activate
                        With Selection
                            
                            ' Insert and style the Page Number text.
                            .TypeText ("Slide " & intSlideNum & ": ")   ' Add the slide number.
                            .Style = "Slide Number"                     ' Set the style.
                            
                            ' Move the cursor to the start of the slide number before importing the image. Otherwise, the anchor will be at the end of the cell contents.
                            .Cells(1).Select
                            .End = .Start
                            
                            ' Make the image path, insert the image, and format it.
                            strThisImgPath = strTempImgDir & "\Slide" & intSlideNum & "." & strExportFormat   ' Concatenate the path to the slide images.
                            ' Perform the import.
                            .InlineShapes.AddPicture FileName:=strThisImgPath, LinkToFile:=False, SaveWithDocument:=True ' Images are previously dimensioned to Width:=600, Height:=450
                            ' Scale and border the image.
                            .Cells(1).Select
                            
                            ' Format the image's word wrap.
                            Call wrapImage
                            
                            ' Move right to start of the cell.
                            .Cells(1).Select
                            .End = .Start
                            
                            ' Move right to create the next row.
                            .MoveRight Unit:=wdCell, Extend:=wdMove     ' Move right again to create the next line.
                            
                            End With    ' Stop working with the selection.
                        End With       ' Stop working with the target document.
                Next sldSlide    ' Move to the next slide.
            End With    ' Stop working with the source presentation file.
        End With    ' Stop working with PowerPoint
        
        ' Return focus to the target temp table.
        docTempTarget.Activate  ' Reactivatite the temp target file.
        
        ' Remove the internal vertical border from the table
        .Tables(1).Select ' Select the row's contents.
        ' Remove the table's borders. This is pointless here as the bordering of the image also re-borders the cell.
        Call setTableFormat
        
        .Start = .End
        .MoveUp Unit:=wdLine, Extend:=wdExtend
        
        ' Remove the last, trailing row of the table; it's a leftover from looping the slides.
        Call trimLastRow     ' Strip the last, trailing row of the table.
        
        Call goEnd           ' And collapse the selection to the end.
        
    End With ' Stop writing to the selection point.
    
' End makePartGuideNotes()
End Sub

Sub makeAtaGlance()
' Capture the module chapters, with slide number and notes.
    
    Debug.Print "Making the At-a-Glance table."
    
    ' Create the target table for the Slide Notes.
    With Selection
        
        Call goEnd  ' Move to the end and collapse.
        
        ' Add a heading.
        docTempTarget.Bookmarks.Add Name:="ataglance", Range:=.Range    ' Create the metadata bookmark.
        
        .TypeText ("At-A-Glance Instructor's Guide")
        .Style = docTempTarget.Styles("Heading 1")
        .TypeText (vbCrLf)
        .ClearFormatting
        
        Call goEnd  ' Move to the end of the document and collapse
        
        ' Create the target table
        Set tblAtaGlance = docTempTarget.Tables.Add(Range:=.Range, NumRows:=1, NumColumns:=1)
        
        With tblAtaGlance
            If .Style <> "Table Grid" Then
                .Style = "Table Grid"
            End If
            .ApplyStyleHeadingRows = True
            .ApplyStyleLastRow = False
            .ApplyStyleFirstColumn = True
            .ApplyStyleLastColumn = False
            .ApplyStyleRowBands = True
            .ApplyStyleColumnBands = False
        End With    ' Stop iterating slides
        
        ' Add the Column Heading
        tblAtaGlance.Cell(1, 1).Select
        .End = .Start
        .TypeText ("Body of Instruction:")
        .MoveRight Unit:=wdCell ' Move to start a new row.
        
        ' Iterate through the Slides
        With objPPT
            .Activate                       ' Activate the PPT application object.
            .Visible = True                 ' Make it visible.
            
            With objSrcFile
                
                Set sldSlides = .Slides             ' Connect to the slide stack
                
                ' Iterate through the slides, checking for titles and adding content to the table.
                ' Set a module number counter.
                Dim intModNum As Integer
                intModNum = 0
                
                intSlideNum = 0
                Dim strModTitle As String
                Dim strModSubtitle As String
                
                ' Flip through the slides.
                For Each sldSlide In sldSlides                  ' Iterate each slide in the range presentaiton object.
                    intSlideNum = sldSlide.SlideNumber          ' Get this slide number
                    If sldSlide.Layout = 1 And intSlideNum <> 1 Then       ' If this is a title slide, but not the first, which is the course title, get the title (module number) and subtitle (module subject)
                        intModNum = intModNum + 1
                        strModTitle = Trim(sldSlide.Shapes.Title.TextFrame.TextRange.Text)
                        strModSubtitle = Trim(sldSlide.Shapes.Item(2).TextFrame.TextRange.Text)
                        
                        ' Output the slide number and title text.
                        With docTempTarget
                            .Activate
                            With Selection
                                .TypeText ("Slide " & intSlideNum & ": " & strModTitle & " " & Chr(151) & " " & strModSubtitle)
                                .TypeBackspace
                                .MoveRight Unit:=wdCell, Extend:=wdMove
                            End With    ' Stop working with the selection.
                        End With     ' Stop working with the target document.
                        
                      ' And output the individual slide numbers and headings.
                      Else
                        strModTitle = Trim(sldSlide.Shapes.Title.TextFrame.TextRange.Text)
                        With docTempTarget
                            .Activate
                            With Selection
                                .TypeText ("Slide " & intSlideNum & ": " & strModTitle)
                                .MoveRight Unit:=wdCell, Extend:=wdMove
                            End With    ' Stop working with the selection.
                        End With    ' Stop working with the document.
                        
                    End If
                Next sldSlide    ' Move to the next slide.
            End With    ' Stop working with the source presentation file.
        End With    ' Stop working with PowerPoint
        
        ' Return focus to the target temp table.
        docTempTarget.Activate  '    Reactivatite the temp target file.
        
        ' Remove the last, trailing row of the table; it's a leftover from looping the slides.
        Call trimLastRow        ' Strip the last, trailing row of the table.
        
        Call goEnd           ' And collapse the selection to the end.
        
    End With ' Stop writing to the selection point.
    
' End makeAtaGlance()
End Sub

Sub makeModuleObjs()
' Capture the module titles and objectives

    Debug.Print "Making the Module Objectives table."
    
    ' Create the target table for the Slide Notes.
    With Selection
        
        Call goEnd  ' Move to the end and collapse
        
        ' add a bookmark and heading
        docTempTarget.Bookmarks.Add Name:="objectives", Range:=.Range    ' Create the metadata bookmark
        
        .TypeText ("Module Objectives")
        .Style = docTempTarget.Styles("Heading 1")
        .TypeText (vbCrLf)
        .ClearFormatting
        
        Call goEnd  ' Move to the end and collapse
        
        ' Create the target table
        Set tblModObjectives = docTempTarget.Tables.Add(Range:=.Range, NumRows:=1, NumColumns:=1)
        
        With tblModObjectives
            If .Style <> "Table Grid" Then
                .Style = "Table Grid"
            End If
            .ApplyStyleHeadingRows = True
            .ApplyStyleLastRow = False
            .ApplyStyleFirstColumn = True
            .ApplyStyleLastColumn = False
            .ApplyStyleRowBands = True
            .ApplyStyleColumnBands = False
        End With    ' Stop iterating slides.
        
        ' Add Column Heading
        tblModObjectives.Cell(1, 1).Select
        .End = .Start
        .TypeText ("Workshop Learning Objectives & Assessments")
        .MoveRight Unit:=wdCell ' Move to start a new row.
        .End = .Start
        
        ' Iterate through the Slides
        With objPPT
            .Activate                       ' Activate the PPT application object.
            .Visible = True                 ' Make it visible.
            
            With objSrcFile
                
                ' Trap any errors and abort gracefully if the presentation file fails to open ... and remain open.
                On Error GoTo keepItGoing
                
                Set sldSlides = .Slides             ' Connect to the slide stack
                
                ' Iterate through the slides, checking for titles and adding content to the table.
                ' Set a module number counter.
                Dim intModNum As Integer        ' This module's number
                intModNum = 0                   ' Initialize as zero.
                ' Dim intSlideNum As Integer      ' This slide's slide nubmer.
                Dim intObjctvsNum As Integer    ' The slide number for the Module Objectives, typically the slide immediately following the module title.
                Dim strModTitle As String       ' The title of the module (e.g., "Module 14")
                Dim strModSubtitle As String    ' The subtitle of the module (e.g., "Course Summary"
                Dim strObjectives As String     ' The text block of the objectives
                
                ' Flip through the slides.
                For Each sldSlide In sldSlides                 ' Iterate each slide in the range presentaiton object.
                    intSlideNum = sldSlide.SlideNumber          ' Get this slide's number.
                    ' NOTE: The first tests for the module title slides, but the first else-if tests for the slide that follows captures the Objectives.
                    If sldSlide.CustomLayout.Name = "Course Module Title" And intSlideNum <> 1 Then
                        intModNum = intModNum + 1                   ' Increment the module number.
                        intObjctvsNum = intSlideNum + 1             ' Infer the next slide's number to find the Objectives.
                        strModTitle = Trim(sldSlide.Shapes.Title.TextFrame.TextRange.Text)
                        strModSubtitle = Trim(sldSlide.Shapes.Item(2).TextFrame.TextRange.Text)
                        Debug.Print "--Module "; intModNum
                        
                        ' Get the Objectives text on the next slide by flagging its slide number.
                        With docTempTarget
                            .Activate
                            Selection.TypeText (strModTitle & ": " & strModSubtitle)
                            Selection.InsertParagraph          ' Add a return so that the objectives will be on a new line.
                            Selection.MoveRight Unit:=wdCharacter, Count:=1 ' Move right one step.
                            ' The new table row in docTempTarget will be created after placing the objectives
                        End With    ' Stop writing out to the docTempTarget document's selection point.
                        
                      ' However, if this is the slide immediately following a module title, get the objectives from the body.
                      ' NOTE: Also check for Objectives layout, otherwise unexpected content will cause infinite loop failure.
                      ElseIf sldSlide.SlideNumber = intObjctvsNum And sldSlide.CustomLayout.Name = "Objectives and Summary" Then
                        ' We're on the Objectives page.
                        Debug.Print "Objectives on "; sldSlide.SlideNumber
                        
                        ' sldSlide.Shapes(2).TextFrame.TextRange.Copy         ' GET THE OBJECTIVES LIST USING COPY AND PASTE. >> This is the previous design. Check object #3, below.
                        sldSlide.Shapes(3).TextFrame.TextRange.Copy         ' GET THE OBJECTIVES LIST USING COPY AND PASTE.
                        
                        With docTempTarget
                            With Selection
                                .Paste
                                
                                ' Select remainder of the cell.
                                .Cells(1).Select
                                
                                ' Clear and reset the bullets as a numbered list.
                                .ClearFormatting
                                
                                .Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
                                    ListGalleries(wdNumberGallery).ListTemplates(1), ContinuePreviousList:= _
                                    False, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
                                    wdWord10ListBehavior
                                
                                ' Clear bullet from first paragraph.
                                .MoveLeft Unit:=wdCharacter, Count:=1
                                .ClearFormatting
                                
                                ' Go to the end of the cell and backspace to remove end pilcrow and straw number.
                                .Cells(1).Select
                                .Start = .End
                                .MoveLeft Unit:=wdCharacter, Count:=1
                                .TypeBackspace
                                
                                ' Clear the clipboard
                                .TypeText (" ")
                                .MoveStart Unit:=wdCharacter, Count:=-1
                                .Cut
                                .MoveRight Unit:=wdCell, Extend:=wdMove    ' Move to the next cell, creating a new row.
                            End With    ' Stop with the selection.
                        End With        ' Stop with the target document.
                        
                      ' >> THE FOLLOWING IS JUST A PLACEHOLDER FOR REFACTORING
                      ' Else
                      '     (Do something else.)
                    End If
                Next sldSlide    ' Move to the next slide.
            End With    ' Stop working with the source presentation file.
        End With    ' Stop working with PowerPoint application.
        
        ' Return focus to the target temp table.
        docTempTarget.Activate  ' Reactivatite the temp target file.
        
        ' Remove the last, trailing row of the table; it's a leftover from looping the slides.
        Call trimLastRow        ' Deletes the last, trailing row.
        
        Call goEnd              ' And collapse the selection to the end.
        
    End With ' stop writing to the selection point.
    
    ' Quit this subroutine because we're done.
    Exit Sub
    
' Error trap for subroutine's On Error statement above.
keepItGoing:
    
    ' Trap errors. Alert the user that PPT closed unexpected.
    Debug.Print "makeModuleObjs: " & "-Error(" & Error(ErrorNumber)
    If inErrorState = False Then
        Debug.Print " -#- Entering Error Mode | makeModuleObjs() -#-"
        MsgBox "PowerPoint had trouble reading content from the Module " & intModNum & ". This is usually a temporary recoverable hiccup. However, check the resulting file when it's complete to ensure that everything worked as expected. Press OK to continue.", Buttons:=vbExclamation, Title:="PowerPoint Stumbled"
    End If
    inErrorState = True
    Resume
    
' End makeModuleObjs()
End Sub

Sub makeMetaData()
' Capture the presentation's metadata.
    
    Debug.Print "Making the Meta Data section."
    ' Open and set the target PowerPoint file object.
    Set objPPT = CreateObject("PowerPoint.Application")  ' Create and initialize the PowerPoint application object.
    With objPPT
        .Activate                       ' Activate the PPT application object.
        .Visible = True                 ' Make it visible.
        
        ' Trap any errors and abort gracefully if the presentation file fails to open ... and remain open.
        On Error GoTo failCleanly
        
        ' Show the Open File dialog.
        Dim dlgOpen As FileDialog
        Set dlgOpen = .FileDialog(Type:=msoFileDialogOpen)
        With dlgOpen
            .Title = "Select an Input Course File"
            .InitialFileName = Environ("USERPROFILE") & "\Documents\" ' Set the default directory path.
            .Show
            .Execute
        End With
        .Activate
        Set objSrcFile = .ActivePresentation    ' Set the active presentation to be the source file.
        
        ' Get the slide's directory path
        strCurrPath = objSrcFile.Path                   ' Get the active presentation's current file directory
        strTempImgDir = strCurrPath & strTempImgDir     ' Concatenate it to the temp export directory's path for exporting and importing slide images.
        
        Debug.Print "Outputting to " & strTempImgDir
        
        ' Get presentation information (title, count, etc.).
        intSlideCnt = objSrcFile.Slides.Count
        strPresentTitle = Trim(objSrcFile.Slides(1).Shapes.Item(1).TextFrame.TextRange.Text)
        
    End With    ' End initializing the PowerPoint object.
    
    ' Add presentation metadata information to the target document.
    docTempTarget.Activate
    With Selection
        
        Call goEnd  ' Get clear of any content
        
        ' Add target doc title
        .TypeText ("Temp Guide Content")
        .Style = "Title"
        .InsertParagraph
        Call goEnd          ' Call a process to move and collapse the selection to the end of the document.
        
        ' Add metadata
        ActiveDocument.Bookmarks.Add Name:="metadata", Range:=.Range    ' Create the metadata bookmark
        .TypeText ("Metadata")         ' Metadata heading
        .Style = "Heading 1"
        .InsertParagraph
        Call goEnd
        .TypeText ("Course Title: " & strPresentTitle)         ' Course Title, from Slide 1
        .InsertParagraph
        Call goEnd
        .TypeText ("File Title: " & objSrcFile.BuiltInDocumentProperties.Item(1).Value)   ' Title property
        .InsertParagraph
        Call goEnd
        .TypeText ("File Name: " & objSrcFile.Name)          ' Filename property
        .InsertParagraph
        Call goEnd
        .TypeText ("Slide Count: " & intSlideCnt)        ' Slidecount property
        .InsertParagraph
        Call goEnd
        .TypeText ("Guide Created: " & DateTime.Time & " " & DateTime.Date)         ' Slidecount property
        .InsertParagraph
        Call goEnd
    End With
    
    ' Exit sub
    Exit Sub
    
' Error trap for subroutine's On Error statement above.
failCleanly:
    
    ' Trap errors. Alert the user that PPT closed unexpected.
    Debug.Print "Error in makeModuleList: " & "-Error(" & Error(ErrorNumber) & "): " & Error(ErrorNumber)
    MsgBox "PowerPoint quit unexpectedly before we could read its file content. To continue, try re-running the macro after we discard the new MS-Word file. If this problem persists, try closing Word completely and restarting it before rerunning the macro again. Press OK to continue.", Buttons:=vbExclamation, Title:="PowerPoint Quit Unexpectedly"
    
    With objPPT
        .Activate               ' Activate PowerPoint
        .Quit                   ' Exit PowerPoint
    End With
    
    docTempTarget.Close SaveChanges:=wdDoNotSaveChanges
    
    Exit Sub
    
' End makeMetaData()
End Sub

Sub makeModuleList()
' Capture the module titles and numbers (mod and slide).
' NOTE: Refactor to keep the presentation file from closing suddenly. Once stable, rework the on error handling.
    
    Debug.Print "Making the Module Listing table."
    
    ' Create the target table for the Slide Notes.
    With Selection
        
        ' Trap any errors and abort gracefully if the presentation file fails to open ... and remain open.
        On Error GoTo failCleanly
        
        Call goEnd  ' Move to the end and collapse
        
        ' add a heading
        docTempTarget.Bookmarks.Add Name:="modulelist", Range:=.Range    ' Create the metadata bookmark
        
        .TypeText ("Module List")
        .Style = docTempTarget.Styles("Heading 1")
        .TypeText (vbCrLf)
        .ClearFormatting
        
        Call goEnd  ' Move to the end and collapse
        
        ' Create the target table
        Set tblModList = docTempTarget.Tables.Add(Range:=.Range, NumRows:=1, NumColumns:=3)
        
        With tblModList
            If .Style <> "Table Grid" Then
                .Style = "Table Grid"
            End If
            .ApplyStyleHeadingRows = True
            .ApplyStyleLastRow = False
            .ApplyStyleFirstColumn = True
            .ApplyStyleLastColumn = False
            .ApplyStyleRowBands = True
            .ApplyStyleColumnBands = False
            
            ' Set Column Widths
            .Columns.PreferredWidthType = wdPreferredWidthPercent
            .Columns(1).PreferredWidth = 20
            .Columns(2).PreferredWidth = 60
            .Columns(3).PreferredWidth = 20
            
            ' Add Column Headings
            .Cell(1, 1).Range.Text = "Module Number"
            .Cell(1, 2).Range.Text = "Module Title"
            .Cell(1, 3).Range.Text = "Slide Number"
            .Cell(1, 3).Select
        End With    ' Stop iterating slides
        
        .MoveRight Unit:=wdCell, Extend:=wdMove
        .End = .Start
        
        ' Iterate through the Slides
        With objPPT
            .Activate                       ' Activate the PPT application object.
            .Visible = True                 ' Make it visible.
            
            With objSrcFile
                
                Set sldSlides = .Slides             ' Connect to the slide stack
                
                ' Iterate through the slides, checking for titles and adding content to the table.
                ' Set a module number counter.
                Dim intModNum As Integer
                intModNum = 0
                Dim strModTitle As String
                Dim strModSubtitle As String
                ' Dim intSlideNum As Integer
                
                ' Flip through the slides.
                For Each sldSlide In sldSlides                 ' Iterate each slide in the range presentaiton object.
                    If sldSlide.CustomLayout.Name = "Course Module Title" And sldSlide.SlideNumber <> 1 Then       ' Is it the right layout and not the first (course title) slide? If so, proceed into loop.
                        intModNum = intModNum + 1
                        strModTitle = Trim(sldSlide.Shapes.Title.TextFrame.TextRange.Text)
                        strModSubtitle = Trim(sldSlide.Shapes.Item(2).TextFrame.TextRange.Text)
                        intSlideNum = sldSlide.SlideNumber
                        
                        With docTempTarget
                            .Activate
                            With Selection
                                .TypeText (intModNum)
                                .MoveRight Unit:=wdCell, Extend:=wdMove
                                ' USE THIS LINE FOR EM DASH SEPARATOR .TypeText (strModTitle & " " & Chr(151) & " " & strModSubtitle) ' The "Chr(151)" is the em-dash delimiter.
                                .TypeText (strModTitle & ": " & strModSubtitle) ' This uses a colon to separate.
                                .MoveRight Unit:=wdCell, Extend:=wdMove
                                .TypeText (intSlideNum)
                                .MoveRight Unit:=wdCell, Extend:=wdMove
                            End With        ' Stop with the selection.
                        End With            ' Stop with the target document.
                        
                      ' Keep this Else statement for refactoring, if alternate action is needed.
                      'Else
                      '  << Do something different. >>
                      
                    End If
                Next sldSlide    ' Move to the next slide.
            End With    ' Stop working with the source presentation file.
        End With    ' Stop working with PowerPoint
        
        ' Return focus to the target temp table.
        docTempTarget.Activate  ' Reactivatite the temp target file.
        
        ' Remove the last, trailing row of the table; it's a leftover from looping the slides.
        Call trimLastRow        ' Strip the last, trailing row of the table.
        
        Call goEnd              ' And collapse the selection to the end.
        
    End With ' Stop writing to the selection point.
  
  Exit Sub  ' Stop the subroutine now so we don't step into the error-trap below, which ironically causes errors to be thrown.

' Error trap for subroutine's On Error statement above.
failCleanly:
    
    ' Trap errors. Alert the user that PPT closed unexpected.
    Debug.Print "Error in makeModuleList: " & "-Error(" & Error(ErrorNumber) & "): " & Error(ErrorNumber)
    MsgBox "PowerPoint quit unexpectedly before we could read its file content. To continue, try re-running the macro after we close the new MS-Word file without saving. If this problem persists, try closing Word completely and restarting it before rerunning the macro again. Press OK to continue.", Buttons:=vbExclamation, Title:="PowerPoint Quit Unexpectedly"
    
    With objPPT
        .Activate               ' Activate PowerPoint
        .Quit                   ' Exit PowerPoint
    End With
    
    docTempTarget.Close SaveChanges:=wdDoNotSaveChanges
    
    Exit Sub
    
' End makeModuleList()
End Sub

Sub styleInstrGuide()
' Style the Instructor Guide table.
' For later refactoring, consider methods for preserving bullet and numbered lists.
    
    Debug.Print "Styling the Instructor Guide notes."
    
    For Each Row In tblInstrGuide.Rows   ' Loop the Rows
        For Each Cell In Row.Cells         ' Loop the Cells
            If Cell.ColumnIndex = 1 Then     ' Skip the first cell
                For Each para In Cell.Range.Paragraphs  ' Loop the paragraphs
                    If para.Style = "Slide Number" Then ' Skip the Slide Number paragraph.
                        ' Do nothing with the slide number.
                      ElseIf para.Range.Font.Italic = -1 Then   ' If it's italic, style it Recommended.
                        para.Style = "Recommended"
                      Else                              ' If it's neither Slide Number nor Italic, style it Slide Text.
                        para.Style = "Slide Text"
                    End If
                Next para   ' Stop looping paras
            End If      ' Stop skipping Cell 1
        Next Cell   ' Stop looping Cells.
    Next Row    ' Stop looping Rows.
    
' End styleInstrGuide()
End Sub

Sub guideFinish()
' Close out the processing of Word documents.

    Debug.Print "Cleaning up the final document."
    
    ' Show the Open File dialog.
    Dim dlgSaveAs As FileDialog
    Set dlgSaveAs = Application.FileDialog(FileDialogType:=msoFileDialogSaveAs)
    With dlgSaveAs
        .Title = "Save the Output Temporary Document"
        .InitialFileName = strCurrPath & "\TempCourseGuideContent" ' Set the default directory path.
        .Show
        .Execute
    End With
    
' End guideFinish()
End Sub

Sub pubInit()
' **** BEGIN THE GUIDE PUBLISHING PROCESS
' NOTE: This currently doesn't do anything, but it's coded into the process as an "off-ramp" for additional processing if needed in later refactoring.
' Initialize the process.
    
    Debug.Print "Intializing the process."
    
    Call pubFinish
    
' End pubInit()
End Sub

Sub pubFinish()
' Initialize the process to finalize the published documents.
' NOTE: This currently doesn't do anything, but it's coded into the process as an "off-ramp" for additional processing if needed in later refactoring.
' This should be fired from the single, final Guide document, once everything's correct.
    
    Debug.Print "Cleaning up final document."
    
' End pubFinish()
End Sub

Sub createStyles()
' Create the custom styles in the target document.
    
    ' Access the Styles gallery.
    With docTempTarget.Styles
    
        ' Create the "Slide Number" style to apply to the slide notes cells.
        .Add Name:="Slide Number"                   ' Create the style.
        With docTempTarget.Styles("Slide Number")   ' Define the style.
            .AutomaticallyUpdate = False
            With .Font
                .Name = "+Body"
                .Size = 11
                .Bold = True
                .Italic = True
            End With    ' Stop working with the font.
        End With    ' Stop working with the Slide Number style.
        
        ' Create the "Recommended" style to apply to the slide notes cell.
        .Add Name:="Recommended"
        With docTempTarget.Styles("Recommended")
            .AutomaticallyUpdate = False
            With .Font
                .Size = 11
                .Bold = False
                .Italic = True
            End With    ' Stop working with the font.
        End With    ' Stop working with the Slide Number style.
        
        ' Create the "Slide Text" style to apply to the slide notes cell.
        .Add Name:="Slide Text"
        With docTempTarget.Styles("Slide Text")
            .AutomaticallyUpdate = False
            With .Font
                .Size = 11
                .Bold = False
                .Italic = False
            End With    ' Stop working with the font.
        End With    ' Stop working with the Slide Number style.
        
    End With ' Stop working with docTempTarget.Styles
    
' End createStyles()
End Sub

Sub wrapImage()
'
' This wraps the placed image to the right
'

    With Selection
        ' Scale and convert the inline image.
        With .InlineShapes(1)               ' Scale it by the percentage dimmed above.
            .LockAspectRatio = msoTrue      ' Lock the aspect ratio
            .ScaleHeight = intImprtScale    ' Scale its width & height.
            .ScaleWidth = intImprtScale     ' This should be redundant, but be sure to preserve the aspect ratio.
            
            ' Add the border.
            .Borders.Enable = True          ' THIS NOR THE CODE BELOW DOES ANYTHING
            .Line.Weight = intBorderWidth   ' Set the line weight to a half point
            
            ' Shift it from being inline to float right.
            .ConvertToShape
            
        End With    ' Stop working with the inline shape.
        
        ' Wrap it flush right.
        With .Range.ShapeRange(1)
            .RelativeHorizontalPosition = wdRelativeHorizontalPositionColumn
            With .WrapFormat
                .Type = wdWrapSquare
                .Side = wdWrapLeft                  ' The "left" refers to where the words will wrap around, not where the image is located.
                .DistanceTop = InchesToPoints(0.1)
                .DistanceBottom = InchesToPoints(0.1)
                .DistanceLeft = InchesToPoints(0.1)
                .DistanceRight = InchesToPoints(0.1)
            End With    ' Stop wrap formatting.
            .LayoutInCell = True
            .Left = wdShapeRight     ' For right-aligned floats
            .AlternativeText = "Slide " & intSlideNum
        End With    ' stop working with the ShapeRange(1)
        
    End With    ' Stop working with the selection
End Sub

Sub trimLastRow()
    ' Delete the last row of the current table. Selection is presumed in cell(n,1) or thereabouts.
    
    Selection.Rows(1).Delete
    
' End trimLastRow()
End Sub
Sub setTableFormat()
'
' Format module slide tables.
' Called from both the instructor and participant guide making subroutines after the target tables have been created and populated.
'
    Selection.Tables(1).Select  ' Select the table
    
    ' Remove any borders in the selected table.
    With Selection.Tables(1)
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
    End With
    
    ' Set default border options.
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = -603930625
    End With
    
    ' Back with the selected table, add bottom and inside horizontal borders.
    ' Changing the defaults above is redundant with this, but it seems to work okay without significant inefficiencies.
    With Selection.Tables(1)
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
    End With
    
End Sub

Sub goEnd()
' Just go to the end of the document, clear of any previous content.

    docTempTarget.Activate
    
    With Selection
        .MoveEnd Unit:=wdStory  ' Get clear of any content
        .Start = .End
    End With
    
End Sub

Sub goStart()
' Just go to the start of the document, clear of any previous content.

    docTempTarget.Activate
    
    With Selection
        .MoveStart Unit:=wdStory  ' Get clear of any content
        .End = .Start             ' Collapse the selection's end to the start.
    End With
    
End Sub

Sub A_Acronymchecker()
'
' A_Acronymchecker Macro
' Finds all acronyms in a target document and dumps them into a table in a new document.
'

Dim oDoc_Source As Document
Dim oDoc_Target As Document
Dim strListSep As String
Dim strAcronym As String
Dim oTable As Table
Dim oRange As Range
Dim n As Long
Dim strAllFound As String

strAllFound = "#"

Set oDoc_Source = ActiveDocument
Set oDoc_Target = Documents.Add

With oDoc_Target
    .Range = ""
    Set oTable = .Tables.Add(Range:=.Range, NumRows:=2, NumColumns:=3)
    With oTable
        .Cell(1, 1).Range.Text = "Acronym"
        .Cell(1, 2).Range.Text = "Definition"
        .Cell(1, 3).Range.Text = "Page"
        .Rows(1).HeadingFormat = True
        .Rows(1).Range.Font.Bold = True
        .PreferredWidthType = wdPreferredWidthPercent
        .Columns(1).PreferredWidth = 20
        .Columns(2).PreferredWidth = 60
        .Columns(3).PreferredWidth = 20
    End With
End With

With oDoc_Source
    Set oRange = .Range
    n = 1
    With oRange.Find
        .Text = "<[A-Z][A-Z\-&/]{1,}>"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = True
        .MatchWildcards = True
        Do While .Execute
            strAcronym = oRange
            If InStr(1, strAllFound, "#" & strAcronym & "#") = 0 Then
                If n > 1 Then oTable.Rows.Add
                strAllFound = strAllFound & strAcronym & "#"
                With oTable
                    .Cell(n + 1, 1).Range.Text = strAcronym
                    .Cell(n + 1, 3).Range.Text = oRange.Information(wdActiveEndPageNumber)
                End With
                n = n + 1
            End If
        Loop
    End With
End With
  
With Selection
    .Sort ExcludeHeader:=True, FieldNumber:="Column 1", SortFieldType:=wdSortFieldAlphanumeric, SortOrder:=wdSortOrderAscending
    .HomeKey (wdStory)
End With

Set oDoc_Source = Nothing
Set oDoc_Target = Nothing
Set oTable = Nothing
MsgBox "Finished extracting " & n - 1 & " acronym(s) to a new document"

End Sub

Sub readStats()
'
' Readability Statistics
'
    StatusBar = "Analyzing this document's readability statistics."
    
    ActiveDocument.Repaginate
    
    Dim fltARI As Double
    Dim txtAlertMessage As String
    
    ActiveDocument.SelectAllEditableRanges
    ActiveDocument.Range.LanguageID = wdEnglishUS
    
    intWord = ActiveDocument.ReadabilityStatistics(1).Value         ' Word count
    intChar = ActiveDocument.ReadabilityStatistics(2).Value         ' Character count
    intPara = ActiveDocument.ReadabilityStatistics(3).Value
    intSent = ActiveDocument.ReadabilityStatistics(4).Value
    intSentPara = ActiveDocument.ReadabilityStatistics(5).Value
    intWordSent = ActiveDocument.ReadabilityStatistics(6).Value
    intCharWord = ActiveDocument.ReadabilityStatistics(7).Value
    ' Debug.Print "Passive: " & ActiveDocument.ReadabilityStatistics(8).Value     ' << THIS IS A KNOWN BUG IN MS-WORD, ALWAYS RETURNING 0.  >>
    intFlesEase = ActiveDocument.ReadabilityStatistics(9).Value
    intFlKiGrad = ActiveDocument.ReadabilityStatistics(10).Value
    intPages = ActiveDocument.ActiveWindow.Panes(1).Pages.Count
    intWordsPage = Round(intWord / intPages, 0)
    fltARI = (4.71 * intCharWord) + (0.5 * intWordSent) - 21.43
    rndARI = Round(fltARI, 1)
    rndAGL = Round(((intFlKiGrad + rndARI) / 2), 1)
    
    Debug.Print ActiveDocument.Name
    Debug.Print "Characters: " & intChar
    Debug.Print "Words: " & intWord
    Debug.Print "Sentences: " & intSent
    Debug.Print "Paragraphs: " & intPara
    Debug.Print "Pages: " & intPages
    Debug.Print "Chars/Word: " & intCharWord
    Debug.Print "Words/Sent: " & intWordSent
    Debug.Print "Sents/Para: " & intSentPara
    Debug.Print "Words/Page: " & intWordsPage
    Debug.Print "Flesch Ease: " & intFlesEase
    Debug.Print "F-K Grade: " & intFlKiGrad
    Debug.Print "ARI: " & rndARI
    Debug.Print "AGL: " & rndAGL
    Debug.Print "CSV: " & _
        intChar & ", " & _
        intWord & ", " & _
        intSent & ", " & _
        intPara & ", " & _
        intPages & ", " & _
        intCharWord & ", " & _
        intWordSent & ", " & _
        intSentPara & ", " & _
        intWordsPage & ", , " & _
        intFlesEase & ", " & _
        intFlKiGrad & ", " & _
        rndARI & ", " & _
        rndAGL
    
    txtAlertMessage = "The following are the readability statistics for this document." & vbCr & " - Flesch Reading Ease: " & intFlesEase & vbCr & " - Flesch-Kincaid Grade: " & intFlKiGrad & vbCr & " - Automated Reading Index: " & rndARI & vbCr & " - Automated Grade Level : " & rndAGL
    
    MsgBox txtAlertMessage, Buttons:=vbInformation, Title:="Readability  Statistics"
    
End Sub


