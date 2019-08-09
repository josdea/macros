Sub changeAllNotes()

Dim aSlide As Slide
Dim noteShp As Shape
Dim slidesModified As Integer

slidesModified = 0 ' keep track of how many slides were changed

For Each aSlide In ActivePresentation.Slides
  For Each noteShp In aSlide.NotesPage.Shapes
    If noteShp.PlaceholderFormat.Type = ppPlaceholderBody Then
       slidesModified = slidesModified + 1 'counts the slide notes that have been modified
          With noteShp.TextFrame.TextRange.Font
          .Size = 12 'font size
          .Name = "Calibri" 'font name
          .Bold = False 'remove all bold
          .Italic = False 'remove all italic
          .Underline = False 'remove all underline
          End With
    End If 'end of placeholder format type
  Next 'next shape in notespage
Next 'next slide
MsgBox ("All Done, Slides Modified: " & slidesModified)
End Sub
