Sub changeAllNotes()

Dim aSlide As Slide
Dim noteShp As Shape
Dim slidesModified As Integer

slidesModified = 0

For Each aSlide In ActivePresentation.Slides
  For Each noteShp In aSlide.NotesPage.Shapes
    If noteShp.PlaceholderFormat.Type = ppPlaceholderBody Then
      If noteShp.HasTextFrame And noteShp.TextFrame.TextRange.Text = "" Then 'makes sure that noteshape has a text frame and the text is blank
       slidesModified = slidesModified + 1 'counts the slide notes that have been modified
       ' If noteShp.TextFrame.TextRange.Text <> "" Then
        ' If noteShp.TextFrame.HasText Then ' detect whether the notes page has been edited at some point
        ' & vbCrLf & ' adds a line break
         ' MsgBox ("notes have been edited on:" & noteShp.TextFrame.TextRange.Text & aSlide.SlideNumber)
         ' noteShp.TextFrame.HasText = False
          noteShp.TextFrame.TextRange.Text = "Slide Purpose" & vbCrLf & vbCrLf & "Instructor Notes" 'sets the text to this
          
          'noteShp.TextFrame.TextRange.Text = "" 'sets the text to this
                    'noteShp.TextFrame.TextRange.Text = Replace(noteShp.TextFrame.TextRange.Text, "\1", "<\tag>")
         'End If
      End If 'end of if hastextfram and text is blank
    End If 'end of placeholder format type
  Next 'next shape in notespage
Next 'next slide
MsgBox ("All Done, Slides Modified: " & slidesModified)
End Sub

