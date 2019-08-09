Sub removeAllPresenterNotes()
  Dim currentSlide As Slide
 
 If MsgBox("Are you sure you want to delete all presenter/instructor notes?", (vbYesNo + vbQuestion), "Delete all Notes?") = vbYes then
 
  For Each currentSlide In ActivePresentation.Slides

if currentSlide.NotesPage.Shapes.Placeholders(2)
    currentSlide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange _
    = ""
  Next currentSlide

else
msgBox("Action canceled.")

  end If
End Sub
