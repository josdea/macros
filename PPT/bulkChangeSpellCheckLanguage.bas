

Option Explicit
Sub ChangeSpellCheckLanguage()
Dim currentSlide As Integer 'current slide number
Dim currentShape As Integer ' current shape on current slide or notes
Dim slideCount As Integer
Dim shapeCount As Integer    'Find out how many slides there are in the presentation
Dim noteShape As Shape
Dim currentLanguage As Integer
Dim totalShapeCount As Integer

totalShapeCount = 0

' ask for language see here for full list https://docs.microsoft.com/en-us/office/vba/api/powerpoint.textrange.languageid
If MsgBox("Do you want UK English Spelling for Slides and notes? Otherwise US English will be applied", vbYesNo) = vbYes Then
currentLanguage = msoLanguageIDEnglishUK 'language set to UK
Else
currentLanguage = msoLanguageIDEnglishUS 'language set to US
End If

slideCount = ActivePresentation.Slides.Count     'Get slide count
MsgBox ("There are " & slideCount & " slides")

For currentSlide = 1 To slideCount        'Find out how many shapes there are so identify all the text boxes
shapeCount = ActivePresentation.Slides(currentSlide).Shapes.Count         'Loop through all the shapes on that slide changing the language option

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

