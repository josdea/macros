Option Explicit

Dim maxSlideWords As Integer

Dim maxSlideWordsSlide As Integer


Dim minSlideWords As Integer

Dim minSlideWordsSlide As Integer


Dim totalSlideWords As Integer

Dim avgSlideWords As Double

Dim includeHiddenSlides As Integer


Sub countWords()
maxSlideWords = 0
maxSlideWordsSlide = 1
minSlideWords = 0
minSlideWordsSlide = 1
totalSlideWords = 0
avgSlideWords = 0

Dim message As String
message = "Do you want to count all words on all slides"

 If MsgBox(message, (vbYesNo + vbQuestion), "Count?") = vbYes Then

includeHiddenSlides = MsgBox("Do you want to include hidden slides?", (vbYesNo + vbQuestion), "Hidden Slides Too?")

iterateSlides
 
MsgBox "Completed Counting" & vbNewLine & "Total Slides: " & ActivePresentation.Slides.Count & vbNewLine & "Total Words: " & totalSlideWords & vbNewLine & "Average Words per Slide: " & avgSlideWords & vbNewLine & "Maximum words on a slide: " & maxSlideWords & " on slide " & maxSlideWordsSlide & vbNewLine & "Minimum Words on a slide: " & minSlideWords & " on slide " & minSlideWordsSlide
 
 
 If MsgBox("Do you want to go to the slide that has the most words?", (vbYesNo + vbQuestion), "Go to max") = vbYes Then
  ActivePresentation.Slides(maxSlideWordsSlide).Select
 
 End If
 
 End If


End Sub


Function iterateSlides()
Dim sld As Slide
 For Each sld In ActivePresentation.Slides        ' iterate slides
            'ActiveWindow.ViewType = ppViewNormal
            'sld.Select
            
            If sld.SlideShowTransition.Hidden = msoFalse Or includeHiddenSlides = vbYes Then
            
iterateSlideShapes sld

End If


Next sld

avgSlideWords = totalSlideWords / ActivePresentation.Slides.Count

End Function

Function iterateSlideShapes(sld As Slide)
Dim shp As Shape
Dim wordCount As Integer
Dim arr() As String
Dim shapeText As String

 For Each shp In sld.Shapes
  If shp.HasTextFrame And shp.name <> "LearnerNotes" Then
  
  shapeText = shp.TextFrame.TextRange.Text
  shapeText = Replace(shapeText, vbNewLine, " ")
  shapeText = Replace(shapeText, vbCr, " ")
  shapeText = Replace(shapeText, vbLf, " ")
  shapeText = Replace(shapeText, vbCrLf, " ")
  
   arr = VBA.Split(shapeText, " ")

    Debug.Print shp.name & ": " & UBound(arr) - LBound(arr) + 1
    
    wordCount = wordCount + UBound(arr) - LBound(arr) + 1
    
  End If
 
 
 Next shp

'MsgBox wordCount & " words on slide"

totalSlideWords = totalSlideWords + wordCount

If wordCount > maxSlideWords Then
maxSlideWords = wordCount
maxSlideWordsSlide = sld.SlideNumber
End If

If wordCount < minSlideWords Or sld.SlideNumber = 1 Then
minSlideWords = wordCount
minSlideWordsSlide = sld.SlideNumber

End If

End Function

