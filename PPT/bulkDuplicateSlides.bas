
' note doesn't duplciate the handout notes portion as straight copying does

Sub duplicateSlide(slidesDesired As Integer, slideToDuplicate As Integer)
Dim i As Integer

    For i = 1 To slidesDesired
        
        Debug.Print "Creating slide " & i & " of " & slidesDesired
  ActivePresentation.Slides(slideToDuplicate).Duplicate
        
    Next i

End Sub