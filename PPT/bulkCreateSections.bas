Sub Sections_Bulk_Create()

Dim sectionsDesired As Integer
Dim sectionPrefix As String
Dim sectioSuffix As String
Dim inputText As String
Dim inputAnswer As Variant
Dim strPrefix As String
Dim strSuffix As String

 inputText = "Please enter number of desired sections to create?"
 
   Do
    inputAnswer = InputBox(inputText)
    
    'Check if user selected cancel button
      If TypeName(inputAnswer) = "Boolean" Then Exit Sub
      
  Loop While inputAnswer <= 0

If inputAnswer > 0 Then
sectionsDesired = inputAnswer
strPrefix = InputBox("Enter optional prefix." & vbNewLine & "(remember to add a space if you need one)", , "Module ")
strSuffix = InputBox("Enter an optional suffix." & vbNewLine & "(remember to add a space if you need one)")

Dim i As Integer

For i = 1 To sectionsDesired
ActivePresentation.SectionProperties.AddBeforeSlide 1, strPrefix & i & strSuffix
Next i
End If

End Sub
