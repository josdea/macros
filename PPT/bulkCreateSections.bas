'Prompts the user for a number, optional prefix, or optional suffix and then
'generates that many sections. Great for easily creating: Module 1, Module 2..
'etc.

Option Explicit
Sub bulkCreateSection()

Dim sectionsDesired As Integer
Dim sectionPrefix As String
Dim sectioSuffix As String
Dim inputText As String
Dim inputAnswer As Variant

 inputText = "Please enter number of desired sections to create?"
 
   Do
    inputAnswer = InputBox(inputText)
    
    'Check if user selected cancel button
      If TypeName(inputAnswer) = "Boolean" Then Exit Sub
      
  Loop While inputAnswer <= 0

If inputAnswer > 0 Then
sectionsDesired = inputAnswer
createSections sectionsDesired, InputBox("Enter optional prefix." & vbNewLine & "(remember to add a space if you need one)", , "Module "), InputBox("Enter an optional suffix." & vbNewLine & "(remember to add a space if you need one)")
End If

End Sub

Sub createSections(sectionsDesired As Integer, sectionPrefix As String, sectionSuffix As String)
Dim i As Integer

For i = 1 To sectionsDesired

' statusOutput "Creating section " & i & " of " & sectionsDesired
ActivePresentation.SectionProperties.AddBeforeSlide 1, sectionPrefix & i & sectionSuffix

Next i

End Sub
