Sub createSections(sectionsDesired As Integer)
Dim i As Integer
For i = 1 To sectionsDesired

statusOutput "Creating section " & i & " of " & sectionsDesired
ActivePresentation.SectionProperties.AddBeforeSlide 1, "Module " & i

Next i

End Sub