Option Explicit
Sub Sections_Bulk_Create()                        ' checked 1/17/20
    'Bulk creates sections with optional prefix and or suffix
    Dim sectionsDesired As Integer
    Dim sectionPrefix As String
    Dim sectioSuffix As String
    Dim inputText As String
    Dim inputAnswer As Variant
    Dim strPrefix As String
    Dim strSuffix As String
    inputText = "Please enter number of desired sections To create?"
    Do
        inputAnswer = InputBox(inputText)
        If TypeName(inputAnswer) = "Boolean" Then Exit Sub
    Loop While inputAnswer <= 0
    If inputAnswer > 0 Then
        sectionsDesired = inputAnswer
        strPrefix = InputBox("Enter Optional prefix." & vbNewLine & "(remember To add a space If you need one)", , "Module ")
        strSuffix = InputBox("Enter an Optional suffix." & vbNewLine & "(remember To add a space If you need one)")
        Dim i  As Integer
        For i = 1 To sectionsDesired
            ActivePresentation.SectionProperties.AddBeforeSlide 1, strPrefix & i & strSuffix
        Next i
    End If
End Sub