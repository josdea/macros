Option Explicit
Sub Text_Language_Toggle_US_UK()                  ' checked 2/17/20
    'Toggles all text in all shapes on all slides and master for spell checking
    Dim currentPresentation As Presentation: Set currentPresentation = ActivePresentation
    Dim currentLanguage As Integer
    Dim totalShapeCount As Integer: totalShapeCount = 0

    Dim strLangSelect As String
    strLangSelect = ""
    
    Dim langList As New Collection
    Dim collMsoLang As New Collection
    langList.Add "English US", "1"
    collMsoLang.Add "1033", "1"
    langList.Add "English UK", "2"
    collMsoLang.Add "2057", "2"
    langList.Add "Arabic", "3"
    collMsoLang.Add "1025", "3"
    langList.Add "Spanish (General)", "4"
    collMsoLang.Add "1034", "4"
    langList.Add "French", "5"
    collMsoLang.Add "1036", "5"
    langList.Add "Russian", "6"
    collMsoLang.Add "1049", "6"
    langList.Add "Polish", "7"
    collMsoLang.Add "1045", "7"
    langList.Add "Romanian", "8"
    collMsoLang.Add "1048", "8"
    
    Dim lang   As Variant
    Dim i      As Integer
    i = 0
    For Each lang In langList
        i = i + 1
        strLangSelect = strLangSelect & ""
        strLangSelect = strLangSelect & i & ". " & lang & vbCrLf
    Next lang
    currentLanguage = collMsoLang(getNumberInput(strLangSelect, "Select from Language", 1, 1, i))
    
    Dim boolUpdateNotesLang As Boolean
    If MsgBox("Do you want to update the notes language as well?", vbYesNo) = vbYes Then
        boolUpdateNotesLang = True
    Else
        boolUpdateNotesLang = False
    End If
    
    Dim sld    As Slide
    Dim shp    As Shape
    For Each sld In currentPresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                shp.TextFrame.TextRange.LanguageID = currentLanguage
                totalShapeCount = totalShapeCount + 1
            End If
        Next shp
        If sld.HasNotesPage And boolUpdateNotesLang = True Then
            For Each shp In sld.NotesPage.Shapes
                If shp.HasTextFrame Then
                    shp.TextFrame.TextRange.LanguageID = currentLanguage
                    totalShapeCount = totalShapeCount + 1
                End If
            Next shp
        End If
    Next sld
    For Each shp In currentPresentation.SlideMaster.Shapes
        If shp.HasTextFrame Then
            shp.TextFrame.TextRange.LanguageID = currentLanguage
            totalShapeCount = totalShapeCount + 1
        End If
    Next shp
    Dim layCustom As CustomLayout
    
    For Each layCustom In currentPresentation.SlideMaster.CustomLayouts
        For Each shp In layCustom.Shapes
            If shp.HasTextFrame Then
                shp.TextFrame.TextRange.LanguageID = currentLanguage
                totalShapeCount = totalShapeCount + 1
            End If
        Next shp
    Next layCustom
    
    MsgBox ("All Done. Total Shapes Set: " & totalShapeCount & ". Press F7 To rerun spellcheck.")
End Sub

Function getNumberInput(strMessage As String, strBoxTitle As String, strDefaultValue As String, intMinNumber As Integer, intMaxNumber As Integer) As Double
    'This function is needed for the above
    Dim strInputField As String
    strMessage = strMessage & vbCrLf & vbCrLf & "Please enter a number between " & intMinNumber & " and " & intMaxNumber & ":"
    
    Do
        'Retrieve an answer from the user
        strInputField = InputBox(strMessage, strBoxTitle, strDefaultValue)
        If StrComp(strInputField, "x", 1) = 0 Then
            End
        ElseIf TypeName(strInputField) = "Boolean" Then 'Check if user selected cancel button
            getNumberInput = -1
        ElseIf Not IsNumeric(strInputField) Then  'input wasnt numeric
            getNumberInput = -1
        Else
            getNumberInput = strInputField        ' Number is numeric
        End If
    Loop While getNumberInput < intMinNumber Or getNumberInput > intMaxNumber Or getNumberInput < 0 'Keep prompting while out of range
    
End Function