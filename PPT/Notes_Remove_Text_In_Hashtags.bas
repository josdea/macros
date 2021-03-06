Option Explicit
Sub Notes_Remove_Text_In_Hashtags()      ' checked 1/17/20
    'Removes all text wrapped in ##double hashtags## in presenter notes
    Dim stringBwDels As String, originalString As String, firstDelPos As Integer, secondDelPos As Integer, stringToReplace As String, replacedString As String
    Dim sld                                       As Slide ' declare slide object
    Dim shp                                       As Shape ' declare shape object
    For Each sld In ActivePresentation.Slides     ' iterate slides
    For Each shp In sld.NotesPage.Shapes          ' iterate note shapes
    If shp.Type = msoPlaceholder Then             ' check if its a placeholder
    If shp.PlaceholderFormat.Type = ppPlaceholderBody Then ' its presenter notes
    originalString = shp.TextFrame2.TextRange.Text
    stringToReplace = ""
    firstDelPos = InStr(originalString, "##") - 1 ' position of start delimiter
    secondDelPos = InStr(firstDelPos + 2, originalString, "##") ' position of end delimiter
    If secondDelPos <> 0 Then
        stringBwDels = Mid(originalString, firstDelPos + 1, secondDelPos - firstDelPos + 2) 'extract the string between two delimiters
    Else
        stringBwDels = 0
    End If
    replacedString = Replace(originalString, stringBwDels, stringToReplace)
    shp.TextFrame2.TextRange.Text = replacedString
End If
End If
Next shp                                          ' end of iterate shapes
Next sld                                          ' end of iterate slides
MsgBox "All Done"
End Sub