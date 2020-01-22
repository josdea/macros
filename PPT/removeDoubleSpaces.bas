Option Explicit
Dim spacesRemoved As Integer

Sub removeDoubleSpaces()
Dim sld As Slide
Dim shp As Shape


spacesRemoved = 0

 If MsgBox("Do you want to replace all instances of multiple spaces with one space", (vbYesNo + vbQuestion), "Remove extra Spaces?") = vbYes Then

For Each sld In ActivePresentation.Slides
    For Each shp In sld.Shapes
        removespace shp
    Next shp
    
    For Each shp In sld.NotesPage.Shapes
        removespace shp
    Next shp
    
Next sld

End If

MsgBox spacesRemoved & " places where extra spacing was removed"
End Sub

Sub removespace(shp As Shape)
Dim shpText As String
If shp.HasTextFrame Then
            shpText = shp.TextFrame.TextRange.text 'Get the shape's text
            Do While InStr(shpText, "  ") > 0
                shpText = Trim(Replace(shpText, "  ", " "))
                spacesRemoved = spacesRemoved + 1
            Loop
            shp.TextFrame.TextRange.text = shpText 'Put the new text in the shape
        Else
            shpText = vbNullString
        End If

End Sub

