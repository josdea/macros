Option Explicit
Sub Text_Remove_Double_Spaces()                   ' checked 1/17/20
    'Removes all text which has double spacebars and replaces with one
    Dim spacesRemoved As Integer
    Dim sld    As Slide
    Dim shp    As Shape
    Dim shpText As String
    spacesRemoved = 0
    Dim shapeCount As Integer
    shapeCount = 0
    If MsgBox("Do you want To replace all instances of multiple spaces With one space", (vbYesNo + vbQuestion), "Remove extra Spaces?") = vbYes Then
        For Each sld In ActivePresentation.Slides
            For Each shp In sld.Shapes
                shapeCount = shapeCount + 1
                If shp.HasTextFrame Then
                    shpText = shp.TextFrame.TextRange.Text 'Get the shape's text
                    Do While InStr(shpText, "  ") > 0
                        shpText = Trim(Replace(shpText, "  ", " "))
                        spacesRemoved = spacesRemoved + 1
                    Loop
                    shp.TextFrame.TextRange.Text = shpText 'Put the new text in the shape
                Else
                    shpText = vbNullString
                End If
            Next shp
            For Each shp In sld.NotesPage.Shapes
                shapeCount = shapeCount + 1
                If shp.HasTextFrame Then
                    shpText = shp.TextFrame.TextRange.Text 'Get the shape's text
                    Do While InStr(shpText, "  ") > 0
                        shpText = Trim(Replace(shpText, "  ", " "))
                        spacesRemoved = spacesRemoved + 1
                    Loop
                    shp.TextFrame.TextRange.Text = shpText 'Put the new text in the shape
                Else
                    shpText = vbNullString
                End If
            Next shp
        Next sld
    End If
    MsgBox spacesRemoved & " places where extra spacing was removed in " & shapeCount & " shapes."
End Sub