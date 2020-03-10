Option Explicit
Sub Text_Remove_Empty_Lines()
    Dim currentPresentation As Presentation: Set currentPresentation = ActivePresentation
    Dim sld    As Slide
    Dim shp    As Shape
    Dim para   As TextRange
    Dim ln     As TextRange

    For Each sld In currentPresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                For Each para In shp.TextFrame.TextRange.Paragraphs
                    Debug.Print para
                    
                    For Each ln In para.Lines
                        Debug.Print ln
                    Next ln
                Next para
                '   With shp.TextFrame.TextRange
                'MsgBox .Text
                ' .Text = removeMultiBlank(.Text)
                '   End With
            End If
        Next shp
        If sld.HasNotesPage Then
            For Each shp In sld.NotesPage.Shapes
                If shp.HasTextFrame Then
                    With shp.TextFrame.TextRange
                        ' MsgBox .Text
                        .Text = removeMultiBlank(.Text)
                    End With
                End If
            Next shp
        End If
    Next sld

End Sub

Function removeMultiBlank(s As String) As String
    With CreateObject("VBScript.RegExp")
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = "^\s"
        
        removeMultiBlank = .Replace(s, "")
    End With
End Function