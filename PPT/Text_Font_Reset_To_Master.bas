Option Explicit
Sub Text_Font_Reset_To_Master()                   ' checked 1/17/20
    'Resets all text in all shapes and notes to master theme font
    Dim sld                                       As Slide
    Dim shp                                       As Shape
    Dim shapesAffected As Integer
    shapesAffected = shapesAffected + 1
    If MsgBox("Are you sure you want To reset all titles, text, And notes To the master font theme", (vbYesNo + vbQuestion), "Reset Font?") = vbYes Then
        For Each sld In ActivePresentation.Slides ' iterate slides
            For Each shp In sld.Shapes
                If shp.HasTextFrame Then
                    If shp.Type = msoPlaceholder Then
                        If shp.PlaceholderFormat.Type = ppPlaceholderTitle Or shp.PlaceholderFormat.Type = ppPlaceholderCenterTitle Then
                            shp.TextFrame.TextRange.Font.Name = "+mj-lt"
                            shapesAffected = shapesAffected + 1
                        Else
                            shp.TextFrame.TextRange.Font.Name = "+mn-lt"
                            shapesAffected = shapesAffected + 1
                        End If
                    Else
                        shp.TextFrame.TextRange.Font.Name = "+mn-lt"
                        shapesAffected = shapesAffected + 1
                    End If
                End If
            Next shp                              ' end of iterate shapes
            For Each shp In sld.NotesPage.Shapes
                If shp.HasTextFrame Then
                    If shp.Type = msoPlaceholder Then
                        If shp.PlaceholderFormat.Type = ppPlaceholderTitle Or shp.PlaceholderFormat.Type = ppPlaceholderCenterTitle Then
                            shp.TextFrame.TextRange.Font.Name = "+mj-lt"
                            shapesAffected = shapesAffected + 1
                        Else
                            shp.TextFrame.TextRange.Font.Name = "+mn-lt"
                            shapesAffected = shapesAffected + 1
                        End If
                    Else
                        shp.TextFrame.TextRange.Font.Name = "+mn-lt"
                        shapesAffected = shapesAffected + 1
                    End If
                End If
            Next shp
        Next sld                                  ' end of iterate slides
        MsgBox "All Done. " & shapesAffected & " shapes have been searched And reset"
    Else
        MsgBox ("Action canceled.")
    End If
End Sub