Option Explicit
Sub PresenterNotes_Add_Comment_If_Empty()         ' checked 1/17/20
    'Adds a comment on every slide with empty presenter notes
    Dim sld    As Slide                           ' declare slide object
    Dim shp                                       As Shape ' declare shape object
    Dim intCommentCount As Integer
    intCommentCount = 0
    For Each sld In ActivePresentation.Slides     ' iterate slides
        For Each shp In sld.NotesPage.Shapes      ' iterate note shapes
            If shp.Type = msoPlaceholder Then     ' check if its a placeholder
                If shp.PlaceholderFormat.Type = ppPlaceholderBody Then ' its presenter notes
                    If shp.TextFrame.HasText = False Then
                        sld.Comments.Add 12, 12, "Auto", "JMD", "TODO: Need Presenter Notes"
                        intCommentCount = intCommentCount + 1
                    ElseIf shp.TextFrame.TextRange.Text = "" Then
                        sld.Comments.Add 12, 12, "Auto", "JMD", "TODO: Need Presenter Notes"
                        intCommentCount = intCommentCount + 1
                    End If
                End If
            End If
        Next shp                                  ' end of iterate shapes
    Next sld                                      ' end of iterate slides
    MsgBox "All Done " & intCommentCount & " comments were added"
End Sub