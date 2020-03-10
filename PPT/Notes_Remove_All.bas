Option Explicit
Sub Presenter_Notes_Remove_All()                  ' checked 1/17/20
    'Deletes all presenter notes on all slides
    Dim oSl    As Slide
    Dim oSh    As Shape
    If MsgBox("Are you sure you want To delete all presenter/instructor notes?", (vbYesNo + vbQuestion), "Delete all Notes?") = vbYes Then
        For Each oSl In ActivePresentation.Slides
            For Each oSh In oSl.NotesPage.Shapes
                If oSh.Type = msoPlaceholder Then
                    If oSh.PlaceholderFormat.Type = ppPlaceholderBody Then
                        If oSh.HasTextFrame Then
                            oSh.TextFrame.TextRange.Text = ""
                        End If
                    End If
                End If
            Next oSh
        Next oSl
    Else
        MsgBox ("Action canceled.")
    End If
End Sub