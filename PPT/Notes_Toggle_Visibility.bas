Option Explicit
Sub PresenterNotes_Toggle_Visibility()            ' checked 1/17/20
    'Toggles visibility of all presenter notes on notes pages
    Dim toggleOn                                  As Boolean
    toggleOn = msoTrue
    If MsgBox("Do you want To hide all presenter note shapes On the notes pages. If you answer `no` then they will all be made visible.", (vbYesNo + vbQuestion), "Toggle Presenter Notes?") = vbYes Then
        toggleOn = msoFalse
    End If
    Dim sld                                       As Slide ' declare slide object
    Dim shp                                       As Shape ' declare shape object
    For Each sld In ActivePresentation.Slides     ' iterate slides
        For Each shp In sld.NotesPage.Shapes      ' iterate note shapes
            If shp.Type = msoPlaceholder Then     ' check if its a placeholder
                If shp.PlaceholderFormat.Type = ppPlaceholderBody Then ' its presenter notes
                    shp.Visible = toggleOn
                End If
            End If
        Next shp                                  ' end of iterate shapes
    Next sld                                      ' end of iterate slides
    If Not CommandBars.GetPressedMso("SelectionPane") Then CommandBars.ExecuteMso ("SelectionPane")
    MsgBox "All done toggling presenter notes"
    ActiveWindow.ViewType = ppViewNotesPage
End Sub