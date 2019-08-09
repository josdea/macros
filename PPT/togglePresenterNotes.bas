Option Explicit

Sub togglePresenterNotes()
    
    Dim toggleOn                                  As Boolean
    toggleOn = msoTrue
    
    If MsgBox("Do you want to hide all presenter note shapes on the notes pages. If you answer 'no' then they will all be made visible.", (vbYesNo + vbQuestion), "Toggle Presenter Notes?") = vbYes Then
    toggleOn = msoFalse
    End If
    
    Dim sld                                       As Slide        ' declare slide object
    Dim shp                                       As Shape        ' declare shape object
    For Each sld In ActivePresentation.Slides        ' iterate slides
        Debug.Print "Slide " & sld.SlideNumber & " of " & ActivePresentation.Slides.Count
        
        For Each shp In sld.NotesPage.Shapes ' iterate note shapes
            If shp.Type = msoPlaceholder Then        ' check if its a placeholder
            If shp.PlaceholderFormat.Type = ppPlaceholderBody Then ' its presenter notes
                shp.visible = toggleOn
            End If
        End If
    Next shp ' end of iterate shapes
Next sld        ' end of iterate slides

    Debug.Print "All done toggling presenter notes";
   
    If Not CommandBars.GetPressedMso("SelectionPane") Then CommandBars.ExecuteMso ("SelectionPane")
    
    MsgBox "All done toggling presenter notes"
    ActiveWindow.ViewType = ppViewNotesPage

End Sub
