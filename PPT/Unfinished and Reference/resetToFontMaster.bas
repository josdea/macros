' UNIFINSIHED TODO

Option Explicit

Sub resetToFontMaster()
    
    Dim sld                                       As Slide
    Dim shp                                       As Shape
    
    If MsgBox("Are you sure you want to reset all titles, text, and notes to the master font theme", (vbYesNo + vbQuestion), "Reset Font?") = vbYes Then
        
        For Each sld In ActivePresentation.Slides        ' iterate slides
            ActiveWindow.ViewType = ppViewNormal
            sld.Select
            
            For Each shp In sld.Shapes
                
                
                shp.Select
                
                If shp.HasTextFrame Then
                If shp.Type = msoPlaceholder Then
                    If shp.PlaceholderFormat.Type = ppPlaceholderTitle Or shp.PlaceholderFormat.Type = ppPlaceholderCenterTitle Then
                        shp.TextFrame.TextRange.Font.name = "+mj-lt"
                        
                    Else
                        shp.TextFrame.TextRange.Font.name = "+mn-lt"
                    End If
                    Else
                    shp.TextFrame.TextRange.Font.name = "+mn-lt"
                    End If
                End If
                
            Next shp        ' end of iterate shapes
            
            ActiveWindow.ViewType = ppViewNotesPage
            For Each shp In sld.NotesPage.Shapes
                       
             shp.Select
             
                      If shp.HasTextFrame Then
                If shp.Type = msoPlaceholder Then
                    If shp.PlaceholderFormat.Type = ppPlaceholderTitle Or shp.PlaceholderFormat.Type = ppPlaceholderCenterTitle Then
                        shp.TextFrame.TextRange.Font.name = "+mj-lt"
                        
                    Else
                        shp.TextFrame.TextRange.Font.name = "+mn-lt"
                    End If
                    Else
                    shp.TextFrame.TextRange.Font.name = "+mn-lt"
                    End If
                End If
                
                
            
            
            Next shp
            
            
        Next sld        ' end of iterate slides
        
        MsgBox "All Done"
        
    Else
        MsgBox ("Action canceled.")
        
    End If
    
End Sub
