Function findObjectives()
    Dim sld    As Slide
    Dim shp    As Shape
    Dim shp2   As Shape
    Debug.Print "break"
    If shp.HasTextFrame = msoTrue Then
        
        If InStr(UCase(shp.TextFrame.TextRange.Text), "OBJECTIVE") Then
            sld.Select
            Debug.Print "found it"
            
            For Each shp2 In sld.Shapes
                If shp2.PlaceholderFormat.Type = ppPlaceholderObject Then
                    'MsgBox shp2.TextFrame.TextRange.Text
                    
                    Dim para As Variant
                    For Each para In shp2.TextFrame.TextRange.Paragraphs
                        
                        MsgBox para.Text
                        
                    Next para
                    
                End If
                
            Next shp2
            
        End If
    End If
    
End Function