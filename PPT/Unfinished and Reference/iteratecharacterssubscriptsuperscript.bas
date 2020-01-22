Function iterateShapeCharacters(sld As Slide, shp As Shape)
    
    ' iterate every character of all text. below finds super and subscript
    
    If shp.TextFrame.HasText Then
        
        With shp.TextFrame.TextRange
            Dim i                                 As Integer
            For i = 1 To .Characters.count
                
                If (.Characters(i).Font.Subscript) Or (.Characters(i).Font.Superscript) Then
                    sld.Select
                    'shp.Select
                    End
                End If
            Next
        End With
        
    End If
    
End Function