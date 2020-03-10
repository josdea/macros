Option Explicit
Sub Image_Add_Comment_For_Missing_AltText()       ' checked 1/17/20
    'Add a comment for all slides with images having missing alternate text
    Dim sld    As Slide                           ' declare slide object
    Dim shp                                       As Shape ' declare shape object
    Dim intCommentCount As Integer
    intCommentCount = 0
    Dim boolImageFound As Boolean
    boolImageFound = False
    For Each sld In ActivePresentation.Slides     ' iterate slides
        For Each shp In sld.Shapes                ' iterate note shapes
            If shp.Type = msoPicture Then         ' check if its a placeholder
                boolImageFound = True
            End If
            
            If shp.Type = msoPlaceholder Then
                If shp.PlaceholderFormat.ContainedType = msoPicture Then
                    boolImageFound = True
                End If
            End If
            
            If boolImageFound = True Then
                boolImageFound = False
                
                If shp.AlternativeText = "" Or InStr(shp.AlternativeText, "generated") Then
                    sld.Comments.Add 12, 12, "Auto", "JMD", "TODO: Need Image ID Or source"
                    intCommentCount = intCommentCount + 1
                End If
            End If
            
        Next shp                                  ' end of iterate shapes
    Next sld                                      ' end of iterate slides
    MsgBox "All Done " & intCommentCount & " comments were added"
End Sub