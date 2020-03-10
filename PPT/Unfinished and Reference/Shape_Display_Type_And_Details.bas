Option Explicit
Sub Shape_Display_Type_And_Details()
    'Provides additional details about currently selected shape or shapes
    Dim currentPresentation As Presentation
    Set currentPresentation = ActivePresentation
    Dim shp    As Shape
    Dim intShapeCount As Integer
    intShapeCount = 0
    Dim strBoxText As String
    
    Dim colBool As New Collection
    colBool.Add "True", "-1"
    colBool.Add "False", "0"
    
    Dim colType As New Collection
    colType.Add "3D model", "30"
    colType.Add "AutoShape", "1"
    colType.Add "Callout", "2"
    colType.Add "Canvas", "20"
    colType.Add "Chart", "3"
    colType.Add "Comment", "4"
    colType.Add "Content Office Add-in", "27"
    colType.Add "Diagram", "21"
    colType.Add "Embedded OLE object", "7"
    colType.Add "Form control", "8"
    colType.Add "Freeform", "5"
    colType.Add "Graphic", "28"
    colType.Add "Group", "6"
    colType.Add "SmartArt graphic", "24"
    colType.Add "Ink", "22"
    colType.Add "Ink comment", "23"
    colType.Add "Line", "9"
    colType.Add "Linked 3D model", "31"
    colType.Add "Linked graphic", "29"
    colType.Add "Linked OLE object", "10"
    colType.Add "Linked picture", "11"
    colType.Add "Media", "16"
    colType.Add "OLE control object", "12"
    colType.Add "Picture", "13"
    colType.Add "Placeholder", "14"
    colType.Add "Script anchor", "18"
    colType.Add "Mixed shape type", "-2"
    colType.Add "Table", "19"
    colType.Add "Text box", "17"
    colType.Add "Text effect", "15"
    colType.Add "Web video", "26"
    
    Dim colPlaceholderType As New Collection
    colPlaceholderType.Add "Bitmap", "9"
    colPlaceholderType.Add "Body", "2"
    colPlaceholderType.Add "Center Title", "3"
    colPlaceholderType.Add "Chart", "8"
    colPlaceholderType.Add "Date", "16"
    colPlaceholderType.Add "Footer", "15"
    colPlaceholderType.Add "Header", "14"
    colPlaceholderType.Add "Media Clip", "10"
    colPlaceholderType.Add "Mixed", "-2"
    colPlaceholderType.Add "Object", "7"
    colPlaceholderType.Add "Organization Chart", "11"
    colPlaceholderType.Add "Picture", "18"
    colPlaceholderType.Add "Slide Number", "13"
    colPlaceholderType.Add "Subtitle", "4"
    colPlaceholderType.Add "Table", "12"
    colPlaceholderType.Add "Title", "1"
    colPlaceholderType.Add "Vertical Body", "6"
    colPlaceholderType.Add "Vertical Object", "17"
    colPlaceholderType.Add "Vertical Title", "5"
    
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        intShapeCount = intShapeCount + 1
        With shp
            
            strBoxText = "Shape is Placeholder: " & colBool(CStr(msoPlaceholder)) & vbCrLf
            
            'TODO here convert above to string before calling mso placeholder was title
            If shp.Type = msoPlaceholder Then
                strBoxText = strBoxText & "PlaceholderFormat.Type: " & colPlaceholderType(.PlaceholderFormat.Type) & vbCrLf _
                           & "PlaceholderFormat.ContainedType: " & colType(.PlaceholderFormat.ContainedType) & vbCrLf
            Else
                strBoxText = strBoxText & "PlaceholderFormat.Type: NA" & vbCrLf _
                           & "PlaceholderFormat.ContainedType: NA" & vbCrLf
            End If
            
            
            
            
            
            
            If shp.HasTextFrame Then
                strBoxText = strBoxText & "Autosize: " & .TextFrame.AutoSize & "(0:no autofit, -2: shrink text, 1: resize shape)" & vbCrLf
            End If
        End With
        
        
        'MsgBox "Shape " & intShapeCount & " of " & ActiveWindow.Selection.ShapeRange.count & vbCrLf
        
        MsgBox strBoxText, vbOKOnly, shp.Name & " (" & intShapeCount & " of " & ActiveWindow.Selection.ShapeRange.count & ")"
        
        
        strBoxText = ""
        
    Next shp
End Sub