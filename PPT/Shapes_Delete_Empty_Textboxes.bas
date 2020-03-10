Option Explicit
Sub Shapes_Delete_Empty_TextBoxes()

    Dim sld    As Slide
    Dim shp    As Shape
    Dim ShapeIndex As Integer

    For Each sld In ActivePresentation.Slides
        
        For ShapeIndex = sld.Shapes.count To 1 Step -1
            
            If sld.Shapes(ShapeIndex).Type = msoTextBox And Not sld.Shapes(ShapeIndex).TextFrame.HasText Then
                sld.Shapes(ShapeIndex).Delete
            End If
        Next
    Next sld
End Sub