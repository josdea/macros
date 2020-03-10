Option Explicit
Sub Shapes_Delete_By_Name()                       ' checked 1/17/20
    'Deletes all shapes that have specified name
    Dim sld    As Slide
    Dim shp    As Shape
    Dim shapeNameToDelete As String
    Dim shapesDeleted As Integer
    shapeNameToDelete = InputBox("What Is the shape name To delete On all slides?")
    For Each sld In ActivePresentation.Slides     ' iterate slides
        For Each shp In sld.Shapes
            If shp.Name = shapeNameToDelete Then
                shapesDeleted = shapesDeleted + 1
                shp.Delete
            End If
        Next shp
    Next sld                                      ' end of iterate slides
    MsgBox shapesDeleted & " shapes have been deleted"
End Sub