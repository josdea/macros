Option Explicit
Sub Shapes_Count_By_Name()                        ' checked 1/17/20
    'Count all shapes that have specified name
    Dim sld    As Slide
    Dim shp    As Shape
    Dim shapeNameToCount As String
    Dim shapesCounted As Integer
    shapeNameToCount = InputBox("What Is the shape name To count On all slides (case sensitive And notes Not counted")
    For Each sld In ActivePresentation.Slides     ' iterate slides
        For Each shp In sld.Shapes
            If shp.Name = shapeNameToCount Then
                shapesCounted = shapesCounted + 1
            End If
        Next shp
    Next sld                                      ' end of iterate slides
    MsgBox shapesCounted & " shapes match the name: " & shapeNameToCount
End Sub