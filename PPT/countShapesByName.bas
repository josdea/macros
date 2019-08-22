Option Explicit

Sub countShapesByName()
Dim sld As Slide
Dim shp As Shape
Dim shapeNameToCount As String
Dim shapesCounted As Integer

shapeNameToCount = InputBox("What is the shape name to count on all slides (case sensitive?")
For Each sld In ActivePresentation.Slides        ' iterate slides
        
         For Each shp In sld.Shapes
         
         If shp.name = shapeNameToCount Then
                  shapesCounted = shapesCounted + 1
                                    
            End If
         Next shp
    
        
    Next sld        ' end of iterate slides
   
MsgBox shapesCounted & " shapes match the name: " & shapeNameToCount

End Sub



