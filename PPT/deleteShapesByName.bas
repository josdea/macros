Option Explicit

Sub deleteShapesByName()
Dim sld As Slide
Dim shp As Shape
Dim shapeNameToDelete As String
Dim shapesDeleted As Integer

shapeNameToDelete = InputBox("What is the shape name to delete on all slides?")
For Each sld In ActivePresentation.Slides        ' iterate slides
        
         For Each shp In sld.Shapes
         
         If shp.name = shapeNameToDelete Then
                  shapesDeleted = shapesDeleted + 1
                  shp.Delete
                     
            End If
         Next shp
    
        
    Next sld        ' end of iterate slides
   
MsgBox shapesDeleted & " shapes have been deleted"

End Sub


