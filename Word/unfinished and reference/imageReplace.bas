Option Explicit

Sub main()
Call convertToInline
Call updateSlideImages

End Sub

Sub updateSlideImages()
Dim shp As InlineShape
Dim imageCount As Integer
Dim imgWidth As Single
Dim imgHeight As Single
Dim strImagePath As String
Dim rng As Range
imageCount = 0
Application.ScreenUpdating = False
For Each shp In ActiveDocument.InlineShapes

If (InStr(shp.AlternativeText, "Slide")) Then
imageCount = imageCount + 1
strImagePath = ActiveDocument.Path & "\slideImages\Slide" & imageCount & ".PNG"
imgWidth = shp.Width
Set rng = shp.Range
shp.Delete
Debug.Print "Updating Slide " & imageCount
Set shp = Selection.InlineShapes.AddPicture(FileName:=strImagePath, LinkToFile:=True, SaveWithDocument:=True, Range:=rng)
With shp
.Select
.LockAspectRatio = msoTrue
.Width = imgWidth
'.AlternativeText = "Slide " & imageCount & " Updated"
End With

End If
Next shp
Application.ScreenUpdating = True
End Sub


Sub convertToInline()

    Dim Shape As Shape
    Dim Index As Integer: Index = 1
    ' Store the count as it will change each time.
    Dim NumberOfShapes As Integer
    NumberOfShapes = ActiveDocument.Shapes.Count

    ' Break out if either all shapes have been checked or there are none left.
    Do While ActiveDocument.Shapes.Count > 0 And Index < NumberOfShapes + 1

        With ActiveDocument.Shapes(Index)
            If .Type = msoPicture Then
                ' If the shape is a picture convert it to inline.
                ' It will be removed from the collection so don't increment the Index.
                
                If (InStr(.AlternativeText, "Slide")) Then
                .ConvertToInlineShape
                End If
            Else
                ' The shape is not a picture so increment the index (move to next shape).
                Index = Index + 1
            End If
        End With

    Loop

End Sub


