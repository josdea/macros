Function addImage()
    Dim currentPresentation As Presentation
    Set currentPresentation = ActivePresentation
    
    Dim sld    As Slide
    Dim shpPlaceholder As Shape
    Dim shpInserted As Shape
    
    Dim strPath As String
    
    Set sld = currentPresentation.Slides(1)
    
    'INSERT IMAGE
    strPath = Environ("USERPROFILE") & "\Downloads\player action.png"
    Set shpInserted = sld.Shapes.AddPicture(fileName:=strPath, LinkToFile:=msoTrue, SaveWithDocument:=msoTrue, Left:=0, Top:=0)
    
    'INSERT VIDEO
    strPath = Environ("USERPROFILE") & "\Downloads\vidinserttest2.mp4"
    Set shpInserted = sld.Shapes.AddMediaObject2(strPath, msoTrue, msoTrue)
    
    'BELOW IS UNEEDED IF THERE IS AN UNUSED PLACEHOLDER AVAILABLE
    Set shpPlaceholder = sld.Shapes(3)
    With shpInserted
        .LockAspectRatio = msoTrue
        If .Width > .Height Then
            .Width = shpPlaceholder.Width
            .Left = shpPlaceholder.Left
            .Top = ((shpPlaceholder.Height - .Height) / 2) + shpPlaceholder.Top
        ElseIf .Width < .Height Then
            .Height = shpPlaceholder.Height
            .Top = shpPlaceholder.Top
            .Left = ((shpPlaceholder.Width - .Width) / 2) + shpPlaceholder.Left
        Else
            .Height = shpPlaceholder.Height
            .Top = shpPlaceholder.Top
            .Left = ((shpPlaceholder.Width - .Width) / 2) + shpPlaceholder.Left
        End If
    End With
    
End Function