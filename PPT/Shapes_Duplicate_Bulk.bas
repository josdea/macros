Option Explicit

Sub Shapes_Duplicate_Bulk()

'REQUIRES ShapeDetails CLASS MODULE

    Dim oPresentation As Presentation
    Dim sld    As Slide
    Dim shp    As Shape
    Dim sldShpColl As New Collection
    Dim ntShpColl As New Collection
    Dim oShpDetails As ShapeDetails
    Dim boolShpPresent As Boolean
    boolShpPresent = False
    Set oPresentation = ActivePresentation
    
    If MsgBox("Use current slide to find shapes?", (vbYesNo + vbQuestion), "Use Current Slide?") = vbYes Then
        Set sld = oPresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex)
        If MsgBox("Use slide shapes? Otherwise notes shapes will be used", (vbYesNo + vbQuestion), "Slide or Notes?") = vbYes Then
            ' ActiveWindow.ViewType = ppViewNormal
            For Each shp In sld.Shapes
                ' shp.Select
                If shp.Visible = True Then
                    If MsgBox("Memorize Slide Shape: " & shp.name, (vbYesNo + vbQuestion), "Memorize?") = vbYes Then
                        
                        Set oShpDetails = New ShapeDetails
                        Set oShpDetails.oshp = shp
                        sldShpColl.Add oShpDetails, shp.name
                    End If
                End If
            Next shp
        Else                                      ' loop note shapes
            ActiveWindow.ViewType = ppViewNotesPage
            For Each shp In sld.NotesPage.Shapes
                If shp.Visible = True Then
                    shp.Select
                    If MsgBox("Memorize Note Shape: " & shp.name, (vbYesNo + vbQuestion), "Memorize?") = vbYes Then
                        
                        Set oShpDetails = New ShapeDetails
                        Set oShpDetails.oshp = shp
                        ntShpColl.Add oShpDetails, shp.name
                    End If
                End If
            Next shp
        End If                                    ' slide or notes
        
        For Each sld In oPresentation.Slides
            For Each oShpDetails In sldShpColl
                For Each shp In sld.Shapes
                    If shp.name = oShpDetails.name Then 'shape exists so update size
                        Call updateShape(shp, oShpDetails)
                        boolShpPresent = True
                        Exit For
                    End If
                Next shp
                If boolShpPresent = False Then    ' shape isnt there so create it
                    Set shp = sld.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, left:=oShpDetails.left, top:=oShpDetails.top, width:=oShpDetails.width, height:=oShpDetails.height)
                    Call updateShape(shp, oShpDetails)
                    shp.TextFrame.TextRange.text = oShpDetails.text
                End If
                boolShpPresent = False
            Next oShpDetails
            For Each oShpDetails In ntShpColl
                For Each shp In sld.NotesPage.Shapes
                    If shp.name = oShpDetails.name Then
                        Call updateShape(shp, oShpDetails)
                        boolShpPresent = True
                        Exit For
                    End If
                Next shp
                If boolShpPresent = False Then
                    Set shp = sld.NotesPage.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, left:=oShpDetails.left, top:=oShpDetails.top, width:=oShpDetails.width, height:=oShpDetails.height)
                    Call updateShape(shp, oShpDetails)
                    shp.TextFrame.TextRange.text = oShpDetails.text
                End If
                boolShpPresent = False
            Next oShpDetails
        Next sld
    End If
    MsgBox "All Done"
End Sub

Private Function updateShape(shp As Shape, oShpDetails As ShapeDetails) As Boolean
    With shp
        .LockAspectRatio = False
        .name = oShpDetails.name
        .left = oShpDetails.left
        .top = oShpDetails.top
        .height = oShpDetails.height
        .width = oShpDetails.width
        .LockAspectRatio = oShpDetails.boolLockAspectRatio
        .fill.ForeColor.RGB = oShpDetails.fillForeColor
        
        If .HasTextFrame Then
            With .TextFrame
                .TextRange.Font.Size = oShpDetails.fontSize
                .TextRange.Font.Color = oShpDetails.fontColor
                .WordWrap = oShpDetails.boolWordWrap
                .MarginBottom = oShpDetails.margBottom
                .MarginTop = oShpDetails.margTop
                .MarginLeft = oShpDetails.margLeft
                .MarginRight = oShpDetails.margRight
                .VerticalAnchor = oShpDetails.alignVertical
                .autoSize = oShpDetails.autoSize
                If oShpDetails.boolOverwriteText = True Then
                    .TextRange.text = oShpDetails.text
                End If
            End With
            .TextEffect.alignment = oShpDetails.alignHorizontal
        End If
        
    End With
End Function

