Option Explicit
Sub Video_Convert_Linked_To_Embedded()
    'Converts all linked videos to embedded
    MsgBox "This can take a few minutes, do not worry. Select ok to continue"
    Dim shp    As Shape
    Dim sld    As Slide
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            On Error Resume Next
            shp.LinkFormat.BreakLink
            On Error GoTo 0
        Next shp
    Next sld
    MsgBox "All Done"
End Sub