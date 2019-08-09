Option Explicit

Sub OutputSlides()
    Dim sld As Slide
    Dim shp As Shape
    Dim i As Long
    Dim oAction As ActionSetting
    Dim oHyperlink As Hyperlink

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            Debug.Print "Shape #" & shp.Id & " (" & shp.Name & ") - Slide: " & sld.SlideNumber & " Position: " & shp.Left & "," & shp.Top _
                 ; " Size: " & shp.Width & "x" & shp.Height


            For Each oAction In shp.ActionSettings
                On Error Resume Next

                If oAction.Action = ppActionHyperlink Then
                    Set oHyperlink = oAction.Hyperlink

                    ''See more: http://www.pptfaq.com/FAQ00162_Hyperlink_-SubAddress_-_How_to_interpret_it.htm
                    Dim parts() As String
                    Dim slideId As Long
                    Dim slideIndex As Long
                    Dim slideTitle As String
                    Dim linkedSlide As Slide

                    parts = Split(oHyperlink.SubAddress, ",")

                    slideId = CLng(parts(0))
                    slideIndex = CLng(parts(1))
                    slideTitle = parts(2)

                    If slideId > 0 Then
                        Debug.Print "  --Internal hyperlink to slide #: " & slideIndex & "(id: " & slideId&; ", title: " & slideTitle & ")"

                        ''this gets you a reference to the linked slide if you need it:
                        ''Set linkedSlide = shp.Parent.Parent.Slides(slideIndex)
                    End If

                End If
            Next oAction
        Next shp
    Next sld
End Sub