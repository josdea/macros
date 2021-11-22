Sub Slide_Display_Slide_ID_Number()

Dim sld As Slide
Set sld = ActiveWindow.Selection.SlideRange(1)
MsgBox "Current SlideID: " & sld.slideid & " Current Slide Index: " & sld.SlideIndex & " Slide Number: " & sld.SlideNumber
End Sub
