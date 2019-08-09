Sub resetSlideLayout(sld As Slide)

Debug.Print "Resetting Slide Layout for Slide " & sld.SlideNumber
    Application.CommandBars.ExecuteMso ("SlideReset")
    
End Sub