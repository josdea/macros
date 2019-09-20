Option Explicit

Function applyTransition(sld As Slide)
 sld.SlideShowTransition.EntryEffect = ppEffectNone
    sld.SlideShowTransition.Duration = 2
End Function


Function gotoAndExit(sld As Slide, shp As Shape)

       sld.Parent.Application.ActiveWindow.View.GotoSlide sld.SlideIndex
        shp.Select
        End
        
End Function

Function applyAnimation(shp As Shape)
 shp.AnimationSettings.Animate = msoFalse
    shp.AnimationSettings.EntryEffect = ppEffectNon
End Function

Function applyLayout(sld As Slide)
'for all layouts see https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
sld.Layout = 23


End Function
