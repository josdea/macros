
Function applyTransition(sld As Slide)
    sld.SlideShowTransition.EntryEffect = ppEffectNone
    sld.SlideShowTransition.Duration = 2
End Function

Function applyAnimation(shp As Shape)
    ' call from certain shapes to disbale animations
    
    shp.AnimationSettings.Animate = msoFalse
    shp.AnimationSettings.EntryEffect = ppEffectNone
End Function