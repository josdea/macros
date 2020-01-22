Function createAllLayoutTypes()
'call from presentation start to create blank layouts for all layouts
 Dim sld As Slide
    
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutBlank)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutBlank (" & sld.Layout & ")"
    
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutChart)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutChart (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutChartAndText)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutChartAndText (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutClipartAndText)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutClipartAndText (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutClipArtAndVerticalText)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutClipArtAndVerticalText (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutComparison)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutComparison (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutContentWithCaption)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutContentWithCaption (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutCustom)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutCustom (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutFourObjects)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutFourObjects (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutLargeObject)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutLargeObject (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutMediaClipAndText)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutMediaClipAndText (" & sld.Layout & ")"
    
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutMixed (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutObject)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutObject (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutObjectAndText)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutObjectAndText (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutObjectAndTwoObjects)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutObjectAndTwoObjects (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutObjectOverText)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutObjectOverText (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutOrgchart)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutOrgchart (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutPictureWithCaption)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutPictureWithCaption (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutSectionHeader)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutSectionHeader (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutTable)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutTable (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutText)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutText (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutTextAndChart)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutTextAndChart (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutTextAndClipart)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutTextAndClipart (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutTextAndMediaClip)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutTextAndMediaClip (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutTextAndObject)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutTextAndObject (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutTextAndTwoObjects)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutTextAndTwoObjects (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutTextOverObject)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutTextOverObject (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutTitle)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutTitle (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutTitleOnly)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutTitleOnly (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutTwoColumnText)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutTwoColumnText (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutTwoObjects)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutTwoObjects (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutTwoObjectsAndObject)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutTwoObjectsAndObject (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutTwoObjectsAndText)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutTwoObjectsAndText (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutTwoObjectsOverText)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutTwoObjectsOverText (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutVerticalText)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutVerticalText (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutVerticalTitleAndText)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutVerticalTitleAndText (" & sld.Layout & ")"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutVerticalTitleAndTextOverChart)
 sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout:  ppLayoutVerticalTitleAndTextOverChart (" & sld.Layout & ")"
    
End Function