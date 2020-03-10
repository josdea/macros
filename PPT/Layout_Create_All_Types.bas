Option Explicit
Sub Layout_Create_All_Types()                     ' checked 1/17/20
    'Creates an example slide of each default layout
    Dim sld    As Slide
    Dim layout As Integer

    'For layout = 1 To 36
    '  Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=layout)
    '    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: " & sld.layout
    'Next layout

    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutTitle)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutTitle (1) - Title"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutText (2) - Text"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutTwoColumnText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutTwoColumnText (3) - Two-column text"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutTable)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutTable (4) - Table"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutTextAndChart)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutTextAndChart (5) - Text and chart"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutChartAndText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutChartAndText (6) - Chart and text"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutOrgchart)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutOrgchart (7) - Organization chart"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutChart)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutChart (8) - Chart"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutTextAndClipart)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutTextAndClipArt (9) - Text and ClipArt"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutClipartAndText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutClipArtAndText (10) - ClipArt and text"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutTitleOnly)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutTitleOnly (11) - Title only"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutBlank)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutBlank (12) - Blank"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutTextAndObject)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutTextAndObject (13) - Text and object"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutObjectAndText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutObjectAndText (14) - Object and text"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutLargeObject)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutLargeObject (15) - Large object"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutObject)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutObject (16) - Object"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutTextAndMediaClip)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutTextAndMediaClip (17) - Text and MediaClip"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutMediaClipAndText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutMediaClipAndText (18) - MediaClip and text"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutObjectOverText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutObjectOverText (19) - Object over text"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutTextOverObject)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutTextOverObject (20) - Text over object"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutTextAndTwoObjects)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutTextAndTwoObjects (21) - Text and two objects"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutTwoObjectsAndText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutTwoObjectsAndText (22) - Two objects and text"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutTwoObjectsOverText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutTwoObjectsOverText (23) - Two objects over text"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutFourObjects)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutFourObjects (24) - Four objects"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutVerticalText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutVerticalText (25) - Vertical text"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutClipArtAndVerticalText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutClipArtAndVerticalText (26) - ClipArt and vertical text"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutVerticalTitleAndText)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutVerticalTitleAndText (27) - Vertical title and text"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutVerticalTitleAndTextOverChart)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutVerticalTitleAndTextOverChart (28) - Vertical title and text over chart"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutTwoObjects)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutTwoObjects (29) - Two objects"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutObjectAndTwoObjects)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutObjectAndTwoObjects (30) - Object and two objects"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutTwoObjectsAndObject)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutTwoObjectsAndObject (31) - Two objects and object"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutCustom)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutCustom (32) - Custom"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutSectionHeader)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutSectionHeader (33) - Section header"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutComparison)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutComparison (34) - Comparison"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutContentWithCaption)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutContentWithCaption (35) - Content with caption"
    Set sld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, layout:=ppLayoutPictureWithCaption)
    sld.NotesPage.Shapes(2).TextFrame2.TextRange.Text = "Layout: ppLayoutPictureWithCaption (36) - Picture with caption"
    MsgBox "36 Slides have been created"
End Sub