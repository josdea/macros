
createSlide(strTopicTitle As String, strObjective As String, strSlideText As String, strPGNotes As String, strIGNotes As String, strHasExercise As String, strExerciseTitle As String, strExerciseDescription As String, strMediaRequired As String, strMediaDetails As String)
strModule, strSubtitle, strDescription, strInstructor, strModuleDuration, 
Dim sld as Slide
trModule 
As String, strSubtitle 
As String, strDescription 
As String, strInstructor 
As String, strModuleDuration 
As String, strTopicTitle 
As String, strObjective 
As String, strSlideText 
As String, strPGNotes 
As String, strIGNotes 
As String, strHasExercise 
As String, strExerciseTitle 
As String, strExerciseDescription 
As String, strMediaRequired 
As String, strMediaDetails 
As String, oPresentation 
As Object)
vbCrLf
ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutBlank
   
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutChart
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutChartAndText
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutClipartAndText
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutClipArtAndVerticalText
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutComparison
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutContentWithCaption
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutCustom
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutFourObjects
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutLargeObject
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutMediaClipAndText
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutMixed
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutObject
ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutObjectAndText
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutObjectAndTwoObjects
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutObjectOverText
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutOrgchart
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutPictureWithCaption
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutSectionHeader
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutTable
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutText
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutTextAndChart
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutTextAndClipart
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutTextAndMediaClip
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutTextAndObject
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutTextAndTwoObjects
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutTextOverObject
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutTitle
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutTitleOnly
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutTwoColumnText
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutTwoObjects
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutTwoObjectsAndObject
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutTwoObjectsAndText
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutTwoObjectsOverText
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutVerticalText
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutVerticalTitleAndText
 ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutVerticalTitleAndTextOverChart





 sld.Layout = ppLayoutBlank
    
 sld.Layout = ppLayoutChart
 sld.Layout = ppLayoutChartAndText
 sld.Layout = ppLayoutClipartAndText
 sld.Layout = ppLayoutClipArtAndVerticalText
 sld.Layout = ppLayoutComparison
 sld.Layout = ppLayoutContentWithCaption
 sld.Layout = ppLayoutCustom
 sld.Layout = ppLayoutFourObjects
 sld.Layout = ppLayoutLargeObject
 sld.Layout = ppLayoutMediaClipAndText
 sld.Layout = ppLayoutMixed
 sld.Layout = ppLayoutObject
    sld.Layout = ppLayoutObjectAndText
 sld.Layout = ppLayoutObjectAndTwoObjects
 sld.Layout = ppLayoutObjectOverText
 sld.Layout = ppLayoutOrgchart
 sld.Layout = ppLayoutPictureWithCaption
 sld.Layout = ppLayoutSectionHeader
 sld.Layout = ppLayoutTable
 sld.Layout = ppLayoutText
 sld.Layout = ppLayoutTextAndChart
 sld.Layout = ppLayoutTextAndClipart
 sld.Layout = ppLayoutTextAndMediaClip
 sld.Layout = ppLayoutTextAndObject
 sld.Layout = ppLayoutTextAndTwoObjects
 sld.Layout = ppLayoutTextOverObject
 sld.Layout = ppLayoutTitle
 sld.Layout = ppLayoutTitleOnly
 sld.Layout = ppLayoutTwoColumnText
 sld.Layout = ppLayoutTwoObjects
 sld.Layout = ppLayoutTwoObjectsAndObject
 sld.Layout = ppLayoutTwoObjectsAndText
 sld.Layout = ppLayoutTwoObjectsOverText
 sld.Layout = ppLayoutVerticalText
 sld.Layout = ppLayoutVerticalTitleAndText
 sld.Layout = ppLayoutVerticalTitleAndTextOverChart