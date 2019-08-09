Sub TableApplyStyles()
'
' TableApplyStyles Macro
'
' show visual basic editor
    ShowVisualBasicEditor = True
    ' find any that have colons

    ' clear formatting from find replace box
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    ' set style of replacement box
    Selection.Find.Replacement.Style = ActiveDocument.Styles( _
        "Slide Purpose Title")

        'search for purpose with a colon and replace without and add style
    With Selection.Find
        .Text = "Slide Purpose:"
        .Replacement.Text = "Slide Purpose"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' search for purpose and add style to it
    With Selection.Find
        .Text = "Slide Purpose"
        .Replacement.Text = "Slide Purpose"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' find instructor notes title with colon and add style
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles( _
        "Instructor Notes Title")
    With Selection.Find
        .Text = "Instructor Notes:"
        .Replacement.Text = "Instructor Notes"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' find instructor notes without colon and add style
    With Selection.Find
        .Text = "Instructor Notes"
        .Replacement.Text = "Instructor Notes"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    ' show navigation panel window
    ActiveWindow.DocumentMap = True
End Sub