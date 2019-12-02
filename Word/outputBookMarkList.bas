Sub BkMarkList()
    Dim J as Integer

    Selection.TypeParagraph
    Selection.InsertBreak Type:=wdColumnBreak
    Selection.TypeText Text:="Bookmark list for "
    Selection.TypeText Text:=ActiveDocument.Name
    Selection.TypeParagraph
    For J = 1 To ActiveDocument.Bookmarks.Count
        Selection.TypeText Text:=Chr(9)
        Selection.TypeText Text:=ActiveDocument.Bookmarks(J).Name
        Selection.TypeParagraph
    Next J
    Selection.InsertBreak Type:=wdColumnBreak
End Sub