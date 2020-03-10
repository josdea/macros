Option Explicit
Sub File_PPTX_Combine_All_In_Folder()
    'Combines all PPTX files in the same folder as current file
    Dim vArray() As String
    Dim x      As Long
    Dim slideCountbeforeInsert As Integer
    slideCountbeforeInsert = ActivePresentation.Slides.count
    EnumerateFiles ActivePresentation.Path & "\", "*.PPTX", vArray
    If MsgBox("Are you sure you want To combine all " & UBound(vArray) & " PPTX files in the current folder? You should be in the first file And the order of the files will be alphabetical. This may require renaming them To 01, 02, 03 etc. If using numbers", (vbYesNo + vbQuestion), "Combine all?") = vbYes Then
        ActivePresentation.SectionProperties.AddBeforeSlide 1, "Module 1"
        With ActivePresentation
            For x = 1 To UBound(vArray)
                If Len(vArray(x)) > 0 Then
                    .Slides.InsertFromFile vArray(x), .Slides.count
                    ActivePresentation.SectionProperties.AddBeforeSlide slideCountbeforeInsert + 1, "Module " & x
                    slideCountbeforeInsert = ActivePresentation.Slides.count
                End If
            Next
            MsgBox "The " & UBound(vArray) & " files have been combined. There are now " & ActivePresentation.Slides.count & " total slides And " & ActivePresentation.SectionProperties.count & " total sections."
        End With
    Else
        MsgBox ("Action canceled.")
    End If
End Sub

Private Function EnumerateFiles(ByVal sDirectory As String, _
        ByVal sFileSpec As String, _
        ByRef vArray As Variant)
    Dim sTemp  As String
    ReDim vArray(1 To 1)
    sTemp = Dir$(sDirectory & sFileSpec)
    Do While Len(sTemp) > 0
        If sTemp <> ActivePresentation.Name Then
            ReDim Preserve vArray(1 To UBound(vArray) + 1)
            vArray(UBound(vArray)) = sDirectory & sTemp
        End If
        sTemp = Dir$
    Loop
End Function