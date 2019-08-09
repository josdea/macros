Option Explicit


Sub combineAllPPTX()
'  Insert all slides from all presentations in the same folder as this one
'  INTO this one; do not attempt to insert THIS file into itself, though.

    Dim vArray() As String
    Dim x As Long
    Dim slideCountbeforeInsert As Integer
    slideCountbeforeInsert = ActivePresentation.Slides.Count

' If MsgBox("Are you sure you want to combine all the PPTX files in the current folder? You should be in the first file and the order of the files will be alphabetical. This means renameing them to 01, 02, 03 etc. if using numbers", (vbYesNo + vbQuestion), "Combine all?") = vbYes Then


    ' Change "*.PPT" to "*.PPTX" or whatever if necessary:
    EnumerateFiles ActivePresentation.Path & "\", "*.PPTX", vArray

If MsgBox("Are you sure you want to combine all " & UBound(vArray) & " PPTX files in the current folder? You should be in the first file and the order of the files will be alphabetical. This may require renaming them to 01, 02, 03 etc. if using numbers", (vbYesNo + vbQuestion), "Combine all?") = vbYes Then



    ActivePresentation.SectionProperties.AddBeforeSlide 1, "Module 1"

    With ActivePresentation
        For x = 1 To UBound(vArray)
             
             
             Debug.Print "x is " & x
            Debug.Print "slide count before is " & slideCountbeforeInsert
            Debug.Print "there are this many slides: " & ActivePresentation.Slides.Count
            
             
            If Len(vArray(x)) > 0 Then
                .Slides.InsertFromFile vArray(x), .Slides.Count
                
            ActivePresentation.SectionProperties.AddBeforeSlide slideCountbeforeInsert + 1, "Module " & x
            slideCountbeforeInsert = ActivePresentation.Slides.Count
            
            End If
                   
        Next
        
        MsgBox "The " & UBound(vArray) & " files have been combined. There are now " & ActivePresentation.Slides.Count & " total slides and " & ActivePresentation.SectionProperties.Count & " total sections."
        
    End With

Else
MsgBox ("Action canceled.")

  End If


End Sub

Sub EnumerateFiles(ByVal sDirectory As String, _
    ByVal sFileSpec As String, _
    ByRef vArray As Variant)
    ' collect all files matching the file spec into vArray, an array of strings

    Dim sTemp As String
    ReDim vArray(1 To 1)

    sTemp = Dir$(sDirectory & sFileSpec)
    Do While Len(sTemp) > 0
        ' NOT the "mother ship" ... current presentation
        If sTemp <> ActivePresentation.name Then
            ReDim Preserve vArray(1 To UBound(vArray) + 1)
            vArray(UBound(vArray)) = sDirectory & sTemp
        End If
        sTemp = Dir$
    Loop

End Sub







