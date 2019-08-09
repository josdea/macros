Sub inputtest()

Dim Msg, Style, Title, Help, Ctxt, Response, MyString, test
Msg = "Do you want to continue ?"    ' Define message.
Style = vbYesNo + vbQuestion + vbDefaultButton2    ' Define buttons.
Title = "MsgBox Demonstration"    ' Define title.
Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then    ' User chose Yes.
    MyString = "Yes"    ' Perform some action.
    test = MsgBox("You clicked yes", vbOKCancel, "Yes")
Else    ' User chose No.
    MyString = "No"    ' Perform some action.
        test = MsgBox("You clicked no", vbOKCancel, "No")
End If

' See https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/msgbox-function
End Sub

Sub removeAllPresenterNotes()
  Dim currentSlide As Slide
 
 If MsgBox("Are you sure you want to delete all presenter/instructor notes?", (vbYesNo + vbQuestion), "Delete all Notes?") = vbYes then

else
msgBox("Action canceled.")

  end If
End Sub

Option Explicit

Sub zoomTest()

Windows(2).View.Zoom = 75


End Sub

Sub createSections(sectionsDesired As Integer)
    Dim i                                         As Integer
    For i = 1 To sectionsDesired
        
        Debug.Print "Creating section " & i & " of " & sectionsDesired
        ActivePresentation.SectionProperties.AddBeforeSlide 1, "Module " & i
        
    Next i
    
End Sub

Sub duplicateSlide(slidesDesired As Integer, slideToDuplicate As Integer)
    Dim i                                         As Integer
    
    For i = 1 To slidesDesired
        
        Debug.Print "Creating slide " & i & " of " & slidesDesired
        ActivePresentation.Slides(slideToDuplicate).Duplicate
        
    Next i
    
End Sub

Sub resetSlideLayout(sld As Slide)
'must be on current slide
ActivePresentation.Slides(sld.SlideNumber).Select
Debug.Print "Resetting Slide Layout for Slide " & sld.SlideNumber
    Application.CommandBars.ExecuteMso ("SlideReset")
    
End Sub


Function temptesting()
    ' TODO this should detect if a slide doesn't have these shapes and then copy from slide one
    
 '   With ActivePresentation

  '  .Slides(1).Shapes.Range(Array(1, 2, 3)).Copy
 '   .Slides(2).Shapes.Paste
 '   Application.CommandBars.ExecuteMso ("SlideReset")

'End With
    
    

' createSlideShapeTextbox ActivePresentation.Slides(1), -288, 0, 252, 540, "LearnerNotes3", "Learner Notes 33", 12

End Function


Sub getShapeInfo(sld As Slide, shp As Shape)


      ' getting current slide shape info, location, and size TODO remove
        Debug.Print "Shape #" & shp.id & " (" & shp.name & ") - Slide: " & sld.SlideNumber & " Position: " & shp.Left & "," & shp.Top _
        ; " Size: " & shp.Width & "x" & shp.Height

End Sub

Sub iterateNoteShapes(sld As Slide)
    Dim shp                                       As Shape        ' declare shape object
    Dim shapeCount                                As Integer
    
    shapeCount = sld.Shapes.Count
    Debug.Print shapeCount & " Note Shapes on Slide"
    
    For Each shp In sld.NotesPage.Shapes
        
        If InStr(shp.name, "Text Placeholder") Then
            Debug.Print "Shape that follows has text placeholder in the shape name"
        End If
        
        Debug.Print (shp.name)
        
        If shp.TextFrame.HasText Then
            Debug.Print "text frame has text"
            
            Debug.Print (shp.TextFrame.TextRange)
        End If
    Next shp
    
End Sub
