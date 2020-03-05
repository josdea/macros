Option Explicit

Sub Export_To_Excel()
  Dim oSH      As Object
  Set oSH = createNewExcel()
  Dim oPres    As Presentation
  Set oPres = ActivePresentation
  Call loopSlides(oSH, oPres)
  Set oSH = Nothing
  Set oPres = Nothing
End Sub

Private Function createNewExcel() As Object
  Dim oXLS     As Object
  Set oXLS = CreateObject("Excel.Application")
  oXLS.Visible = True
  Dim oWB      As Object
  Set oWB = oXLS.Workbooks.Add
  Set createNewExcel = oWB.Sheets(1)
End Function

Private Function loopSlides(oSH As Object, oPres As Presentation)
  Dim sld      As Slide
  Dim shp      As Shape
  Dim cl       As Object
  Set cl = oSH.Range("A1")
  cl.Value = "Slide"
  Set cl = oSH.Range("B1")
  cl.Value = "Image"                              ' Slide4.GIF
  Set cl = oSH.Range("C1")
  cl.Value = "Notes"
  
  For Each sld In oPres.Slides
    Set cl = oSH.Range("A" & sld.SlideNumber + 1)
    cl.Value = sld.SlideNumber
    Set cl = oSH.Range("B" & sld.SlideNumber + 1)
    cl.Value = "Slide" & sld.SlideNumber & ".GIF" ' CHANGE TO PNG IF DESIRED
    If sld.HasNotesPage Then                      ' slide has notes
      For Each shp In sld.NotesPage.Shapes
        If shp.Type = msoPlaceholder Then         ' is a placeholder
          If shp.PlaceholderFormat.Type = ppPlaceholderBody Then 'is presenter notes
            Set cl = oSH.Range("C" & sld.SlideNumber + 1)
            cl.Value = shp.TextFrame.TextRange.Text
            Exit For
          End If
        End If
      Next shp
    End If
  Next sld
End Function



