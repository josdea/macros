Option Explicit

Sub PPTNotes_To_Excel()
  Dim oSH      As Object
  Dim oPres    As Presentation
  Dim sld      As Slide
  Dim shp      As Shape
  Dim cl       As Object

  Set oSH = createNewExcel()
  Set oPres = ActivePresentation
  Set cl = oSH.Range("A1")
  cl.Value = "Slide"
  Set cl = oSH.Range("B1")
  cl.Value = "Image"
  Set cl = oSH.Range("C1")
  cl.Value = "Notes"
  
  For Each sld In oPres.Slides
    Set cl = oSH.Range("A" & sld.SlideNumber + 1)
    cl.Value = sld.SlideNumber
    Set cl = oSH.Range("B" & sld.SlideNumber + 1)
    cl.Value = "Slide" & sld.SlideNumber & ".GIF" ' CHANGE TO PNG IF DESIRED
    If sld.HasNotesPage Then
      For Each shp In sld.NotesPage.Shapes
        If shp.Type = msoPlaceholder Then
          If shp.PlaceholderFormat.Type = ppPlaceholderBody Then
            Set cl = oSH.Range("C" & sld.SlideNumber + 1)
            cl.Value = shp.TextFrame.TextRange.Text
            Exit For
          End If
        End If
      Next shp
    End If
  Next sld
  
End Sub

Private Function createNewExcel() As Object
  Dim oXLS     As Object
  Set oXLS = CreateObject("Excel.Application")
  oXLS.Visible = True
  Dim oWB      As Object
  Set oWB = oXLS.Workbooks.Add
  Set createNewExcel = oWB.Sheets(1)
End Function



