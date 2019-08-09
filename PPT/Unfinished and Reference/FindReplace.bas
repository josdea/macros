'unfinished multi find replace function works on slides only not notes
'find replace notes works with strings i think but not line breaks

Sub Multi_FindReplace()
'PURPOSE: Find & Replace a list of text/values throughout entire PowerPoint presentation
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim sld As Slide
Dim shp As Shape
Dim ShpTxt As TextRange
Dim TmpTxt As TextRange
Dim FindList As Variant
Dim ReplaceList As Variant
Dim x As Long

FindList = Array("  ", vbCrLf & vbCrLf, vbCr & vbCr, vbLf & vbLf)
ReplaceList = Array(" ", vbCrLf, vbCr, vbLf)

'Loop through each slide in Presentation
  For Each sld In ActivePresentation.Slides
    
        For Each shp In sld.Shapes
      'Store shape text into a variable
        Set ShpTxt = shp.TextFrame.TextRange
      
      'Ensure There is Text To Search Through
        If ShpTxt <> "" Then
          For x = LBound(FindList) To UBound(FindList)
            
            'Store text into a variable
              Set ShpTxt = shp.TextFrame.TextRange
            
            'Find First Instance of "Find" word (if exists)
              Set TmpTxt = ShpTxt.Replace( _
               FindWhat:=FindList(x), _
               Replacewhat:=ReplaceList(x), _
               WholeWords:=False)
'        textS = Replace(text, vbCr, ""
            'Find Any Additional instances of "Find" word (if exists)
              Do While Not TmpTxt Is Nothing
                Set ShpTxt = ShpTxt.Characters(TmpTxt.Start + TmpTxt.Length, ShpTxt.Length)
                Set TmpTxt = ShpTxt.Replace( _
                 FindWhat:=FindList(x), _
                 Replacewhat:=ReplaceList(x), _
                 WholeWords:=False)
              Loop
              
          Next x
          
        End If
        
    Next shp
      
  Next sld
MsgBox ("all done")

End Sub

Sub findReplaceNotes()

Dim oPres As Presentation
Dim oSld As Slide
Dim oShp As Shape
Dim FindWhat As String
Dim ReplaceWith As String

'FindList = Array("  ", vbCrLf & vbCrLf, vbCr & vbCr, vbLf & vbLf)
'ReplaceList = Array(" ", vbCrLf, vbCr, vbLf)

FindWhat = vbLf & vbLf
ReplaceWith = vbLf
For Each oPres In Application.Presentations
     For Each oSld In oPres.Slides
        For Each oShp In oSld.Shapes
            Call ReplaceText(oShp, FindWhat, ReplaceWith)
        Next oShp
    Next oSld
Next oPres
MsgBox (" all done")
End Sub

Sub ReplaceText(oShp As Object, FindString As String, ReplaceString As String)
Dim oTxtRng As TextRange
Dim oTmpRng As TextRange
Dim I As Integer
Dim iRows As Integer
Dim iCols As Integer
Dim oShpTmp As Shape

' Always include the 'On error resume next' statement below when you are working with text range object.
' I know of at least one PowerPoint bug where it will error out - when an image has been dragged/pasted
' into a text box. In such a case, both HasTextFrame and HasText properties will return TRUE but PowerPoint
' will throw an error when you try to retrieve the text.
On Error Resume Next
Select Case oShp.Type
Case 19 'msoTable
    For iRows = 1 To oShp.Table.Rows.Count
        For iCol = 1 To oShp.Table.Rows(iRows).Cells.Count
            Set oTxtRng = oShp.Table.Rows(iRows).Cells(iCol).Shape.TextFrame.TextRange
            Set oTmpRng = oTxtRng.Replace(FindWhat:=FindString, _
                                  Replacewhat:=ReplaceString, WholeWords:=True)
            Do While Not oTmpRng Is Nothing
            Set oTmpRng = oTxtRng.Replace(FindWhat:=FindString, _
                                Replacewhat:=ReplaceString, _
                                After:=oTmpRng.Start + oTmpRng.Length, _
                                WholeWords:=True)
            Loop
        Next
    Next
Case msoGroup 'Groups may contain shapes with text, so look within it
    For I = 1 To oShp.GroupItems.Count
        Call ReplaceText(oShp.GroupItems(I), FindString, ReplaceString)
    Next I
Case 21 ' msoDiagram
    For I = 1 To oShp.Diagram.Nodes.Count
        Call ReplaceText(oShp.Diagram.Nodes(I).TextShape, FindString, ReplaceString)
    Next I
Case Else
    If oShp.HasTextFrame Then
        If oShp.TextFrame.HasText Then
            Set oTxtRng = oShp.TextFrame.TextRange
            Set oTmpRng = oTxtRng.Replace(FindWhat:=FindString, _
                Replacewhat:=ReplaceString, WholeWords:=True)
            Do While Not oTmpRng Is Nothing
                Set oTmpRng = oTxtRng.Replace(FindWhat:=FindString, _
                            Replacewhat:=ReplaceString, _
                            After:=oTmpRng.Start + oTmpRng.Length, _
                            WholeWords:=True)
            Loop
       End If
    End If
End Select
End Sub