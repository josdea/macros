Attribute VB_Name = "NewMacros"
Sub CreateTable()
Attribute CreateTable.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.CreateTable"
'
' CreateTable Macro
'
'

Dim tableNew As Table
Dim inputTest As String
inputTest = InputBox("Type some text")

    ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=3, NumColumns:= _
        3, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
        Set tableNew = Selection.Tables(1)
    With Selection.Tables(1)
        If .Style <> "Table Grid" Then
            .Style = "Table Grid"
        End If
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = False
        .ApplyStyleRowBands = True
        .ApplyStyleColumnBands = False
        .Cell(1, 3).Range.InsertAfter inputTest
        
     
    End With
   
End Sub
