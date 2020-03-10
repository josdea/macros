Option Explicit


Public left    As Long
Public top     As Long
Public width   As Long
Public height  As Long
Public name    As String
Public fontSize As Integer
Public text    As String
Public boolOverwriteText As Boolean
Public boolLockAspectRatio As Boolean
Public autoSize As Integer
Public boolWordWrap As Boolean
Public margBottom As Long
Public margTop As Long
Public margLeft As Long
Public margRight As Long
Public alignHorizontal As Integer
Public alignVertical As Integer
Public fillForeColor As Long
Public fontColor As Long


Private Sub Class_Initialize()
     boolOverwriteText = False
End Sub


Public Property Set oshp(shp As Shape)


     With shp
          left = .left
          top = .top
          width = .width
          height = .height
          name = .name
          boolLockAspectRatio = .LockAspectRatio
          fillForeColor = .fill.ForeColor
          
          If .HasTextFrame Then
               alignHorizontal = .TextEffect.alignment
               
               With .TextFrame
                    If .autoSize < 0 Then
                         autoSize = .autoSize * -1
                    Else
                         autoSize = .autoSize
                    End If
                    boolWordWrap = .WordWrap
                    margBottom = .MarginBottom
                    margTop = .MarginTop
                    margLeft = .MarginLeft
                    margRight = .MarginRight
                    fontSize = .TextRange.Font.Size
                    fontColor = .TextRange.Font.Color
                    alignVertical = .VerticalAnchor
               End With
          End If
     End With
     
     If shp.HasTextFrame Then
          If shp.TextFrame.HasText Then
               If MsgBox("Memorize and use text?", (vbYesNo + vbQuestion), "Use Text?") = vbYes Then
                    If MsgBox("Overwrite existing text?", (vbYesNo + vbQuestion), "Overwrite Text?") = vbYes Then
                         boolOverwriteText = True
                    End If
                    text = shp.TextFrame.TextRange.text
               End If                             'prompt for text
          End If                                  ' has text
     End If                                       ' has text frame
     
End Property


