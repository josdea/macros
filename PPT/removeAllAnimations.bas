Option Explicit


Sub RemoveAllAnimations()
'PURPOSE: Remove All PowerPoint Animations From Slides
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim sld As Slide
Dim x As Long
Dim Counter As Long

'Loop Through Each Slide in ActivePresentation
  For Each sld In ActivePresentation.Slides
    
    'Loop through each animation on slide
      For x = sld.TimeLine.MainSequence.Count To 1 Step -1
        
        'Remove Each Animation
          sld.TimeLine.MainSequence.Item(x).Delete
        
        'Maintain Deletion Stat
          Counter = Counter + 1
          
      Next x
  
  Next sld

'Completion Notification
MsgBox Counter & " Animation(s) were removed from you PowerPoint presentation!"

End Sub