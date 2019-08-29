Option Explicit

Sub exportCommentsToDisk()
Dim n As Integer
Dim sld As Slide
Dim myComment As Comment
Dim Comment As String

For Each sld In ActivePresentation.Slides
For Each myComment In sld.Comments
Comment = Comment & vbCrLf
Comment = Comment & "Slide " & sld.SlideNumber & " of " & ActivePresentation.Slides.Count & " (" & myComment.DateTime & ")" & vbCrLf
Comment = Comment & "" & myComment.Text & vbCrLf
Next myComment
Next sld

n = FreeFile()
Open "C:\Users\25d\Desktop\ppt_comments_" & Format(Now(), "yymmddhhmm") & ".txt" For Output As #n
'Debug.Print Comment ' write to immediate
Print #n, Comment ' write to file
Close #n

MsgBox "All Comments have been written to a text file (ppt_comments) on your desktop"

End Sub



