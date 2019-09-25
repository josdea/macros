Option Explicit

Sub exportCommentsToDisk()
Dim replyCount As Integer
Dim commentCount As Integer
Dim sld As Slide
Dim myComment As Comment
Dim Comment As String
Dim commentSearch As String
Dim reply As String
Dim replyIndex As Integer
Dim tempComment As String

commentSearch = InputBox("Enter Search term within comments or leave blank for all")

For Each sld In ActivePresentation.Slides
For Each myComment In sld.Comments
tempComment = "Slide " & sld.SlideNumber & " of " & ActivePresentation.Slides.Count & vbCrLf
tempComment = tempComment & "  " & myComment.text & " (" & myComment.DateTime & ")" & vbCrLf

replyCount = 0
If myComment.Replies.Count > 0 Then
For replyIndex = 1 To myComment.Replies.Count
tempComment = tempComment & "    \-" & myComment.Replies(replyIndex).text & " (" & myComment.Replies(replyIndex).DateTime & ")" & vbCrLf
replyCount = replyCount + 1
Next 'next reply
End If 'there are replies

If (InStr(UCase(tempComment), UCase(commentSearch))) Or commentSearch = "" Then
Comment = Comment & tempComment
Comment = Comment & vbCrLf
commentCount = commentCount + 1 + replyCount
End If 'search matches or is blank

Next myComment

Next sld

Call writeFile(Comment, commentSearch)
MsgBox commentCount & " comments and/or replies have been written to a text file (ppt_comments) on your desktop"

End Sub

Function writeFile(Comment As String, commentSearch As String)

Dim n As Integer
n = FreeFile()
Open Environ("USERPROFILE") & "\Desktop\ppt_comments_" & Format(Now(), "yymmdd hhmm") & "-" & commentSearch & ".txt" For Output As #n
'Debug.Print Comment ' write to immediate
Print #n, Comment ' write to file
Close #n

End Function




