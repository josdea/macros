Option Explicit

Sub exportCommentsToDiskWSearch()
Dim n As Integer
Dim sld As Slide
Dim myComment As Comment
Dim Comment As String
Dim commentSearch As String

commentSearch = InputBox("Enter Search term within comments or leave blank for all")


For Each sld In ActivePresentation.Slides
For Each myComment In sld.Comments

If (InStr(UCase(myComment.Text), UCase(commentSearch))) Or commentSearch = "" Then

Comment = Comment & vbCrLf
Comment = Comment & "Slide " & sld.SlideNumber & " of " & ActivePresentation.Slides.Count & " (" & myComment.DateTime & ")" & vbCrLf
Comment = Comment & "" & myComment.Text & vbCrLf

End If

Next myComment
Next sld

n = FreeFile()
Open Environ("USERPROFILE") & "\Desktop\ppt_comments_" & Format(Now(), "yymmddhhmm") & "-" & commentSearch & ".txt" For Output As #n
'Debug.Print Comment ' write to immediate
Print #n, Comment ' write to file
Close #n

MsgBox "All Comments have been written to a text file (ppt_comments) on your desktop"

End Sub





