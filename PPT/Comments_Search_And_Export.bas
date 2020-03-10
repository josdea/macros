Option Explicit
Sub Comments_Search_And_Export()                  ' checked 1/17/20
    'Exports all comments or based on search
    Dim replyCount As Integer
    Dim commentCount As Integer
    Dim sld    As Slide
    Dim myComment As Comment
    Dim Comment As String
    Dim commentSearch As String
    Dim reply  As String
    Dim replyIndex As Integer
    Dim tempComment As String
    commentSearch = InputBox("Enter Search term within comments Or leave blank For all")
    For Each sld In ActivePresentation.Slides
        For Each myComment In sld.Comments
            tempComment = "Slide " & sld.SlideNumber & " of " & ActivePresentation.Slides.count & vbCrLf
            tempComment = tempComment & "  " & myComment.Text & " (" & myComment.DateTime & ")" & vbCrLf
            replyCount = 0
            If myComment.Replies.count > 0 Then
                For replyIndex = 1 To myComment.Replies.count
                    tempComment = tempComment & "    \-" & myComment.Replies(replyIndex).Text & " (" & myComment.Replies(replyIndex).DateTime & ")" & vbCrLf
                    replyCount = replyCount + 1
                Next                              'next reply
            End If                                'there are replies
            If (InStr(UCase(tempComment), UCase(commentSearch))) Or commentSearch = "" Then
                Comment = Comment & tempComment
                Comment = Comment & vbCrLf
                commentCount = commentCount + 1 + replyCount
            End If                                'search matches or is blank
        Next myComment
    Next sld
    
    If commentCount > 0 Then                      ' there were comments
        Dim n      As Integer
        n = FreeFile()
        Open Environ("USERPROFILE") & "\Desktop\ppt_comments_" & Format(Now(), "yymmdd hhmm") & "-" & commentSearch & ".txt" For Output As #n
        Print #n, Comment                         ' write to file
        Close #n
        
        MsgBox commentCount & " comments and/or replies have been written To a Text file (ppt_comments) On your desktop"
    Else
        MsgBox "There were no comments found"
    End If
End Sub