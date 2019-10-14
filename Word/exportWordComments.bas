Option Explicit

'Macro last updated 2019-10-4 by Josh Dean deanjm@ornl.gov

Sub searchAndExportComments()
    Dim doc    As Document
    Set doc = ActiveDocument
    
    Dim intCommentCount As Integer
    Dim oComment As Comment
    Dim strCommentsData As String
    Dim strCommentSearch As String
    
    strCommentsData = "Number,Comment Text,Date,Resolved,Author" & vbCrLf
    strCommentSearch = InputBox("Enter Search term within comments Or leave blank For all")
    
    For Each oComment In doc.Comments
        
        If (InStr(UCase(oComment.Range.Text), UCase(strCommentSearch))) Or strCommentSearch = "" Then
            intCommentCount = intCommentCount + 1
            strCommentsData = strCommentsData & intCommentCount & "," & oComment.Range.Text & "," & oComment.Date & ","
            If oComment.Done = True Then
                strCommentsData = strCommentsData & "true,"
            Else
                strCommentsData = strCommentsData & "false,"
            End If
            strCommentsData = strCommentsData & Replace(oComment.Author, ",", " ") & vbCrLf
        End If        'search matches or is blank
        
    Next oComment
    
    If strCommentSearch = "" Then
        strCommentSearch = "ALL"
    End If
    
    Call writeFile(strCommentsData, strCommentSearch)
    MsgBox intCommentCount & " comments and/or replies have been written To a csv file (word_comments) On your desktop"
End Sub
Private Function writeFile(strCommentsData As String, strCommentSearch As String)
    
    Dim n      As Integer
    n = FreeFile()
    Open Environ("USERPROFILE") & "\Desktop\word_comments_" & Format(Now(), "yymmdd hhmmss") & "-" & strCommentSearch & ".csv" For Output As #n
    'Debug.Print strCommentsData ' write to immediate
    Print #n, strCommentsData        ' write to file
    Close #n
End Function
