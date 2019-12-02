Option Explicit
' change directory below to open edit and save all the docs in that folder

Dim fileList As String
Function main()

fileList = ""

Dim file As String
Dim path As String

' Path to your folder. MY folder is listed below. I bet yours is different.
' make SURE you include the terminating "\"
'YOU MUST EDIT THIS.
path = "c:\users\25d\desktop\bulkmacro\"

'Change this file extension to the file you are opening. .htm is listed below. You may have rtf or docx.
'YOU MUST EDIT THIS.
file = Dir(path & "*.docx")
Do While file <> ""
Documents.Open FileName:=path & file

' This is the call to the macro you want to run on each file the folder
'YOU MUST EDIT THIS. lange01 is my macro name. You put yours here.
Call doStuff

' Saves the file
ActiveDocument.Save
ActiveDocument.Close
' set file to next in Dir
file = Dir()
Loop

Call writeFile(fileList)

End Function

Function doStuff()


Dim myRange As Word.Range, myChr
For Each myRange In ActiveDocument.StoryRanges
  Do
    For Each myChr In myRange.Characters

        If myChr.Font.Superscript = True Then
            ' myChr.Font.Superscript = False
           ' myChr.InsertBefore "^"
           fileList = fileList & vbCrLf & ActiveDocument.FullName
        End If

      

    Next
    Set myRange = myRange.NextStoryRange
  Loop Until myRange Is Nothing
Next
End Function

Function writeFile(Comment As String)

Dim n As Integer
n = FreeFile()
Open Environ("USERPROFILE") & "\Desktop\filelist_comments_" & Format(Now(), "yymmdd hhmm") & "-" & ".txt" For Output As #n
'Debug.Print Comment ' write to immediate
Print #n, Comment ' write to file
Close #n

End Function
