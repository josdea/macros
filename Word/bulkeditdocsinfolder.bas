Option Explicit
' change directory below to open edit and save all the docs in that folder
Function main()

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

End Function

Function doStuff()


Debug.Print ActiveDocument.Name
    
    Selection.WholeStory
    Selection.LanguageID = wdEnglishUK
    Selection.NoProofing = False
    Application.CheckLanguage = False


End Function
