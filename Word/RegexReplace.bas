Sub RegexReplace()

    Dim RegEx As Object
    Set RegEx = CreateObject("VBScript.RegExp")
    
     With RegEx
            .Global = True
            .MultiLine = False
            .IgnoreCase = True
            .Pattern = InputBox("Find what:")
        End With
    
    On Error Resume Next
    
    ActiveDocument.Range = _
        RegEx.Replace(ActiveDocument.Range, InputBox("Replace with:"))
    
    Set RegEx = Nothing

End Sub
