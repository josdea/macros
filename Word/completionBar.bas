Private Function completionBar(num As Integer, den As Integer)
Dim strLine As String
num = 100 / den * num
den = 100

Dim i As Integer
For i = 1 To num
strLine = strLine & "*"
Next i
For i = num To den
    strLine = strLine & "_"
Next i
Debug.Print strLine & "|"

End Function