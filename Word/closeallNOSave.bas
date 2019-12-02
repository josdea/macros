Option Explicit



Sub Close_All_Documents_NO_Save()

If (MsgBox("Do you want to close all documents WITHOUT saving?", (vbYesNo + vbQuestion), "Close All") = vbYes) Then
Application.Quit SaveChanges:=wdDoNotSaveChanges
End If

End Sub


