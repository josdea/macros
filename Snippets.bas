Sub inputtest()

Dim Msg, Style, Title, Help, Ctxt, Response, MyString, test
Msg = "Do you want to continue ?"    ' Define message.
Style = vbYesNo + vbQuestion + vbDefaultButton2    ' Define buttons.
Title = "MsgBox Demonstration"    ' Define title.
Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then    ' User chose Yes.
    MyString = "Yes"    ' Perform some action.
    test = MsgBox("You clicked yes", vbOKCancel, "Yes")
Else    ' User chose No.
    MyString = "No"    ' Perform some action.
        test = MsgBox("You clicked no", vbOKCancel, "No")
End If


' See https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/msgbox-function
End Sub

Sub removeAllPresenterNotes()
  Dim currentSlide As Slide
 
 If MsgBox("Are you sure you want to delete all presenter/instructor notes?", (vbYesNo + vbQuestion), "Delete all Notes?") = vbYes then



else
msgBox("Action canceled.")

  end If
End Sub


