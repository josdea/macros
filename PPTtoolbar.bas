Option Explicit

Sub Auto_Open()
    Dim oToolbar                                  As CommandBar
    Dim MyToolbar                                 As String
    
    ' Give the toolbar a name
    MyToolbar = "Josh        's PPT Tools"
    
    On Error Resume Next
    ' so that it doesn't stop on the next line if the toolbar's already there
    CommandBars(MyToolbar).Delete        ' delete the toolbar if it exists already in order to update it each time PPT opens
    ' Create the toolbar; PowerPoint will error if it already exists
    Set oToolbar = CommandBars.Add(name:=MyToolbar, _
    Position:=msoBarFloating, Temporary:=True)
    If Err.number <> 0 Then
        ' The toolbar's already there, so we have nothing to do
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    Dim btnSetLanguage                            As CommandBarButton
    ' Now add a button to the new toolbar
    ' And set some of the button's properties
    Set btnSetLanguage = oToolbar.Controls.Add(Type:=msoControlButton)
    With btnSetLanguage
        .DescriptionText = "Set Language of all objects and notes to US or UK"
        'Tooltip text when mouse if placed over button
        .Caption = "Set Language2"
        'Text if Text in Icon is chosen
        .OnAction = "bulkChangeSpellCheckLanguage"
        'Runs the Sub Button1() code when clicked
        .Style = msoButtonIconAndCaption
        ' Button displays                         as icon, not text or both
        .FaceId = 7385
        ' chosen icon
    End With
    ' end of button creation
    
    Dim btnRemovePresenterNotes                   As CommandBarButton
    ' Now add a button to the new toolbar
    Set btnRemovePresenterNotes = oToolbar.Controls.Add(Type:=msoControlButton)
    ' And set some of the button's properties
    With btnRemovePresenterNotes
        .DescriptionText = "Removes all Presenter Notes"
        'Tooltip text when mouse if placed over button
        .Caption = "Remove All Notes"
        'Text if Text in Icon is chosen
        .OnAction = "deletePresenterNotes"
        'Runs the Sub Button1() code when clicked
        .Style = msoButtonIconAndCaption
        ' Button displays                         as icon, not text or both
        .FaceId = 9408
        ' chosen icon
        
    End With
    ' end of button creation
    
 
    
    ' You can set the toolbar position and visibility here if you like
    ' By default, it'll be visible when created. Position will be ignored in PPT 2007 and later
    oToolbar.Top = 150
    oToolbar.Left = 150
    oToolbar.Visible = True
    
NormalExit:
    Exit Sub        ' so it doesn't go on to run the errorhandler code
    
ErrorHandler:
    'Just in case there is an error
    MsgBox Err.number & vbCrLf & Err.Description
    Resume NormalExit:
End Sub

