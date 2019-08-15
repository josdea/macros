Option Explicit

Sub Auto_Open()
    Dim oToolbar                                  As CommandBar
    Dim MyToolbar                                 As String
    
    ' Give the toolbar a name
    MyToolbar = "Josh's PPT Tools"
    
    On Error Resume Next
    ' so that it doesn't stop on the next line if the toolbar's already there
    CommandBars(MyToolbar).Delete        ' delete the toolbar if it exists already in order to update it each time PPT opens
    ' Create the toolbar; PowerPoint will error if it already exists
    Set oToolbar = CommandBars.Add(name:=MyToolbar, _
    Position:=msoBarFloating, Temporary:=False)
    If Err.number <> 0 Then
        ' The toolbar's already there, so we have nothing to do
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
     Dim btnbulkCreateSections                            As CommandBarButton
    Set btnbulkCreateSections = oToolbar.Controls.Add(Type:=msoControlButton)
    With btnbulkCreateSections
        .Caption = "Bulk Create Sections"
        .OnAction = "bulkCreateSection"
        .Style = msoButtonIconAndCaption        'change to msoButtonIconAndCaption to use the icon below
        .FaceId = 7385
    End With
    
    
    Dim btnSetLanguage                            As CommandBarButton
    Set btnSetLanguage = oToolbar.Controls.Add(Type:=msoControlButton)
    With btnSetLanguage
        .Caption = "Bulk Set Language"
        .OnAction = "bulkChangeSpellCheckLanguage"
        .Style = msoButtonIconAndCaption        'change to msoButtonIconAndCaption to use the icon below
        .FaceId = 7385
    End With
    
    Dim btncombineAllPPTX                            As CommandBarButton
    Set btncombineAllPPTX = oToolbar.Controls.Add(Type:=msoControlButton)
    With btncombineAllPPTX
        .Caption = "Combine PPTX"
        .OnAction = "combineAllPPTX"
        .Style = msoButtonIconAndCaption        'change to msoButtonIconAndCaption to use the icon below
        .FaceId = 7385
    End With
    
    Dim btngoToNextVideo                            As CommandBarButton
    Set btngoToNextVideo = oToolbar.Controls.Add(Type:=msoControlButton)
    With btngoToNextVideo
        .Caption = "Go to Next Video"
        .OnAction = "goToNextVideo"
        .Style = msoButtonIconAndCaption        'change to msoButtonIconAndCaption to use the icon below
        .FaceId = 7385
    End With

   Dim btnRemoveDoubleSpaces                            As CommandBarButton
    Set btnRemoveDoubleSpaces = oToolbar.Controls.Add(Type:=msoControlButton)
    With btnRemoveDoubleSpaces
        .Caption = "Remove Double Spaces"
        .OnAction = "removeDoubleSpaces"
        .Style = msoButtonIconAndCaption        'change to msoButtonIconAndCaption to use the icon below
        .FaceId = 7385
    End With

     Dim btnCountWords                            As CommandBarButton
    Set btnCountWords = oToolbar.Controls.Add(Type:=msoControlButton)
    With btnCountWords
        .Caption = "Count Slide Words"
        .OnAction = "countWords"
        .Style = msoButtonIconAndCaption        'change to msoButtonIconAndCaption to use the icon below
        .FaceId = 7385
    End With

    Dim btncreateSlideTextboxes                            As CommandBarButton
    Set btncreateSlideTextboxes = oToolbar.Controls.Add(Type:=msoControlButton)
    With btncreateSlideTextboxes
        .Caption = "Create Learner Textboxes on Slides"
        .OnAction = "createSlideTextBoxesAllSlides"
        .Style = msoButtonIconAndCaption        'change to msoButtonIconAndCaption to use the icon below
        .FaceId = 7385
    End With
    
    Dim btnCreateNoteTextBoxes                            As CommandBarButton
    Set btnCreateNoteTextBoxes = oToolbar.Controls.Add(Type:=msoControlButton)
    With btnCreateNoteTextBoxes
        .Caption = "Create Learner Textboxes on Notes"
        .OnAction = "createNoteTextBoxesAllSlides"
        .Style = msoButtonIconAndCaption        'change to msoButtonIconAndCaption to use the icon below
        .FaceId = 7385
    End With
    
    Dim btnCopySlideTextboxes                            As CommandBarButton
    Set btnCopySlideTextboxes = oToolbar.Controls.Add(Type:=msoControlButton)
    With btnCopySlideTextboxes
        .Caption = "Copy Learner box Slides to Notes"
        .OnAction = "copySlideTextBoxes"
        .Style = msoButtonIconAndCaption        'change to msoButtonIconAndCaption to use the icon below
        .FaceId = 7385
    End With
    
    Dim btnTogglePresenterNotes                            As CommandBarButton
    Set btnTogglePresenterNotes = oToolbar.Controls.Add(Type:=msoControlButton)
    With btnTogglePresenterNotes
        .Caption = "Toggle Presenter Notes"
        .OnAction = "togglePresenterNotes"
        .Style = msoButtonIconAndCaption        'change to msoButtonIconAndCaption to use the icon below
        .FaceId = 7385
    End With
    
    ' You can set the toolbar position and visibility here if you like
    ' By default, it'll be visible when created. Position will be ignored in PPT 2007 and later
    oToolbar.Top = 150
    oToolbar.Left = 150
    oToolbar.visible = True
    
NormalExit:
    Exit Sub        ' so it doesn't go on to run the errorhandler code
    
ErrorHandler:
    'Just in case there is an error
    MsgBox Err.number & vbCrLf & Err.Description
    Resume NormalExit:
End Sub

