Option Explicit

Private Function createButton(strOnAction As String, strCaption As String, strTooltip As String, intIcon As Integer, oToolbar As CommandBar)
    Dim btnButton As CommandBarButton
    Set btnButton = oToolbar.Controls.Add(Type:=msoControlButton)
    With btnButton
        .Caption = strCaption
        .OnAction = strOnAction
        .Style = msoButtonIconAndCaption          'change to msoButtonIconAndCaption to use the icon below
        .FaceId = intIcon
        .TooltipText = strTooltip
    End With
End Function

Sub Auto_Open()
    Dim oToolbarComment                                  As CommandBar
    Dim oToolbarNavigation                                  As CommandBar
    Dim oToolbarText                                  As CommandBar
    Dim oToolbarUtility                                 As CommandBar
    Dim oToolbarMedia                                  As CommandBar
    Dim currentToolbar As CommandBar


    On Error Resume Next
    ' so that it doesn't stop on the next line if the toolbar's already there
    CommandBars("1").Delete                       ' delete the toolbar if it exists already in order to update it each time PPT opens
    CommandBars("2").Delete                       ' delete the toolbar if it exists already in order to update it each time PPT opens
    CommandBars("3").Delete                       ' delete the toolbar if it exists already in order to update it each time PPT opens
    CommandBars("4").Delete                       ' delete the toolbar if it exists already in order to update it each time PPT opens
    CommandBars("5").Delete                       ' delete the toolbar if it exists already in order to update it each time PPT opens
    CommandBars("6").Delete                       ' delete the toolbar if it exists already in order to update it each time PPT opens
    CommandBars("7").Delete                       ' delete the toolbar if it exists already in order to update it each time PPT opens
    CommandBars("Josh's PPT Tools").Delete
    
    If Err.number <> 0 Then
        ' The toolbar's already there, so we have nothing to do
        '    Exit Sub
    End If
    
    '    On Error GoTo ErrorHandler
    
    Set currentToolbar = CommandBars.Add(Name:="1", Position:=msoBarFloating, Temporary:=True)
    Call createButton("Comments_Search_And_Export", "Comments Export", "Exports all comments or based on search", 201, currentToolbar)
    Call createButton("File_PPTX_Combine_All_In_Folder", "Files Combine PPTX", "Combines all PPTX files in the same folder as current file", 1679, currentToolbar)
    Call createButton("Image_Add_Comment_For_Missing_AltText", "Images AltText?", "Add a comment for all slides with images having missing alternate text", 6382, currentToolbar)
    currentToolbar.Visible = True
    Set currentToolbar = CommandBars.Add(Name:="2", Position:=msoBarFloating, Temporary:=True)
    Call createButton("Image_Go_To_Next", "Image Next", "Go to the next slide that has an image", 6578, currentToolbar)
    Call createButton("Layout_Create_All_Types", "Layout All", "Creates an example slide of each default layout", 4389, currentToolbar)
    Call createButton("Presenter_Notes_Remove_All", "Notes Remove", "Deletes all presenter notes on all slides", 7440, currentToolbar)
    currentToolbar.Visible = True
    Set currentToolbar = CommandBars.Add(Name:="3", Position:=msoBarFloating, Temporary:=True)
    Call createButton("PresenterNotes_Add_Comment_If_Empty", "Notes Empty?", "Adds a comment on every slide with empty presenter notes", 7385, currentToolbar)
    Call createButton("PresenterNotes_Remove_Text_In_Hashtags", "Notes Hashtags", "Removes all text wrapped in ##double hashtags## in presenter notes", 2991, currentToolbar)
    Call createButton("PresenterNotes_Toggle_Visibility", "Notes Toggle", "Toggles visibility of all presenter notes on notes pages", 7803, currentToolbar)
    
    currentToolbar.Visible = True
    Set currentToolbar = CommandBars.Add(Name:="4", Position:=msoBarFloating, Temporary:=True)
    Call createButton("Create_Progress_Bar", "Progress Bar Create", "Creates a progress bar from slide 2 until the end. Can be deleted using the Shape Delete button", 6796, currentToolbar)
    Call createButton("Sections_Bulk_Create", "Sections Bulk", "Bulk creates sections with optional prefix and or suffix", 7460, currentToolbar)
    Call createButton("Shape_Display_Type_And_Details", "Shape Details", "Provides additional details about currently selected shape or shapes", 7169, currentToolbar)
    
    currentToolbar.Visible = True
    Set currentToolbar = CommandBars.Add(Name:="5", Position:=msoBarFloating, Temporary:=True)
    Call createButton("Shapes_Count_By_Name", "Shapes Count", "Count all shapes that have specified name", 6185, currentToolbar)
    Call createButton("Shapes_Delete_By_Name", "Shapes Delete", "Deletes all shapes that have specified name", 1716, currentToolbar)
    Call createButton("Shapes_Go_To_Next_Non_Placeholder", "Shapes Non Placeholder Next", "Go to the next slide which has a non placeholder shape", 5689, currentToolbar)
    
    currentToolbar.Visible = True
    Set currentToolbar = CommandBars.Add(Name:="6", Position:=msoBarFloating, Temporary:=True)
    Call createButton("Slide_Go_To", "Slide Go To", "Go to a slide by number", 29, currentToolbar)
    Call createButton("Text_Font_Reset_To_Master", "Text Font Reset", "Resets all text in all shapes and notes to master theme font", 2010, currentToolbar)
    Call createButton("Text_Language_Toggle_US_UK", "Text Language", "Toggles all text in all shapes between US and UK spelling", 5768, currentToolbar)
    
    currentToolbar.Visible = True
    Set currentToolbar = CommandBars.Add(Name:="7", Position:=msoBarFloating, Temporary:=True)
    'Call createButton("Video_Convert_Embedded_To_Linked", "Video Embed/Link", "Converts all embedded videos to linked", 4348, currentToolbar)
    Call createButton("Text_Remove_Double_Spaces", "Text  Spaces", "Removes all text which has double spacebars and replaces with one", 2124, currentToolbar)
    Call createButton("Video_Convert_Linked_To_Embedded", "Video Link/Embed", "Converts all linked videos to embedded", 9780, currentToolbar)
    Call createButton("Video_Go_To_Next", "Video Go To", "Go to the next slide that has a video", 9230, currentToolbar)
    currentToolbar.Visible = True
    
NormalExit:
    Exit Sub                                      ' so it doesn't go on to run the errorhandler code
    
ErrorHandler:
    'Just in case there is an error
    MsgBox Err.number & vbCrLf & Err.Description
    Resume NormalExit:
End Sub











