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
    Dim oToolbar                                  As CommandBar
    Dim MyToolbar                                 As String
    
    ' Give the toolbar a name
    MyToolbar = "Josh's PPT Tools"
    
    On Error Resume Next
    ' so that it doesn't stop on the next line if the toolbar's already there
    CommandBars(MyToolbar).Delete                 ' delete the toolbar if it exists already in order to update it each time PPT opens
    ' Create the toolbar; PowerPoint will error if it already exists
    Set oToolbar = CommandBars.Add(name:=MyToolbar, _
    Position:=msoBarFloating, Temporary:=False)
    If Err.number <> 0 Then
        ' The toolbar's already there, so we have nothing to do
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    Call createButton("Comments_Search_And_Export", "Comments Export", "Exports all comments or based on search", 201, oToolbar)
    Call createButton("File_PPTX_Combine_All_In_Folder", "Files Combine PPTX", "Combines all PPTX files in the same folder as current file", 1679, oToolbar)
    Call createButton("Image_Add_Comment_For_Missing_AltText", "Images AltText?", "Add a comment for all slides with images having missing alternate text", 6382, oToolbar)
    Call createButton("Image_Go_To_Next", "Image Next", "Go to the next slide that has an image", 6578, oToolbar)
    Call createButton("Layout_Create_All_Types", "Layout All", "Creates an example slide of each default layout", 4389, oToolbar)
    Call createButton("Presenter_Notes_Remove_All", "Notes Remove", "Deletes all presenter notes on all slides", 7440, oToolbar)
    Call createButton("PresenterNotes_Add_Comment_If_Empty", "Notes Empty?", "Adds a comment on every slide with empty presenter notes", 7385, oToolbar)
    Call createButton("PresenterNotes_Remove_Text_In_Hashtags", "Notes Hashtags", "Removes all text wrapped in ##double hashtags## in presenter notes", 2991, oToolbar)
    Call createButton("PresenterNotes_Toggle_Visibility", "Notes Toggle", "Toggles visibility of all presenter notes on notes pages", 7803, oToolbar)
    Call createButton("Sections_Bulk_Create", "Sections Bulk", "Bulk creates sections with optional prefix and or suffix", 7460, oToolbar)
    Call createButton("Shape_Display_Type_And_Details", "Shape Details", "Provides additional details about currently selected shape or shapes", 7169, oToolbar)
    Call createButton("Shapes_Count_By_Name", "Shapes Count", "Count all shapes that have specified name", 6185, oToolbar)
    Call createButton("Shapes_Delete_By_Name", "Shapes Delete", "Deletes all shapes that have specified name", 1716, oToolbar)
    Call createButton("Shapes_Go_To_Next_Non_Placeholder", "Shapes Non Placeholder", "Go to the next slide which has a non placeholder shape", 5689, oToolbar)
    Call createButton("Slide_Go_To", "Slide Go To", "Go to a slide by number", 29, oToolbar)
    Call createButton("Text_Font_Reset_To_Master", "Text Font Reset", "Resets all text in all shapes and notes to master theme font", 2010, oToolbar)
    Call createButton("Text_Language_Toggle_US_UK", "Text UK/US", "Toggles all text in all shapes between US and UK spelling", 5768, oToolbar)
    Call createButton("Text_Remove_Double_Spaces", "Text  Spaces", "Removes all text which has double spacebars and replaces with one", 2124, oToolbar)
    Call createButton("Video_Go_To_Next", "Video Go To", "Go to the next slide that has a video", 9230, oToolbar)
    Call createButton("Video_Convert_Linked_To_Embedded", "Video Link/Embed", "Converts all linked videos to embedded", 9780, oToolbar)
    Call createButton("Video_Convert_Embedded_To_Linked", "Video Embed/Link", "Converts all embedded videos to linked", 4348, oToolbar)
    
    oToolbar.visible = True
    
NormalExit:
    Exit Sub                                      ' so it doesn't go on to run the errorhandler code
    
ErrorHandler:
    'Just in case there is an error
    MsgBox Err.number & vbCrLf & Err.Description
    Resume NormalExit:
End Sub







