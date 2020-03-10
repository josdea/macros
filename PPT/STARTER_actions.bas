Dim counter                                       As Integer
Option Explicit
Function presentationActionsStart(currentPresentation As Presentation)
    counter = 0        'USE THE FOLLOWING IN ANY PLACE TO AUGMENT COUNTER or type "End" to stop at next instancce "counter = counter + 1"
    'ENTIRE PRESENTATION - CALLED FIRST
End Function
Function presentationActionsEnd(currentPresentation As Presentation)
    'MsgBox "Counter: " & counter     'USE THIS LINE TO COUNT INSTANCES
    'ENTIRE PRESENTATION - CALLED LAST
    
End Function
Function sectionStartAction(currentPresentation As Presentation, sld As Slide)
    'FIRST SLIDE AT THE BEGINNING A NEW SECTION
    
End Function
Function sectionEndAction(currentPresentation As Presentation, sld As Slide)
    'LAST SLIDE AT THE END OF SECTION
    
End Function
Function slideActions(sld As Slide)
    'EACH SLIDE
    
    If sld.Master.Name = "SlideMaster" Then
        'EACH SLIDE USING DEFAULT MASTER LAYOUT
        
        Select Case sld.layout
            Case Is = ppLayoutTitle        'TITLE SLIDE
                
            Case Is = ppLayoutObject        'TITLE AND CONTENT
                
            Case Is = ppLayoutSectionHeader        'SECTION HEADER
                
            Case Is = ppLayoutTwoObjects        'TWO CONTENT OBJECTS
                
            Case Is = ppLayoutComparison        'COMPARISON - TWO OBJECTS TWO TEXT BOXES
                
            Case Is = ppLayoutTitleOnly        'JUST TITLE
                
            Case Is = ppLayoutBlank        'NO PLACEHOLDERS
                
            Case Is = ppLayoutContentWithCaption        'TITLE, CONTENT, AND TEXTBOX
                
            Case Is = ppLayoutPictureWithCaption        ' TITLE, PICTURE, AND TEXTBOX
                
        End Select
    Else
        'EACH SLIDE USING A NON DEFAULT MASTER LAYOUT
        sld.Select
        
    End If
    
End Function
Function slideShapeActions(sld As Variant, shp As Variant)
    'EACH SHAPE OF A SLIDE
    
    If shp.Type = msoPlaceholder Then
        'EACH SLIDE PlACEHOLDER SHAPE
        
        Select Case shp.PlaceholderFormat.Type
            Case Is = ppPlaceholderTitle        'EACH TITLE PLACEHOLDER
             
            Case Is = ppPlaceholderObject        'EACH CONTENT PLACEHOLDER
                
            Case Is = ppPlaceholderDate        'EACH DATE PLACEHOLDER
                
            Case Is = ppPlaceholderFooter        'EACH FOOTER PLACEHOLDER
                
            Case Is = ppPlaceholderSlideNumber        'EACH SLIDE NUMBER PLACEHOLDER
                
            Case Is = ppPlaceholderBody        'EACH TEXT PLACEHOLDER
                
            Case Is = ppPlaceholderPicture        'EACH PICTURE PLACEHOLDER
                
            Case Is = ppPlaceholderChart        'EACH CHART PLACEHOLDER
                
            Case Is = ppPlaceholderTable        'EACH TABLE PLACEHOLDER
                
            Case Is = ppPlaceholderOrgChart        'EACH SMARTART PLACEHOLDER
                
            Case Is = ppPlaceholderMediaClip        'EACH MEDIA CLIP PLACEHOLDER
              
            Case Else        'EACH SHAPE THAT IS A PLACEHOLDER BUT NOT THE ABOVE
                
        End Select
    Else
        'EACH SLIDE NON-PLACEHOLDER SHAPE
        
        Select Case shp.Type
      
        Case Is = msoMedia        'EACH MEDIA OBJECT NON-PLACEHOLDER
           
            If shp.MediaType = ppMediaTypeMovie Then        'EACH VIDEO NON-PLACEHOLDER
              
            End If
        Case Is = msoTable        'EACH TABLE NON-PLACEHOLDER
            
        Case Is = msoPicture        'EACH PICTURE NON-PLACEHOLDER
            
        Case Is = msoAutoShape        'EACH SHAPE NON-PLACEHOLDER
            
        Case Is = msoSmartArt        'EACH SMARTART NON-PLACEHOLDER
            
        Case Is = msoChart        'EACH CHART NON-PLACEHOLDER
            
        Case Is = msoTextBox        'EACH TEXTBOX NON-PLACEHOLDER
          
        Case Is = msoGroup        'EACH GROUP OF SHAPES OR GROUP OF GROUPS
            
        Case Is = -2        'EACH SLIDE ZOOM SHAPE
      
       
    End Select
End If
End Function
Function slideCommentActions(sld As Slide, cmt As Comment)
    'EACH COMMENT OR COMMENT REPLY OF A SLIDE
    
End Function
Function noteShapeActions(sld As Slide, shp As Shape)
    'EACH SHAPE OF SLIDE NOTES
    
    If shp.Type = msoPlaceholder Then
        'EACH NOTES PlACEHOLDER SHAPE
        
        Select Case shp.PlaceholderFormat.Type
            Case Is = ppPlaceholderTitle        'EACH SLIDE IMAGE PLACEHOLDER
                
            Case Is = ppPlaceholderDate        'EACH DATE PLACEHOLDER
                
            Case Is = ppPlaceholderFooter        'EACH FOOTER PLACEHOLDER
                
            Case Is = ppPlaceholderSlideNumber        'EACH SLIDE NUMBER PLACEHOLDER
                
            Case Is = ppPlaceholderBody        'EACH PRESENTER NOTES PLACEHOLDER
              
            Case Is = ppPlaceholderHeader        'EACH HEADER PLACEHOLDER
                
            Case Else        'EACH SHAPE THAT IS A PLACEHOLDER BUT NOT THE ABOVE
                
        End Select
    Else
        'EACH NOTES NON-PLACEHOLDER SHAPE
        
        Select Case shp.Type
            Case Is = msoTable        'EACH TABLE NON-PLACEHOLDER
                
            Case Is = msoPicture        'EACH PICTURE NON-PLACEHOLDER
                
            Case Is = msoAutoShape        'EACH SHAPE NON-PLACEHOLDER
                
            Case Is = msoSmartArt        'EACH SMARTART NON-PLACEHOLDER
                
            Case Is = msoChart        'EACH CHART NON-PLACEHOLDER
                
            Case Is = msoTextBox        'EACH TEXTBOX NON-PLACEHOLDER
                
            Case Is = msoGroup        'EACH GROUP OF SHAPES
                
            Case Else        'EACH SHAPE THAT IS NOT A PLACEHOLDER AND NOT THE ABOVE
                
        End Select
    End If
End Function
