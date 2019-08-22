Option Explicit

Sub RemoveHeadAndFoot()
    Dim oSec                                      As Section
    Dim oHead                                     As HeaderFooter
    Dim oFoot                                     As HeaderFooter
    Dim headersRemovedCount                       As Integer
    headersRemovedCount = 0
    Dim footersRemovedCount                       As Integer
    footersRemovedCount = 0
    Dim removeHeaders                             As Boolean
    removeHeaders = False
    Dim removeFooters                             As Boolean
    removeFooters = False
    
    If MsgBox("Do you want To remove all headers in all sections?", (vbYesNo + vbQuestion), "Go To max") = vbYes Then
        removeHeaders = True
    End If
    
    If MsgBox("Do you want To remove all footers in all sections?", (vbYesNo + vbQuestion), "Go To max") = vbYes Then
        removeFooters = True
    End If
    
    For Each oSec In ActiveDocument.Sections
        
        If removeHeaders = True Then
            For Each oHead In oSec.Headers
                If oHead.Exists Then oHead.Range.Delete
                headersRemovedCount = headersRemovedCount + 1
            Next oHead
        End If
        
        If removeFooters = True Then
            For Each oFoot In oSec.Footers
                If oFoot.Exists Then oFoot.Range.Delete
                footersRemovedCount = footersRemovedCount + 1
            Next oFoot
            
        End If
        
    Next oSec
    
    MsgBox headersRemovedCount & " headers have been removed Or reset"
    MsgBox footersRemovedCount & " footers removed have been removed Or reset"
    
End Sub
