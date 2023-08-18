Attribute VB_Name = "Module11"
Sub ObjVisibility()

    Dim oDoc As Document
    Set oDoc = ThisApplication.ActiveDocument
    
    Dim selectedSet As SelectSet
    Set selectedSet = oDoc.SelectSet
    
    Dim isAssembly As Boolean
    If oDoc.DocumentType = kAssemblyDocumentObject Then
        isAssembly = True
    End If
    
    
    Dim obj As Object
    For Each obj In selectedSet
        'To pass if "Object doesn't support this property or method" occure (if you have e.g. Feature priority filter)
        On Error Resume Next
        
        'To add new case add breakpoint and check obj.Type property
        Select Case obj.Type
            Case kBrowserFolderObject 'for hide documents in folders (if at least one is visible then objectVisibility will be true
                Dim objectVisibility As Boolean
                objectVisibility = False
                
                Dim objectInFolder As Variant
    
                For Each objectInFolder In obj.BrowserNode.BrowserNodes
                    If CheckIfVisible(objectInFolder.NativeObject) Then
                        objectVisibility = True
                        Exit For
                    End If
                Next
                
                For Each objectInFolder In obj.BrowserNode.BrowserNodes
                    Call ObjectsInFolderVisibility(objectInFolder.NativeObject, objectVisibility)
                Next
                
            Case kFaceProxyObject
                If isAssembly Then
                    obj.Parent.Parent.Visible = Not obj.Parent.Parent.Visible
                Else
                    obj.Parent.Visible = Not obj.Parent.Visible
                End If
                
            Case kFaceObject
                If isAssembly Then
                    obj.Parent.Parent.Visible = Not obj.Parent.Parent.Visible
                Else
                    obj.Parent.Visible = Not obj.Parent.Visible
                End If
                
            Case kSurfaceBodyProxyObject
                If isAssembly Then
                    obj.Parent.Visible = Not obj.Parent.Visible
                Else
                    obj.Visible = Not obj.Visible
                End If
                
            Case Else 'On error
                obj.Visible = Not obj.Visible
        End Select
            
     Next
        
End Sub

Private Sub ObjectsInFolderVisibility(obj As Variant, ov As Boolean)
    
    If ov Then
        obj.Visible = False
    Else
        obj.Visible = True
    End If

End Sub

Private Function CheckIfVisible(obj As Variant) As Boolean

    If obj.Visible Then
        CheckIfVisible = True
    Else
        CheckIfVisible = False
    End If
    
End Function



