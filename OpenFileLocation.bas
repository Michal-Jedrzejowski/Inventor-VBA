Attribute VB_Name = "Module9"
Sub OpenFileLocation()

    Dim selectedSet As SelectSet
    Set selectedSet = ThisApplication.ActiveDocument.SelectSet
    
    Dim obj As Object
    
    If selectedSet.Count = 1 Then
        If TypeOf selectedSet(1) Is ComponentOccurrence Then
            Set obj = selectedSet(1)
        Else
            Set obj = ThisApplication.CommandManager.Pick(kAssemblyLeafOccurrenceFilter, "Select item")
        End If
    Else
        Set obj = ThisApplication.CommandManager.Pick(kAssemblyLeafOccurrenceFilter, "Select item")
    End If

    Dim ObjPath As String
    ObjPath = obj.ReferencedFileDescriptor.FullFileName
    Shell "C:\Windows\explorer.exe /select," & ObjPath, vbMaximizedFocus
        
End Sub
