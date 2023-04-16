Attribute VB_Name = "Module25"
Sub ShowWorkPlanes()

    Dim obj As Object
    Set obj = ThisApplication.CommandManager.Pick(kAssemblyLeafOccurrenceFilter, "Select item")

    Dim oPlanes As WorkPlanes
    Set oPlanes = obj.Definition.WorkPlanes
    
    For i = 1 To 3
        oPlanes(i).Visible = Not oPlanes(i).Visible
    Next
    
    For i = 4 To oPlanes.Count
        oPlanes(i).Visible = False
    Next
    
    ThisApplication.ActiveView.Update
    
End Sub
