Private Sub Workbook_AddinInstall()
    Dim cmbBar As CommandBar
    Dim cmbControl As CommandBarControl
     
    Set cmbBar = Application.CommandBars("Worksheet Menu Bar")
    Set cmbControl = cmbBar.Controls.Add(Type:=msoControlPopup, temporary:=True) 'adds a menu item to the Menu Bar
    
    With cmbControl
        .Caption = "&Table Aggregation" 'names the menu item
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Add Sheet" 'adds a description to the menu item
            .OnAction = "worksheetAdd" 'runs the specified macro
            .FaceId = 3044 'assigns an icon to the dropdown
        End With
    End With
End Sub

Private Sub Workbook_AddinUninstall()
    On Error Resume Next 'in case the menu item has already been deleted
    Application.CommandBars("Worksheet Menu Bar").Controls("Table Aggregation").Delete 'delete the menu ite
End Sub

Private Sub Workbook_Open()
    Dim cmbBar As CommandBar
    Dim cmbControl As CommandBarControl
     
    Set cmbBar = Application.CommandBars("Worksheet Menu Bar")
    Set cmbControl = cmbBar.Controls.Add(Type:=msoControlPopup, temporary:=True) 'adds a menu item to the Menu Bar
    
    With cmbControl
        .Caption = "&Table Aggregation" 'names the menu item
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Add Sheet" 'adds a description to the menu item
            .OnAction = "worksheetAdd" 'runs the specified macro
            .FaceId = 3044 'assigns an icon to the dropdown
        End With
    End With
End Sub
Sub manualMenuDelete()
    On Error Resume Next 'in case the menu item has already been deleted
    Application.CommandBars("Worksheet Menu Bar").Controls("Table Aggregation").Delete 'delete the menu item
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next 'in case the menu item has already been deleted
    Application.CommandBars("Worksheet Menu Bar").Controls("Table Aggregation").Delete 'delete the menu item
End Sub
