'Face ID Icon List
'https://bettersolutions.com/vba/ribbon/face-ids-2003.htm

'Workbook_AddinInstall event instead of the Workbook_Open event


Private Sub Workbook_Open() 
    Dim cmbBar As CommandBar 
    Dim cmbControl As CommandBarControl 
     
    Set cmbBar = Application.CommandBars("Worksheet Menu Bar") 
    Set cmbControl = cmbBar.Controls.Add(Type:=msoControlPopup, temporary:=True) 'adds a menu item to the Menu Bar
    With cmbControl 
        .Caption = "&My Macros" 'names the menu item
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "My Macro No 1" 'adds a description to the menu item
            .OnAction = "RunMyMacro1" 'runs the specified macro
            .FaceId = 1098 'assigns an icon to the dropdown
        End With 
        With .Controls.Add(Type:=msoControlButton) 
            .Caption = "My Macro No 2" 
            .OnAction = "RunMyMacro2" 
            .FaceId = 108 
        End With 
        With .Controls.Add(Type:=msoControlButton) 
            .Caption = "My Macro No 3" 
            .OnAction = "RunMyMacro3" 
            .FaceId = 21 
        End With 
    End With 
End Sub 
 
 
Private Sub Workbook_BeforeClose(Cancel As Boolean) 
    On Error Resume Next 'in case the menu item has already been deleted
    Application.CommandBars("Worksheet Menu Bar").Controls("My Macros").Delete 'delete the menu item
End Sub 
 
