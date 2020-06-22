
	


 
  
Excel
Add custom menu items to the Menu Bar

Ease of Use
Intermediate
Version tested with
2000, 2002 
Submitted by:
Glaswegian
Description:
Add custom menu items to the Menu Bar. You can then assign specific macros to run from these items 
Discussion:
It can be useful to add custom menu items to a variety of workbooks. These items can be added to a specific workbook, to an add-in or to your Personal.xls. This allows you to run specific macros direct from the Menu Bar - something that is easy even for inexperienced users. By adding the code to your Personal.xls you can assign your favourite or most commonly used macros - ready to run when you need them. This can also be useful when creating an add-in and allows you to create client specific menu items - which increases the professional look of your work. There are two pieces of code involved - the first creates the menu items on opening and the second deletes them on closing the workbook. 
Code:
instructions for use
			
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
 
