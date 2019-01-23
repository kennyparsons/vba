https://stackoverflow.com/questions/25978315/auto-run-when-an-appointment-is-updated
Macro event when appointment is updated

Public WithEvents myOlItems As Outlook.Items 

Public Sub Application_Startup() 
  Set myOlItems = _
    Application.GetNamespace("MAPI").GetDefaultFolder(olFolderCalendar).Items 
End Sub 

Private Sub myOlItems_ItemChange(ByVal Item As Object) 
    debug.print item.subject
End Sub