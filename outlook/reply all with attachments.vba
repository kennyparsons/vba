Sub ReplyAllWithAttachments()
    Dim rpl As Outlook.MailItem
    Dim itm As Object
     
    Set itm = GetCurrentItem()
    If Not itm Is Nothing Then
        Set rpl = itm.ReplyAll
        CopyAttachments itm, rpl
        rpl.Display
    End If
     
    Set rpl = Nothing
    Set itm = Nothing
End Sub
Sub CopyAttachments(objSourceItem, objTargetItem)
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set fldTemp = fso.GetSpecialFolder(2) ' TemporaryFolder
   strPath = fldTemp.Path & "\"
   For Each objAtt In objSourceItem.Attachments
      strFile = strPath & objAtt.FileName
      objAtt.SaveAsFile strFile
      objTargetItem.Attachments.Add strFile, , , objAtt.DisplayName
      fso.DeleteFile strFile
   Next
 
   Set fldTemp = Nothing
   Set fso = Nothing
End Sub
