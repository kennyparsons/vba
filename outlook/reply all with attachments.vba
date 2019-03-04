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
