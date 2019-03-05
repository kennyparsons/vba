Sub vbEmail(varTO, varCC, varBCC, varSubject, varHTMLbody, varAttachment As String, Optional vDisplay As Boolean = False, Optional vSend As Boolean = False)

    Dim signature As String
    Dim OutApp As Object
    Dim OutMail As Object

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    On Error Resume Next
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With
    signature = GetEmailSig
    With OutMail
        .To = varTO
        .CC = varCC
        .BCC = varBCC
        .subject = varSubject
        .HTMLBody = varHTMLbody & signature
        If varAttachment = "" Then
            'No attachemnt, do nothing
        Else
            .Attachments.Add varAttachment
        End If
        'Application.Wait (Now + TimeValue("0:00:03"))
        If vDisplay = True Then .Display
        If vSend = True Then .Send
    End With
    
    On Error GoTo 0
    
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With
    
    Set OutMail = Nothing
    Set OutApp = Nothing

End Sub
