'https://www.slipstick.com/outlook/rules/add-category-accepted-meetings/
'automatically add categories to accepted meetings
Sub AcceptedMeetings(oRequest As MeetingItem)
If oRequest.MessageClass <> "IPM.Schedule.Meeting.Resp.Pos" Then
  Exit Sub
End If
 
Dim oAppt As AppointmentItem
Set oAppt = oRequest.GetAssociatedAppointment(True)
oAppt.Categories = "Green"
oAppt.Save

End Sub