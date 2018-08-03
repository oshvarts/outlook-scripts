' This macro creates reminders on incoming meetings if they were organizer did not specify them,
' and clears reminders from all-day events that don't need them.

Private WithEvents Items As Outlook.Items


Private Sub log(s As String)

    Debug.Print s ' write to immediate log

    '' To log to a file, uncomment the section below.
    
    ' Dim n As Integer
    ' n = FreeFile()
    ' Open "C:\d\outlook.log" For Append As #n
    ' Print #n, DateTime.Now
    ' Print #n, s
    ' Close #n
End Sub


Private Sub Application_Startup()
  log ("Application startup ENTRY")
  Dim Ns As Outlook.NameSpace

  Set Ns = Application.GetNamespace("MAPI")
  Set Items = Ns.GetDefaultFolder(olFolderCalendar).Items
  log ("Application startup EXIT")
End Sub

Private Sub Items_ItemAdd(ByVal Item As Object)
  On Error Resume Next
  log ("ItemAdd ENTRY")
  Dim Appt As Outlook.AppointmentItem

  If TypeOf Item Is Outlook.AppointmentItem Then

    Set Appt = Item
    log ("processing meeting: " + Appt.Subject)
    
    'Checks to see if all day and if it has a reminder set to true
    If Appt.AllDayEvent = True And Appt.ReminderSet = True Then

       'msgbox block - 3 lines
       'If MsgBox("Do you want to remove the reminder?", vbYesNo) = vbNo Then
       '  Exit Sub
       'End If

       'appt.reminderset block - 2 lines
        log ("Clearing reminder for all day meeting")
        Appt.ReminderSet = False
        Appt.Save
        log ("Saved")
    ElseIf Appt.AllDayEvent = False And Appt.ReminderSet = True And Appt.Duration = 60 * 24 Then
       ' 24-hour meeting, all-day equivalent meeting in another time zone.
        log ("Clearing reminder for 24-hour meeting")
        log (Appt.Duration + " minutes")
        Appt.ReminderSet = False
        Appt.Save
        log ("Saved")                    
    ElseIf Appt.AllDayEvent = False And Appt.ReminderSet = False Then
       'msgbox block - 3 lines
       'If MsgBox("Do you want to change the reminder to 15 minutes?", vbYesNo) = vbNo Then
       '  Exit Sub
       'End If
       
        log ("Setting reminder on un-remindered meeting")
        Appt.ReminderSet = True
        Appt.ReminderMinutesBeforeStart = 15
        Appt.Save
        log ("Saved")
    End If
  End If
  log ("ItemAdd EXIT")
End Sub
