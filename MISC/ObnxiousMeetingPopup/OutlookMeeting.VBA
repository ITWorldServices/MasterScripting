'Open VBA in Outlook Developer, then copy and paste this to ThisOutlookSession
Private Sub Application_Reminder(ByVal Item As Object)
    ' This will trigger when any reminder (like for a meeting) goes off
    If Item.MessageClass = "IPM.Appointment" Then
        Call ShowFullScreenNotification(Item.Subject)
    End If
End Sub

Sub ShowFullScreenNotification(meetingName As String)
    ' Prepare the command to call the Python script with the meeting name
    Dim command As String
    meetingName = Replace(meetingName, """", """""") ' Escape double quotes
    command = "pythonw ""C:\temp\meetingalert.py"" """ & meetingName & """"
    
    ' Call the Python script
    Shell command, vbNormalFocus
End Sub
