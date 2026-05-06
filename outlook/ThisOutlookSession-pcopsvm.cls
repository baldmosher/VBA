VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

'https://www.slipstick.com/developer/code-samples/running-outlook-macros-schedule/
' The Private subs go in ThisOutlookSession
Private WithEvents olRemind As Outlook.Reminders
Attribute olRemind.VB_VarHelpID = -1

Private Sub Application_Quit()

Dim objNS As Outlook.NameSpace: Set objNS = GetNamespace("MAPI")
Dim olFolder As Outlook.MAPIFolder
Set olFolder = objNS.GetDefaultFolder(olFolderInbox)
Dim sInbox As String
sInbox = olFolder.Parent

modEmail.SendEmail "recipient@company.com; recipient@company.com" _
    , _
    , cPCIVMemail _
    , "WARNING: Outlook is shut down on " & Environ("Computername") & " " & sInbox _
    , "Need to switch it back on again.  If this was done as part of reboot process, wait a couple of minutes and log back on to the VM" _
    , , , , , , False, True

'MsgBox "DO NOT LEAVE OUTLOOK CLOSED ON VM FOR LONG!!!", vbCritical, ""

'SendReceiveGroup

'Sleep 3000

'MsgBox "quitting"

End Sub

Private Sub Application_Startup()

Dim objNS As Outlook.NameSpace: Set objNS = GetNamespace("MAPI")
Dim olFolder As Outlook.MAPIFolder
Set olFolder = objNS.GetDefaultFolder(olFolderInbox)
Dim sInbox As String
sInbox = olFolder.Parent

modEmail.SendEmail "recipient@company.com; recipient@company.com" _
    , _
    , cPCIVMemail _
    , "Outlook restarted on " & Environ("Computername") & " " & sInbox _
    , "" _
    , , , , , , False, True

olReminders.Relaunch_Appointments

End Sub

Private Sub Application_Reminder(ByVal Item As Object)

On Error Resume Next

Set olRemind = Outlook.Reminders

If Item.MessageClass <> "IPM.Appointment" Then
  Exit Sub
End If
 
If Item.Categories <> sAppointmentCategory Then
  Exit Sub
End If
 
If Item.Subject = sAppointmentArchiveSubject Then
'run your code
    Call Audit_All_MailItems
'Delete Appt from calendar when finished
    Item.Delete
' Create another appt to repeat the process
    Call CreateAppointment(sAppointmentArchiveSubject)

'ElseIf Item.Subject = sAppointmentInboxesSubject Then
''run your code
'    Call olAuditCops.Audit_SQL_Inbox_Total
''Delete Appt from calendar when finished
'    Item.Delete
'' Create another appt to repeat the process
'    'Call CreateAppointment(sAppointmentInboxesSubject)

End If

End Sub

' dismiss reminder
Private Sub olRemind_BeforeReminderShow(Cancel As Boolean)
Dim objRem As Outlook.Reminder, sAppointmentSubject As String
On Error Resume Next
    For Each objRem In olRemind
        If objRem.Caption = sAppointmentSubject Then
            If objRem.IsVisible Then
                objRem.Dismiss
                Cancel = True
            End If
            Exit For
        End If
    Next objRem
End Sub





