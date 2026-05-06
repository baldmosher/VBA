'--- CC-SUMMARY-START ---
'===========================================================================
' ThisOutlookSession: Application-level Outlook automation entry point
'
' Wires Application_Reminder and Application_NewMailEx events to the report
' trigger and archive routines. This is the top-level entry point for all
' automated Outlook activity — nothing runs without this class being active.
' Three variants exist for different virtual machines (see filename suffix).
' Requires: olTriggers, olReminders, modEmail
'===========================================================================
'--- CC-SUMMARY-END ---
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

Private Const cVMWinVer As String = "VM7"
Private Const cPCIVMemail As String = "pcinsightvm@company.com"
Private Const cTeamEmail As String = "P&CInsight@company.com"

Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
'auto-correction for VM reply-to address
Dim Email_ReplyTo As String
With Item
    If Email_ReplyTo = "" Or Email_ReplyTo = cPCIVMemail Then Email_ReplyTo = cTeamEmail
    While .ReplyRecipients.Count > 0
        .ReplyRecipients.Remove 1
    Wend
    .ReplyRecipients.Add Email_ReplyTo
    .ReplyRecipients.ResolveAll
    '.Display
End With
End Sub

Private Sub Application_Quit()

modEmail.SendEmail "recipient@company.com; recipient@company.com" _
    , _
    , cPCIVMemail _
    , "WARNING: Outlook is shut down on " & cVMWinVer _
    , "Need to switch it back on again.  If this was done as part of reboot process, wait a couple of minutes and log back on to the VM" _
    , , , , , , False, True

MsgBox "DO NOT LEAVE OUTLOOK CLOSED ON VM FOR LONG!!!", vbCritical, ""

'SendReceiveGroup

'Sleep 3000

'MsgBox "quitting"

End Sub

Public Sub SendReceiveGroup()
     
    Dim mySyncObjects As Outlook.SyncObjects
    Dim syc As Outlook.SyncObject
     
    Set mySyncObjects = Application.GetNamespace("MAPI").SyncObjects
 
        Set syc = mySyncObjects(1)
        syc.Start
         
End Sub

Private Sub Application_Startup()

modEmail.SendEmail "recipient@company.com; recipient@company.com" _
    , _
    , cPCIVMemail _
    , "Outlook is restarted on " & cVMWinVer _
    , "" _
    , , , , , , False, True

End Sub
