' =====================================================================
' CC the Send On Behalf Of address on every mail
' Version: 2026-01-31_203750840
' RE: https://forums.slipstick.com/threads/102150-automatically-cc-the-send-on-behalf-of-account/
' Must be in ThisOutlookSession because of the ItemSend event hook.
' General Help with Outlook VBA: https://www.slipstick.com/developer/how-to-use-outlooks-vba-editor/
'
Private Sub Application_ItemSend(ByVal Item As Object, ByRef Cancel As Boolean)

    '   If Not a MailItem - done
    '
    If Not TypeOf Item Is Outlook.MailItem Then Exit Sub

    '   Cast the Item object to a MailItem
    '
    Dim Mail As Outlook.MailItem
    Set Mail = Item

    '   If not sending from an SMTP account - done
    '
    If Not Mail.SenderEmailType = "SMTP" Then Exit Sub

    '   If no On Behalf Of Name - done
    '
    If Mail.SentOnBehalfOfName = "" Then Exit Sub

    '   Add the Sender as a CC
    '
    Dim CCRecipient As Outlook.Recipient
    Set CCRecipient = Mail.Recipients.Add(Mail.SenderEmailAddress)
    CCRecipient.Type = Outlook.OlMailRecipientType.olCC

    '   Make sure the Sender address resolves
    '
    CCRecipient.Resolve
    If Not CCRecipient.Resolved Then
        MsgBox "CCRecipient did not Resolve." & vbCrLf & vbCrLf & CCRecipient.Address
        Mail.Recipients.Remove CCRecipient.Index
        Cancel = True
        Exit Sub
    End If

End Sub
' =====================================================================
