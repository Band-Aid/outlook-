Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    Dim recipents As Outlook.Recipients
    Dim recipent As Outlook.Recipient
    Dim pa As Outlook.PropertyAccessor
    Dim prompt As String
    Dim strMsg As String
	
	

    Const PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"

    Set recipents = Item.Recipients
    For Each recipent In recipents
        Set pa = recipent.PropertyAccessor
        If InStr(LCase(pa.GetProperty(PR_SMTP_ADDRESS)), "@ドメインネームをここに指定") = 0 Then
            strMsg = strMsg & "   " & pa.GetProperty(PR_SMTP_ADDRESS) & vbNewLine
        End If
   Next

    If strMsg <> "" Then
        prompt = "管理ドメイン 以外のユーザが含まれています。to:" & vbNewLine & strMsg & "本当に送りますか?"
        If MsgBox(prompt, vbYesNo + vbExclamation + vbMsgBoxSetForeground, "Check Address") = vbNo Then
            Cancel = True
        End If
    End If
End Sub
