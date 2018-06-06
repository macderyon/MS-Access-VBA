# MS-Access-VBA

With little chane to the library setting (tools) of MS Access, this code will authomate As Access form to send out an email to MS 
outlook user.
-------------------------------------------------------------------------------------------------------------------------------------
'To clear the textbox, Email, Subject and Body of the form. 
Private Sub btnClear_Click()
    Me.txtEmail = Null
    Me.txtSubject = Null
    Me.txtBody = Null
End Sub

------------------------------------------------------------------------------------------------------------------------------------
Private Sub btnSend_Click()

    Dim oApp As New Outlook.Application
    Dim oEmail As Outlook.MailItem
    Set oEmail = oApp.CreateItem(olMailItem)
    
    oEmail.To = Me.cboemail
    oEmail.Subject = Me.txtSubject
    oEmail.body = Me.txtBody
  
    With oEmail
        If Not IsNull(.To) And Not IsNull(.Subject) And Not IsNull(.body) Then
            .Send
            MsgBox "Email Sent!"
        Else
            MsgBox "Please fill out all the fields"
        End If
    End With
    
End Sub
--------------------------------------------------------------------------------------------------------------------------------------
'Close button fuction

Private Sub cmdCancel_Click()
    DoCmd.Close
End Sub
---------------------------------------------------------------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
    DoCmd.Restore
    txtBody = ""
    txtBody.SetFocus
End Sub
