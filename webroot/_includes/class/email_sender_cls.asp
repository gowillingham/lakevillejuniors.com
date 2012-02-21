<%

' new email class ..
Class cEmail
	Public m_Subject
	Public m_Body
	Public m_ToAddress
	Public m_FromAddress
	
	Private m_Email		' email object
	
	Public Sub Send()
		m_Email.TextBody = m_Body
		m_Email.Subject = m_Subject
		m_Email.To = m_ToAddress
		m_Email.From = m_FromAddress
		Call m_Email.Send()
	End Sub
	
	Public Sub QuickSend(toAddress, fromAddress, subject, body)
		m_Email.TextBody = body
		m_Email.Subject = subject
		m_Email.To = toAddress
		m_Email.From = fromAddress
		Call m_Email.Send()
	End Sub
	
	Private Sub Class_Initialize()
		If Not IsObject(m_Email) Then 
			Set m_Email = Server.CreateObject("CDO.Message")
		End If
	
		If Application.Value("IsLiveSite") Then
			' config info for remote server
			m_Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 ' send using remote smtp
			m_Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") ="mail.lakevillejuniors.com"
			m_Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
			m_Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False 'Use SSL for the connection (True or False)
			m_Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
			' remote server authentication credentials
			m_Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic (clear-text) authentication
			m_Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") ="webapplication@lakevillejuniors"
			m_Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") ="reneepaul"
		Else
			m_Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 1 ' send using pickup folder
			m_Email.Configuration.Fields.Item(cdoSMTPServerPickupDirectory) = "c:\inetpub\mailroot\pickup"
		End If
		
		m_Email.Configuration.Fields.Update
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(m_Email) Then Set m_Email = Nothing
	End Sub
End Class


' deprecating this class ..
Class cEmailSender
	Private m_oEmail, m_oConfig
	Private m_iMsgCount, m_iSentCount, m_iErrCount, m_sErrMsg, m_sAddressWithErrList
	Private m_sPickupFolderPath
	
	Public Function SendMessage(sTo, sFrom, sSubject, sBody)
		Dim i, arr
		SendMessage = 0
		With m_oEmail
			.To = sTo
			.From = sFrom
			.Subject = sSubject
			.TextBody = sBody
		End With

'		On Error Resume Next
			m_oEmail.Send
			If Err.Number <> 0 Then 
				SendMessage = Err.Number
				m_sErrMsg = Err.Description
				'increment error counter
				m_iErrCount = m_iErrCount + 1
				'add this address to list of addresses with errors
				If Len(m_sAddressWithErr) = 0 Then
					m_sAddressWithErrList = sTo
				Else
					m_sAddressWithErrList = m_sAddressWithErrList & " @@ " & sTo
				End If
			Else
				'increment successful send counter
				m_iSentCount = m_iSentCount + 1
			End If
			'increment msg counter
			m_iMsgCount = m_iMsgCount + 1
		On Error GoTo 0
	End Function
	
	Public Function AddAttachment(val)
		m_oEmail.AddAttachment val
	End Function
	
	Public Property Get ErrorCount()
		ErrorCount = m_iErrCount
	End Property
	
	Public Property Get TotalMessageCount()
		TotalMessageCount = m_iMsgCount
	End Property
	
	Public Property Get SentMessageCount()
		SentMessageCount = m_iSentCount
	End Property
	
	Public Property Get ErrorDescription()
		'this will only return the last error if more than 
		'one message is being sent by the object
		ErrorDescription = m_sErrMsg
		m_sErrMsg = 0 
	End Property
	
	Public Property Get AddressWithErrorList()
		'return list of addresses that threw error as array
		AddressWithErrorlist = Split(m_sAddressWithErrList, " @@ ")
	End Property
	
	Private Sub Class_Initialize()
		m_iMsgCount = 0
		m_iSentCount = 0
		m_iErrCount = 0
	
		Set m_oEmail = Server.CreateObject("CDO.Message")
		Set m_oConfig = m_oEmail.Configuration
		
		With m_oConfig
			'pickup folder
			m_sPickupFolderPath = Application.Value("cEmailSender.PICKUP_FOLDER")
			If Len(m_sPickupFolderPath) > 0 Then 
				.Fields.Item(cdoSMTPServerPickupDirectory) = m_sPickupFolderPath
			Else
				.Fields.Item(cdoSMTPServerPickupDirectory) = "c:\inetpub\mailroot\pickup"
			End If
			
			'send using method
			.Fields.Item(cdoSendUsingMethod) = 1 'pickup
			
			.Fields.Update
		End With
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(m_oConfig) Then Set m_oConfig = Nothing
		If IsObject(m_oEmail) Then Set m_oEmail = Nothing
	End Sub
End Class
%>