<%

	Class cRegistration

		Private m_RegistrationID		'as guid
		Private m_RegistrationNumber	'bigint
		Private m_NameFirstPlayer		'as string
		Private m_NameLastPlayer		'as string
		Private m_NameFirstParent1		'as string
		Private m_NameLastParent1		'as string
		Private m_AddressLine1		'as string
		Private m_AddressLine2		'as string
		Private m_City		'as string
		Private m_StateID		'as string
		Private m_Zip		'as string
		Private m_Phone		'as string
		Private m_PhoneRaw	'as string
		Private m_Email		'as string
		Private m_EmailRetype 'as string
		Private m_School		'as string
		Private m_TShirtSize		'as string
		Private m_Grade		'as small int
		Private m_Session	' as small int
		Private m_IsParentHelper		'as small int
		Private m_Notes		'as string
		Private m_DateCreated		'as date
		Private m_DateModified		'as date
		Private m_IsOnlineRegistration		'as small int
		Private m_RegisFee		'as string
		Private m_RegisPaid		'as string
		Private m_CoachTShirtSize		'as string
		Private m_Team		'as string
		Private m_IsHeadCoach		'as small int
		Private m_IsPaymentConfirmed		'as small int
		Private m_PayPalTransactionID		' as str
		Private m_PayPalIsSandbox	' as small int
		Private m_PayPalPaymentStatus		' as str
		Private m_PayPalPaymentStatusReason		' as str		
		Private m_HasRelease		' as tinyint		
		
		Private m_sSQL		'as string
		Private m_cnn		'as ADODB.Connection
		Private m_rs		'as ADODB.Recordset
		
		Private CLASS_NAME	'as string
		
		Public Property Get RegistrationID() 'As int
			RegistrationID = Replace(Replace(m_RegistrationID, "}", ""), "{", "")
		End Property

		Public Property Let RegistrationID(val) 'As int
			m_RegistrationID = val
		End Property
		
		Public Property Get RegistrationNumber()
			RegistrationNumber = m_RegistrationNumber
		End Property
		
		Public Property Get NameFirstPlayer() 'As string
			NameFirstPlayer = m_NameFirstPlayer
		End Property

		Public Property Let NameFirstPlayer(val) 'As string
			m_NameFirstPlayer = val
		End Property
		
		Public Property Get NameLastPlayer() 'As string
			NameLastPlayer = m_NameLastPlayer
		End Property

		Public Property Let NameLastPlayer(val) 'As string
			m_NameLastPlayer = val
		End Property
		
		Public Property Get NameFirstParent1() 'As string
			NameFirstParent1 = m_NameFirstParent1
		End Property

		Public Property Let NameFirstParent1(val) 'As string
			m_NameFirstParent1 = val
		End Property
		
		Public Property Get NameLastParent1() 'As string
			NameLastParent1 = m_NameLastParent1
		End Property

		Public Property Let NameLastParent1(val) 'As string
			m_NameLastParent1 = val
		End Property
		
		Public Property Get AddressLine1() 'As string
			AddressLine1 = m_AddressLine1
		End Property

		Public Property Let AddressLine1(val) 'As string
			m_AddressLine1 = val
		End Property
		
		Public Property Get AddressLine2() 'As string
			AddressLine2 = m_AddressLine2
		End Property

		Public Property Let AddressLine2(val) 'As string
			m_AddressLine2 = val
		End Property
		
		Public Property Get City() 'As string
			City = m_City
		End Property

		Public Property Let City(val) 'As string
			m_City = val
		End Property
		
		Public Property Get StateID() 'As string
			StateID = m_StateID
		End Property

		Public Property Let StateID(val) 'As string
			m_StateID = val
		End Property
		
		Public Property Get Zip() 'As string
			Zip = m_Zip
		End Property

		Public Property Let Zip(val) 'As string
			m_Zip = val
		End Property
		
		Public Property Get Phone() 'As string
			Phone = m_Phone
		End Property

		Public Property Let Phone(val) 'As string
			m_Phone = val
		End Property
		
		Public Property Get PhoneRaw()
			PhoneRaw = m_PhoneRaw
		End Property
		
		Public Property Get Email() 'As string
			Email = m_Email
		End Property

		Public Property Let Email(val) 'As string
			m_Email = val
		End Property
		
		Public Property Get EmailRetype() 'As string
			EmailRetype = m_EmailRetype
		End Property

		Public Property Let EmailRetype(val) 'As string
			m_EmailRetype = val
		End Property
		
		Public Property Get School() 'As string
			School = m_School
		End Property

		Public Property Let School(val) 'As string
			m_School = val
		End Property
		
		Public Property Get TShirtSize() 'As string
			TShirtSize = m_TShirtSize
		End Property

		Public Property Let TShirtSize(val) 'As string
			m_TShirtSize = val
		End Property
		
		Public Property Get Grade() 'As small int
			Grade = m_Grade
		End Property
		
		Public Property Let Grade(val) 'As small int
			m_Grade = val
		End Property
		
		Public Property Get Session() 
			Session = m_Session
		End Property

		Public Property Let Session(val)
			m_Session = val
		End Property
		
		Public Property Get IsParentHelper() 'As small int
			IsParentHelper = m_IsParentHelper
		End Property

		Public Property Let IsParentHelper(val) 'As small int
			m_IsParentHelper = val
		End Property
		
		Public Property Get Notes() 'As string
			Notes = m_Notes
		End Property

		Public Property Let Notes(val) 'As string
			m_Notes = val
		End Property
		
		Public Property Get DateCreated() 'As date
			DateCreated = m_DateCreated
		End Property

		Public Property Get IsOnlineRegistration() 'As small int
			IsOnlineRegistration = m_IsOnlineRegistration
		End Property

		Public Property Let IsOnlineRegistration(val) 'As small int
			m_IsOnlineRegistration = val
		End Property
		
		Public Property Get RegisFee() 'As string
			RegisFee = m_RegisFee
		End Property

		Public Property Get RegisPaid() 'As string
			RegisPaid = m_RegisPaid
		End Property

		Public Property Let RegisPaid(val) 'As string
			m_RegisPaid = val
		End Property
		
		Public Property Get CoachTShirtSize() 'As string
			CoachTShirtSize = m_CoachTShirtSize
		End Property

		Public Property Let CoachTShirtSize(val) 'As string
			m_CoachTShirtSize = val
		End Property
		
		Public Property Get Team() 'As string
			Team = m_Team
		End Property

		Public Property Let Team(val) 'As string
			m_Team = val
		End Property
		
		Public Property Get IsHeadCoach() 'As small int
			IsHeadCoach = m_IsHeadCoach
		End Property

		Public Property Let IsHeadCoach(val) 'As small int
			m_IsHeadCoach = val
		End Property
		
		Public Property Get IsPaymentConfirmed() 'As small int
			IsPaymentConfirmed = m_IsPaymentConfirmed
		End Property

		Public Property Let IsPaymentConfirmed(val) 'As small int
			m_IsPaymentConfirmed = val
		End Property
		
		Public Property Get PayPalTransactionID() 'As str
			PayPalTransactionID = m_PayPalTransactionID
		End Property

		Public Property Let PayPalTransactionID(val) 'As small int
			m_PayPalTransactionID = val
		End Property
		
		Public Property Get PayPalIsSandbox() 'As str
			PayPalIsSandbox = m_PayPalIsSandbox
		End Property

		Public Property Let PayPalIsSandbox(val) 'As small int
			m_PayPalIsSandbox = val
		End Property
		
		Public Property Get PayPalPaymentStatus() 'As str
			PayPalPaymentStatus = m_PayPalPaymentStatus
		End Property

		Public Property Let PayPalPaymentStatus(val) 'As small int
			m_PayPalPaymentStatus = val
		End Property
		
		Public Property Get PayPalPaymentStatusReason() 'As str
			PayPalPaymentStatusReason = m_PayPalPaymentStatusReason
		End Property

		Public Property Let PayPalPaymentStatusReason(val) 'As small int
			m_PayPalPaymentStatusReason = val
		End Property
		
		Public Property Get HasRelease() 'As str
			HasRelease = m_HasRelease
		End Property

		Public Property Let HasRelease(val) 'As small int
			m_HasRelease = val
		End Property
		
		Private Sub Class_Initialize()
			m_sSQL = Application.Value("CNN_STR")
			CLASS_NAME = "cRegistrations"
		End Sub
		
		Private Sub Class_Terminate()
			If IsObject(m_rs) Then
				If m_rs.State = adStateOpen Then m_rs.Close
				Set m_rs = Nothing
			End If
			If IsObject(m_cnn) Then
				If m_cnn.State = adStateOpen Then m_cnn.Close
				Set m_cnn = Nothing
			End If
		End Sub
		
		Public Function GetRegistrationListByEmail(email)
			Dim cnn			: Set cnn = Server.CreateObject("ADODB.Connection")
			Dim rs			: SEt rs = Server.CreateObject("ADODB.Recordset")
			
			cnn.Open Application.Value("CNN_STR")
			
			' 0-RegistrationID 1-RegistrationNumber 2-NameFirstPlayer 3-NameLastPlayer 4-NameFirstParent
			' 5-NameLastParent 6-AddressLine1 7-AddressLine2 8-City 9-StateID 10-Zip
			' 11-Phone 12-Email 13-school 14-TShirtSize 15-Grade 16-IsParentHelper 17-Notes 18-dateCreated
			' 19-DateModified 20-IsOnlineRegistration 21-RegisFee 22-RegisPaid 23-CoachTShirtSize
			' 24-Team 25-IsHeadCoach 26-IsPaymentConfirmed 27-PayPalTransactionId 28-PayPalIsSandbox
			' 29-PayPalPaymentStatus 30-PayPalPaymentStatusReason 31-HasRelease

			cnn.up_playerGetRegistrationListByEmail CStr(email), rs
			If Not rs.EOF Then GetRegistrationListByEmail = rs.GetRows()
		End Function
		
		Public Function List()
			Dim cnn			: Set cnn = Server.CreateObject("ADODB.Connection")
			Dim rs			: SEt rs = Server.CreateObject("ADODB.Recordset")
			
			cnn.Open Application.Value("CNN_STR")
			
			' 0-RegistrationID 1-RegistrationNumber 2-NameFirstPlayer 3-NameLastPlayer 4-NameFirstParent
			' 5-NameLastParent 6-AddressLine1 7-AddressLine2 8-City 9-StateID 10-Zip
			' 11-Phone 12-Email 13-school 14-TShirtSize 15-Grade 16-IsParentHelper 17-Notes 18-dateCreated
			' 19-DateModified 20-IsOnlineRegistration 21-RegisFee 22-RegisPaid 23-CoachTShirtSize
			' 24-Team 25-IsHeadCoach 26-IsPaymentConfirmed 27-PayPalTransactionId 28-PayPalIsSandbox
			' 29-PayPalPaymentStatus 30-PayPalPaymentStatusReason 31-Session 32-HasRelease

			cnn.up_playerGetRegistrationList rs
			If Not rs.EOF Then List = rs.GetRows()
			
			cnn.Close(): Set cnn = Nothing
		End Function
		
		Public Sub Load() 'As Boolean
			If Len(m_RegistrationID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Load():", "Required parameter RegistrationID not provideded.")
		
			If Not IsObject(m_cnn) Then 
				Set m_cnn = Server.CreateObject("ADODB.Connection")
				m_cnn.Open m_sSQL
			End If
			If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
			
			m_cnn.up_playerGetRegistration m_RegistrationID, m_rs
			If Not m_rs.EOF Then
				m_RegistrationID = m_rs("RegistrationID").Value
				m_RegistrationNumber = m_rs("RegistrationNumber").Value
				m_NameFirstPlayer = m_rs("NameFirstPlayer").Value
				m_NameLastPlayer = m_rs("NameLastPlayer").Value
				m_NameFirstParent1 = m_rs("NameFirstParent1").Value
				m_NameLastParent1 = m_rs("NameLastParent1").Value
				m_AddressLine1 = m_rs("AddressLine1").Value
				m_AddressLine2 = m_rs("AddressLine2").Value
				m_City = m_rs("City").Value
				m_StateID = m_rs("StateID").Value
				m_Zip = m_rs("Zip").Value
				m_Phone = m_rs("Phone").Value
				m_PhoneRaw = m_rs("PhoneRaw").Value
				m_Email = m_rs("Email").Value
				m_EmailRetype = m_rs("Email").Value
				m_School = m_rs("School").Value
				m_TShirtSize = m_rs("TShirtSize").Value
				m_Grade = m_rs("Grade").Value
				m_Session = m_rs("Session").Value
				m_IsParentHelper = m_rs("IsParentHelper").Value
				m_Notes = m_rs("Notes").Value
				m_DateCreated = m_rs("DateCreated").Value
				m_IsOnlineRegistration = m_rs("IsOnlineRegistration").Value
				m_RegisFee = m_rs("RegisFee").Value
				m_RegisPaid = m_rs("RegisPaid").Value
				m_CoachTShirtSize = m_rs("CoachTShirtSize").Value
				m_Team = m_rs("Team").Value
				m_IsHeadCoach = m_rs("IsHeadCoach").Value
				m_IsPaymentConfirmed = m_rs("IsPaymentConfirmed").Value
				m_PayPalTransactionID = m_rs("PayPalTransactionID").Value
				m_PayPalIsSandbox = m_rs("PayPalIsSandbox").Value
				m_PayPalPaymentStatus = m_rs("PayPalPaymentStatus").Value
				m_PayPalPaymentStatusReason = m_rs("PayPalPaymentStatusReason").Value
				m_HasRelease = m_rs("HasRelease").Value
			End If
			
			If m_rs.State = adStateOpen Then m_rs.Close
		End Sub
		
		Public Sub Add(ByRef outError) 'As Boolean
			Dim cmd
			
			m_DateCreated = Now()
			m_DateModified = Now()

			If Not IsObject(m_cnn) Then 
				Set m_cnn = Server.CreateObject("ADODB.Connection")
				m_cnn.Open m_sSQL
			End If
			Set cmd = Server.CreateObject("ADODB.Command")

			With cmd
				.CommandType = adCmdStoredProc
				.CommandText = "dbo.up_playerInsertRegistration"
				.ActiveConnection = m_cnn
			End With
			
			cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
			cmd.Parameters.Append cmd.CreateParameter("@NameFirstPlayer", adVarChar, adParamInput, 50, m_NameFirstPlayer)
			cmd.Parameters.Append cmd.CreateParameter("@NameLastPlayer", adVarChar, adParamInput, 50, m_NameLastPlayer)
			cmd.Parameters.Append cmd.CreateParameter("@NameFirstParent1", adVarChar, adParamInput, 50, m_NameFirstParent1)
			cmd.Parameters.Append cmd.CreateParameter("@NameLastParent1", adVarChar, adParamInput, 50, m_NameLastParent1)
			cmd.Parameters.Append cmd.CreateParameter("@AddressLine1", adVarChar, adParamInput, 100, m_AddressLine1)
			cmd.Parameters.Append cmd.CreateParameter("@AddressLine2", adVarChar, adParamInput, 100, m_AddressLine2)
			cmd.Parameters.Append cmd.CreateParameter("@City", adVarChar, adParamInput, 50, m_City)
			cmd.Parameters.Append cmd.CreateParameter("@StateID", adVarChar, adParamInput, 2, m_StateID)
			cmd.Parameters.Append cmd.CreateParameter("@Zip", adVarChar, adParamInput, 50, m_Zip)
			cmd.Parameters.Append cmd.CreateParameter("@Phone", adVarChar, adParamInput, 50, m_Phone)
			cmd.Parameters.Append cmd.CreateParameter("@Email", adVarChar, adParamInput, 100, m_Email)
			cmd.Parameters.Append cmd.CreateParameter("@School", adVarChar, adParamInput, 3, m_School)
			cmd.Parameters.Append cmd.CreateParameter("@TShirtSize", adVarChar, adParamInput, 5, m_TShirtSize)
			cmd.Parameters.Append cmd.CreateParameter("@Grade", adUnsignedTinyInt, adParamInput, 0, m_Grade)
			cmd.Parameters.Append cmd.CreateParameter("@Session", adUnsignedTinyInt, adParamInput, 0, m_Session)
			cmd.Parameters.Append cmd.CreateParameter("@IsParentHelper", adUnsignedTinyInt, adParamInput, 0, m_IsParentHelper)
			cmd.Parameters.Append cmd.CreateParameter("@Notes", adVarChar, adParamInput, 2000, m_Notes)
			cmd.Parameters.Append cmd.CreateParameter("@DateCreated", adDate, adParamInput, 0, m_DateCreated)
			cmd.Parameters.Append cmd.CreateParameter("@IsOnlineRegistration", adUnsignedTinyInt, adParamInput, 0, m_IsOnlineRegistration)
			cmd.Parameters.Append cmd.CreateParameter("@RegisPaid", adCurrency, adParamInput, 0, m_RegisPaid)
			cmd.Parameters.Append cmd.CreateParameter("@CoachTShirtSize", adVarChar, adParamInput, 5, m_CoachTShirtSize)
			If Len(m_Team) = 0 Then
				cmd.Parameters.Append cmd.CreateParameter("@Team", adVarChar, adParamInput, 25, Null)
			Else
				cmd.Parameters.Append cmd.CreateParameter("@Team", adVarChar, adParamInput, 25, m_Team)
			End If
			If Len(m_IsHeadCoach) = 0 Then
				cmd.Parameters.Append cmd.CreateParameter("@IsHeadCoach", adUnsignedTinyInt, adParamInput, 0, Null)
			Else
				cmd.Parameters.Append cmd.CreateParameter("@IsHeadCoach", adUnsignedTinyInt, adParamInput, 0, m_IsHeadCoach)
			End If
			cmd.Parameters.Append cmd.CreateParameter("@IsPaymentConfirmed", adUnsignedTinyInt, adParamInput, 0, m_IsPaymentConfirmed)
			cmd.Parameters.Append cmd.CreateParameter("@PayPalTransactionID", adVarChar, adParamInput, 256, m_PayPalTransactionID)
			cmd.Parameters.Append cmd.CreateParameter("@PayPalIsSandbox", adUnsignedTinyInt, adParamInput, 0, m_PayPalIsSandbox)
			cmd.Parameters.Append cmd.CreateParameter("@PayPalPaymentStatus", adVarChar, adParamInput, 256, m_PayPalPaymentStatus)
			cmd.Parameters.Append cmd.CreateParameter("@PayPalPaymentStatusReason", adVarChar, adParamInput, 256, m_PayPalPaymentStatusReason)
			cmd.Parameters.Append cmd.CreateParameter("@HasRelease", adUnsignedTinyInt, adParamInput, 0, m_HasRelease)
			cmd.Parameters.Append cmd.CreateParameter("@NewID", adGuid, adParamOutput)
		
			cmd.Execute ,,adExecuteNoRecords
			outError = cmd.Parameters("@RETURN_VALUE").Value
			m_RegistrationID = cmd.Parameters("@NewID").Value

			Set cmd = Nothing
		End Sub
		
		Public Sub Save(ByRef outError) 'As Boolean
			Dim cmd
			
			If Len(m_RegistrationID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Save():", "Required parameter RegistrationID not provideded.")
			m_DateModified = Now()
			
			If Not IsObject(m_cnn) Then 
				Set m_cnn = Server.CreateObject("ADODB.Connection")
				m_cnn.Open m_sSQL
			End If
			Set cmd = Server.CreateObject("ADODB.Command")

			With cmd
				.CommandType = adCmdStoredProc
				.CommandText = "dbo.up_playerUpdateRegistration"
				.ActiveConnection = m_cnn
			End With
			
			cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
			cmd.Parameters.Append cmd.CreateParameter("@RegistrationID", adGuid, adParamInput, 0, m_RegistrationID)
			cmd.Parameters.Append cmd.CreateParameter("@NameFirstPlayer", adVarChar, adParamInput, 50, m_NameFirstPlayer)
			cmd.Parameters.Append cmd.CreateParameter("@NameLastPlayer", adVarChar, adParamInput, 50, m_NameLastPlayer)
			cmd.Parameters.Append cmd.CreateParameter("@NameFirstParent1", adVarChar, adParamInput, 50, m_NameFirstParent1)
			cmd.Parameters.Append cmd.CreateParameter("@NameLastParent1", adVarChar, adParamInput, 50, m_NameLastParent1)
			cmd.Parameters.Append cmd.CreateParameter("@AddressLine1", adVarChar, adParamInput, 100, m_AddressLine1)
			cmd.Parameters.Append cmd.CreateParameter("@AddressLine2", adVarChar, adParamInput, 100, m_AddressLine2)
			cmd.Parameters.Append cmd.CreateParameter("@City", adVarChar, adParamInput, 50, m_City)
			cmd.Parameters.Append cmd.CreateParameter("@StateID", adVarChar, adParamInput, 2, m_StateID)
			cmd.Parameters.Append cmd.CreateParameter("@Zip", adVarChar, adParamInput, 50, m_Zip)
			cmd.Parameters.Append cmd.CreateParameter("@Phone", adVarChar, adParamInput, 50, m_Phone)
			cmd.Parameters.Append cmd.CreateParameter("@Email", adVarChar, adParamInput, 100, m_Email)
			cmd.Parameters.Append cmd.CreateParameter("@School", adVarChar, adParamInput, 3, m_School)
			cmd.Parameters.Append cmd.CreateParameter("@TShirtSize", adVarChar, adParamInput, 5, m_TShirtSize)
			cmd.Parameters.Append cmd.CreateParameter("@Grade", adUnsignedTinyInt, adParamInput, 0, m_Grade)
			cmd.Parameters.Append cmd.CreateParameter("@Session", adUnsignedTinyInt, adParamInput, 0, m_Session)
			cmd.Parameters.Append cmd.CreateParameter("@IsParentHelper", adUnsignedTinyInt, adParamInput, 0, m_IsParentHelper)
			cmd.Parameters.Append cmd.CreateParameter("@Notes", adVarChar, adParamInput, 2000, m_Notes)
			cmd.Parameters.Append cmd.CreateParameter("@DateModified", adDate, adParamInput, 0, m_DateModified)
			cmd.Parameters.Append cmd.CreateParameter("@IsOnlineRegistration", adUnsignedTinyInt, adParamInput, 0, m_IsOnlineRegistration)
			cmd.Parameters.Append cmd.CreateParameter("@RegisPaid", adCurrency, adParamInput, 0, m_RegisPaid)
			cmd.Parameters.Append cmd.CreateParameter("@CoachTShirtSize", adVarChar, adParamInput, 5, m_CoachTShirtSize)
			If Len(m_Team) = 0 Then
				cmd.Parameters.Append cmd.CreateParameter("@Team", adVarChar, adParamInput, 25, Null)
			Else
				cmd.Parameters.Append cmd.CreateParameter("@Team", adVarChar, adParamInput, 25, m_Team)
			End If
			If Len(m_IsHeadCoach) = 0 Then
				cmd.Parameters.Append cmd.CreateParameter("@IsHeadCoach", adUnsignedTinyInt, adParamInput, 0, Null)
			Else
				cmd.Parameters.Append cmd.CreateParameter("@IsHeadCoach", adUnsignedTinyInt, adParamInput, 0, m_IsHeadCoach)
			End If
			cmd.Parameters.Append cmd.CreateParameter("@IsPaymentConfirmed", adUnsignedTinyInt, adParamInput, 0, m_IsPaymentConfirmed)
			cmd.Parameters.Append cmd.CreateParameter("@PayPalTransactionID", adVarChar, adParamInput, 256, m_PayPalTransactionID)
			cmd.Parameters.Append cmd.CreateParameter("@PayPalIsSandbox", adUnsignedTinyInt, adParamInput, 0, m_PayPalIsSandbox)
			cmd.Parameters.Append cmd.CreateParameter("@PayPalPaymentStatus", adVarChar, adParamInput, 256, m_PayPalPaymentStatus)
			cmd.Parameters.Append cmd.CreateParameter("@PayPalPaymentStatusReason", adVarChar, adParamInput, 256, m_PayPalPaymentStatusReason)
			cmd.Parameters.Append cmd.CreateParameter("@HasRelease", adUnsignedTinyInt, adParamInput, 0, m_HasRelease)
			cmd.Execute ,,adExecuteNoRecords
			outError = cmd.Parameters("@RETURN_VALUE").Value
			
			Set cmd = Nothing
		End Sub
		
		Public Sub Delete(ByRef outError) 'As Boolean
			Dim cmd
			
			If Len(m_RegistrationID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Delete():", "Required parameter RegistrationID not provideded.")

			If Not IsObject(m_cnn) Then 
				Set m_cnn = Server.CreateObject("ADODB.Connection")
				m_cnn.Open m_sSQL
			End If
			Set cmd = Server.CreateObject("ADODB.Command")

			With cmd
				.CommandType = adCmdStoredProc
				.CommandText = "dbo.up_playerDeleteRegistration"
				.ActiveConnection = m_cnn
			End With
			
			cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
			cmd.Parameters.Append cmd.CreateParameter("@RegistrationID", adGuid, adParamInput, 0, "{" & m_RegistrationID & "}")

			cmd.Execute ,,adExecuteNoRecords
			outError = cmd.Parameters("@RETURN_VALUE").Value

			Set cmd = Nothing
		End Sub
	End Class

%>