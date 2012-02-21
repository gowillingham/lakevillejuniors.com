<%
Class cLeagueSession
	Private m_LeagueSessionID		'as small int
	Private m_Name		'as string
	Private m_Description		'as string
	Private m_DisplayOrder		'as small int
	Private m_Price		'as string

	Private m_sSQL		'as string
	Private m_cnn		'as ADODB.Connection
	Private m_rs		'as ADODB.Recordset
	
	Private CLASS_NAME	'as string
	
	Public Property Get LeagueSessionID() 'As small int
		LeagueSessionID = m_LeagueSessionID
	End Property

	Public Property Let LeagueSessionID(val) 'As small int
		m_LeagueSessionID = val
	End Property
	
	Public Property Get Name() 'As string
		Name = m_Name
	End Property

	Public Property Let Name(val) 'As string
		m_Name = val
	End Property
	
	Public Property Get Description() 'As string
		Description = m_Description
	End Property

	Public Property Let Description(val) 'As string
		m_Description = val
	End Property
	
	Public Property Get DisplayOrder() 'As small int
		DisplayOrder = m_DisplayOrder
	End Property

	Public Property Let DisplayOrder(val) 'As small int
		m_DisplayOrder = val
	End Property
	
	Public Property Get Price() 'As string
		Price = m_Price
	End Property

	Public Property Let Price(val) 'As string
		m_Price = val
	End Property
	
	
	Private Sub Class_Initialize()
		m_sSQL = Application.Value("CNN_STR")
		CLASS_NAME = "cLeagueSession"
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
	
	Public Function List() ' as array

	End Function
	
	Public Sub Load() 'As Boolean
		If Len(m_LeagueSessionID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Load():", "Required parameter LeagueSessionID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_GetLeagueSession CInt(m_LeagueSessionID), m_rs
		If Not m_rs.EOF Then
			m_LeagueSessionID = m_rs("LeagueSessionID").Value
			m_Name = m_rs("Name").Value
			m_Description = m_rs("Description").Value
			m_DisplayOrder = m_rs("DisplayOrder").Value
			m_Price = m_rs("Price").Value
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
			.CommandText = "dbo.up_InsertLeagueSession"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 200, m_Name)
		cmd.Parameters.Append cmd.CreateParameter("@Description", adVarChar, adParamInput, 2000, m_Description)
		cmd.Parameters.Append cmd.CreateParameter("@DisplayOrder", adUnsignedTinyInt, adParamInput, 1, m_DisplayOrder)
		cmd.Parameters.Append cmd.CreateParameter("@Price", adCurrency, adParamInput, 8, m_Price)
		cmd.Parameters.Append cmd.CreateParameter("@NewLeagueSessionID", adBigInt, adParamOutput)
	
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		m_LeagueSessionID = cmd.Parameters("@NewLeagueSessionID").Value

		Set cmd = Nothing
	End Sub
	
	Public Sub Save(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(m_LeagueSessionID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Save():", "Required parameter LeagueSessionID not provided.")
		m_DateModified = Now()
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_UpdateLeagueSession"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@LeagueSessionID", adUnsignedTinyInt, adParamInput, 1, m_LeagueSessionID)
		cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 200, m_Name)
		cmd.Parameters.Append cmd.CreateParameter("@Description", adVarChar, adParamInput, 2000, m_Description)
		cmd.Parameters.Append cmd.CreateParameter("@DisplayOrder", adUnsignedTinyInt, adParamInput, 1, m_DisplayOrder)
		cmd.Parameters.Append cmd.CreateParameter("@Price", adCurrency, adParamInput, 8, m_Price)
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		
		Set cmd = Nothing
	End Sub
	
	Public Sub Delete(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(m_LeagueSessionID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Delete():", "Required parameter LeagueSessionID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_DeleteLeagueSession"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@LeagueSessionID", adUnsignedTinyInt, adParamInput, 1, m_LeagueSessionID)

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value

		Set cmd = Nothing
	End Sub
End Class
%>

