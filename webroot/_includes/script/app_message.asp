<%

Function ApplicationMessageToString(level, header, text)
	Dim str
	
	Dim messageClass
	
	Select Case level
		Case "confirm"
			messageClass = "confirm"
		Case "warning"
			messageClass = "warning"
		Case "error"
			messageClass = "error"
		Case Else
		
	End Select
	
	str = str & "<div class=""application-message"">"
	str = str & "<p class=""" & messageClass & """><span>" & header & "</span> " & text & "</p>"
	str = str & "</div>"
	
	ApplicationMessageToString = str
End Function

Sub ShowAppMsg(iMessageID)
	Dim arrID, arrMsg, msg, sMsg, sPath, sFirstLine, sFinalMsg, i
	'some styles for the message
	Dim sStyleMsgFirstLine, sStyleMsgBody, sStyleMsg
	msg = ""
	sFinalMsg = ""
	
	'check for no MessageID or illegal data
	If Len(CStr(iMessageID)) = 0 Then Exit Sub
	If Not IsNumeric(iMessageID) Then Exit Sub
	
	'account for multiple messages
	arrID = Split(iMessageID, ",")
	If IsArray(arrID) Then
		For i = 0 To UBound(arrID)
			arrMsg = GetMessageByMessageID(arrID(i))
			If IsArray(arrMsg) Then
				'1-Message Text, 2-Message Importance, 3-Path to graphic file
				'don't provide first parm - sLine1 as GetCustomAppMsg will fill it in if missing
				'provide last parm - sPath as this will keep GetCustomAppMsg from hitting db again
				sFinalMsg = sFinalMsg & GetCustomAppMsg("",arrMsg(1,0), arrMsg(2,0), arrMsg(3,0))
			End If
		Next
	End If
	Response.Write sFinalMsg
End Sub

Function GetCustomAppMsg(ByVal sLine1, sLine2, sImportance, sPath)
	Dim cnn, rs, msg, sStyleMsgFirstLine, sStyleMsgBody, sStyleMsg
	
	If Len(sLine1 & sLine2) = 0 Then
		msg = ""
		Exit Function
	End If
	
	If Len(sPath) = 0 Then 
		'get the path to the graphic from the db only if it wasn't passed
		Set cnn = Server.CreateObject("ADODB.Connection")
		cnn.Open Application.Value("CNN_STR")
		Set rs = Server.CreateObject("ADODB.Recordset")
		cnn.up_adminGetMessageGraphicByImportance sImportance, rs
		If Not rs.EOF Then
			sPath = rs("PathToGraphic").Value	
		End IF
		rs.Close: Set rs = Nothing
		cnn.Close: Set cnn = Nothing
	End If	
	
	Select Case sImportance
		'set some text depending on importance
		Case "Error"
			If sLine1 = "" Then sLine1 = "Sorry, an error has occurred. "
			sStyleMsgFirstLine = "errFirstLine"
			sStyleMsgBody = "errText"
			sStyleMsg = "errMessage"
		Case "Critical Error"
			If sLine1 = "" Then sLine1 = "Sorry, a critical error has occurred! "
			sStyleMsgFirstLine = "errFirstLine"
			sStyleMsgBody = "errText"
			sStyleMsg = "errMessage"
		Case "Confirm"
			If sLine1 = "" Then sLine1 = "Thank you! "
			sStyleMsgFirstLine = "msgFirstLine"
			sStyleMsgBody = "msgText"
			sStyleMsg = "msgMessage"
		Case "Info"
			If sLine1 = "" Then sLine1 = "OK! "
			sStyleMsgFirstLine = "msgFirstLine"
			sStyleMsgBody = "msgText"
			sStyleMsg = "msgMessage"
		Case Else
			If sLine1 = "" Then sLine1 = ""	
			sStyleMsgFirstLine = "msgFirstLine"
			sStyleMsgBody = "msgText"
			sStyleMsg = "msgMessage"
	End Select

	'this code writes error message to page as small table
	msg = "<table class=""" & sStyleMsg & """>"	& vbCrLf
	msg = msg & "<tr>"	& vbCrLf
	msg = msg & "<td class=""msgGraphic"">"	& vbCrLf
	msg = msg & "<img src=""" & sPath & """ alt=""" & sImportance & """/>"	& vbCrLf
	msg = msg & "</td>"	& vbCrLf
	msg = msg & "<td>"	& vbCrLf
	msg = msg & "<span class=""" & sStyleMsgFirstLine & """>" & vbCrLf
	msg = msg & sLine1 & "</span>"
	msg = msg & "<span class=""" & sStyleMsgBody & """>" & vbCrLf
	msg = msg & sLine2 & "</span>" & vbCrLf
	msg = msg & "</td>"	& vbCrLf
	msg = msg & "</tr>"	& vbCrLf
	msg = msg & "</table>" & vbCrLf
	
	GetCustomAppMsg = msg
End Function

Function GetRawMessage(iMessageID)
	'returns unformatted string msg from db
	Dim arr
	
	GetRawMessage = ""
	arr = GetMessageByMessageID(iMessageID)
	If IsArray(arr) Then
		GetRawMessage = arr(0,1)
	End If
End Function

Function GetMessageByMessageID(iMessageID)
	'return multi-dimensional, one item (row) array of message information
	Dim cnn, rs
	
	GetMessageByMessageID = ""
	If Len(iMessageID) = 0 Then Exit Function
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.Open Application.Value("CNN_STR")
	Set rs = Server.CreateObject("ADODB.Recordset")
	cnn.up_adminGetMessageByMessageID CInt(iMessageID), rs
	If Not rs.EOF Then GetMessageByMessageID = rs.GetRows
	
	rs.Close: Set rs = Nothing
	cnn.Close: Set cnn = Nothing
End Function

%>