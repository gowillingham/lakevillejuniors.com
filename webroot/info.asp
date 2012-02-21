<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>

<%Call Main()%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
	<link href="_includes/styles/default.css" rel="stylesheet" type="text/css" />
	<title>Lakeville Juniors - Team info</title>
</head>
	<body>
		<%=m_bodyText%>
	</body>
</html>

<%

Dim m_bodyText

Sub Main()
	Dim str, i
	
	Dim page					: Set page = New cPage
	
	Call OnPageLoad(page)
	
	Dim registration			: Set registration = New cRegistration
	Dim leagueSession			: Set leagueSession = New cLeagueSession
	Dim list
	
	Dim count					: count = 0
	Dim coachText
			
	registration.RegistrationId = page.RegistrationId
	If Len(registration.RegistrationId) > 0 Then Call registration.Load()
	
	leagueSession.LeagueSessionId = registration.Session
	If Len(leagueSession.LeagueSessionId) > 0 Then Call leagueSession.Load()
	
	list = registration.List()
	
	str = str & "<div class=""summary"">"
	str = str & "<h1>Player info</h1>"
	str = str & "<table class=""listing"">"
	str = str & "<tr><td class=""label"">Player</td><td>" & Server.HtmlEncode(registration.NameLastPlayer & ", " & registration.NameFirstPlayer) & "</td></tr>"
	str = str & "<tr><td class=""label"">Grade</td><td>" & registration.Grade & "</td></tr>"
	str = str & "<tr><td class=""label"">School</td><td>" & registration.School & "</td></tr>"
	str = str & "<tr><td class=""label"">Team</td><td>" & registration.Team & "" & "</td></tr>"
	str = str & "<tr><td class=""label"">Session</td><td>" & leagueSession.Name & "</td></tr>"
	str = str & "</table>"
	
	str = str & "<h1>Parent info</h1>"
	str = str & "<table class=""listing"">"
	str = str & "<tr><td class=""label"">Parent</td><td>" & Server.HtmlEncode(registration.NameLastParent1 & ", " & registration.NameFirstParent1) & "</td></tr>"

	' commented out private info ..
''	str = str & "<tr><td class=""label"">Address</td><td>" & Server.HtmlEncode(registration.AddressLine1)
''	str = str & "<br />" & Server.HtmlEncode(registration.City) & ", " & registration.StateId & " " & registration.Zip & "</td></tr>"
''	str = str & "<tr><td class=""label"">Phone</td><td>" & registration.Phone & "</td></tr>"
''	str = str & "<tr><td class=""label"">Email</td><td>" & Server.HtmlEncode(registration.Email) & "</td></tr>"

	coachText = "No"
	If CStr(registration.IsParentHelper & "") = CStr(1) Then coachText = "Assistant coach"
	If CStr(registration.IsHeadCoach & "") = CStr(1) Then coachText = "Lead coach"
	str = str & "<tr><td class=""label"">Coaching</td><td>" & coachText & "</td></tr>"
	str = str & "</table>"

	' 0-RegistrationID 1-RegistrationNumber 2-NameFirstPlayer 3-NameLastPlayer 4-NameFirstParent
	' 5-NameLastParent 6-AddressLine1 7-AddressLine2 8-City 9-StateID 10-Zip
	' 11-Phone 12-Email 13-school 14-TShirtSize 15-Grade 16-IsParentHelper 17-Notes 18-dateCreated
	' 19-DateModified 20-IsOnlineRegistration 21-RegisFee 22-RegisPaid 23-CoachTShirtSize
	' 24-Team 25-IsHeadCoach 26-IsPaymentConfirmed 27-PayPalTransactionId 28-PayPalIsSandbox
	' 29-PayPalPaymentStatus 30-PayPalPaymentStatusReason 31-Session 32-HasRelease

	If Len(registration.Team & "") > 0 Then
		
		If IsArray(list) Then
			str = str & "<h1>" & registration.Team & " team roster</h1>"
			str = str & "<table class=""listing"">"
			str = str & "<tr><th>&nbsp;</th><th>Player</th><th>School</th><th>Grade</th></tr>"
			For i = 0 To UBound(list,2)
				If (registration.Team = list(24,i)) And (registration.Session = list(31,i)) Then
					count = count + 1
					
					str = str & "<tr><td>" & count & "</td>"
					str = str & "<td>" & Server.HtmlEncode(list(3,i) & ", " & list(2,i)) & "</td>"
					str = str & "<td>" & list(13,i) & "</td>"
					str = str & "<td>" & list(15,i) & "</td></tr>"
				End If
			Next
			str = str & "</table>"
		End If
	End If
	
	
	str = str & "</div>"
	
	m_bodyText = str
End Sub

Sub OnPageLoad(page)
	page.RegistrationId = Request.QueryString("id")
End Sub
%>
<!--#INCLUDE VIRTUAL="/_includes/constants/constants.asp"-->
<!--#INCLUDE VIRTUAL="/_includes/class/registration_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_includes/class/league_session_cls.asp"-->
<%
Class cPage
	Public RegistrationId
End Class
%>
