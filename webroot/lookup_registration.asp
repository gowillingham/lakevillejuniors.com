<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%

Sub OnPageLoad(page)
	page.Action = Request.QueryString("act")
	
	page.LookupEmailAddress = Request.Form("email")
	
	Set page.Registration = New cRegistration
	
End Sub
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
	<link href="_includes/styles/default.css" rel="stylesheet" type="text/css" />
	<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.2.6/jquery.min.js"></script>
	<title>Lakeville Juniors Online Registration</title>
	
	<script type="text/javascript" language="javascript">
		$(document).ready(function(){
			
		});
	</script>
</head>
	<body>
		<!--#INCLUDE VIRTUAL="/_includes/maincontent.asp"-->
	</body>
</html>
<%
Sub Main
	Dim str, rv
	Dim page			: Set page = New cPage
	Dim count
	
	Call OnPageLoad(page)

	Select Case page.Action
	
		Case Else
			If Request.Form("form_lookup_registration_is_postback") = IS_POSTBACK Then
				If ValidFormLookup(page.LookupEmailAddress) Then
					Call DoEmailRegistrationReport(page, count)
					str = str & FormLookupRegistrationToString(page)
					str = str & ConfirmReportToString(page, count)
				Else
				response.Write "<p>here .."
					str = str & FormLookupRegistrationToString(page)
				End If
			Else
				str = str & FormLookupRegistrationToString(page)
			End If

	End Select

	Response.Write str
	Set page = Nothing
End Sub

Function ConfirmReportToString(page, count)
	Dim str
	Dim headerText
	
	Dim countText		: countText = count & " registration"
	If count <> 1 Then countText = countText & "s"		
	
	If count > 0 Then
		headerText = "There are " & countText & " associated with that email address. "
		str = str & "We have sent an email message to <strong>" & server.HTMLEncode(page.LookupEmailAddress) & "</strong> with the details for each registration "
		str = str & "(please check that any spam filters you use are set to allow email sent from @lakevillejuniors.com). "
		str = str & "If you have any other questions about an in-house volleyball registration, you should inquire at <a href=""mailto:info@lakevillejuniors.com"">info@lakevillejuniors.com</a> and we'll try to get back to you right away. "
	
		str = ApplicationMessageToString("confirm", headerText, str)
		
		' clear the email address for the form ..
		page.LookupEmailAddress = ""
	Else
		headerText = "No registrations were found for that address!"
		str = str & "No registrations in our database matched the email address that you provided (" & server.HTMLEncode(page.LookupEmailAddress) & "). "
		str = str & "Perhaps you registered with a different email address or you have mispelled the address. "
		str = str & "If you believe you are registered but nothing is being returned, please inquire at <a href=""mailto:info@lakevillejuniors.com"">info@lakevillejuniors.com</a> and we'll try to get back to you right away. "
		
		str = ApplicationMessageToString("error", headerText, str)
	End If
	
	ConfirmReportToString = str
End Function

Sub DoEmailRegistrationReport(page, count)
	Dim str, i
	Dim email			: Set email = New cEmail
	
	count = 0
	
	Dim team			: team = ""
	Dim subject			: subject = ""
	Dim countText
	Dim list			: list = page.Registration.GetRegistrationListByEmail(page.LookupEmailAddress)
	
	' exit if no records returned ..
	If Not IsArray(list) Then Exit Sub
	
	' set registration count for output ..
	count = UBound(list,2) + 1
	
	countText = count & " registration"
	If count <> 1 Then countText = countText & "s"
	
	' 0-RegistrationID 1-RegistrationNumber 2-NameFirstPlayer 3-NameLastPlayer 4-NameFirstParent
	' 5-NameLastParent 6-AddressLine1 7-AddressLine2 8-City 9-StateID 10-Zip
	' 11-Phone 12-Email 13-school 14-TShirtSize 15-Grade 16-IsParentHelper 17-Notes 18-dateCreated
	' 19-DateModified 20-IsOnlineRegistration 21-RegisFee 22-RegisPaid 23-CoachTShirtSize
	' 24-Team 25-IsHeadCoach 26-IsPaymentConfirmed 27-PayPalTransactionId 28-PayPalIsSandbox
	' 29-PayPalPaymentStatus 30-PayPalPaymentStatusReason 31-SessionName 32-SessionDescription

	' send an email ..
	subject = "[" & Application.Value("COMPANY_NAME") & "] **In-house registration details for " & page.LookupEmailAddress & "**"
	str = str & "Hello"
	str = str & vbCrLf & vbCrLf & "Thank you for your inquiry regarding " & Application.Value("COMPANY_NAME") & " in-house volleyball registration. "
	str = str & "It looks like you have " & countText & " with this email address (" & page.LookupEmailAddress & "). "
	str = str & "Please find below a listing of the players you have registered. "
	
	For i = 0 To UBound(list,2)
		team = "Not yet available ..":			If Len(list(24,i) & "") > 0 Then team = list(24,i)
		
		str = str & vbCrLf & vbCrLf
		str = str & "Player: " & list(3,i) & ", " & list(2,i)
		str = str & vbCrLf & "Parents: " & list(4,i) & " " & list(5,i)
		str = str & vbCrLf & "Grade: " & list(15,i)
		str = str & vbCrLf & "School: " & list(13,i)
		str = str & vbCrLf & "Session: " & list(32,i)
		str = str & vbCrLf & "Balance due: " & FormatCurrency(list(21,i) - list(22,i), 2)
		str = str & vbCrLf & "Team: " & team
		str = str & vbCrLf & "Registration received: " & FormatDateTime(list(18,i), vbLongDate)
		str = str & vbCrLf & "Registration ID: " & Replace(Replace(list(0,i), "}", ""), "{", "")
	Next
	
	str = str & vbCrLf & vbCrLf
	str = str & "Thanks again for your registration. "
	str = str & "If you have an outstanding balance, you may bring payment to the first league session. "
	str = str & "Any other questions about your in-house volleyball registration may be directed to mailto:" & Application.Value("INFO_EMAIL_ADDRESS") & ". "
	str = str & vbCrLf & vbCrLf & "See you at volleyball! "
	str = str & vbCrLf & vbCrLf & "--"
	str = str & vbCrLf & vbCrLf & "Lakeville Junior Volleyball"
	
	Call email.QuickSend(page.LookupEmailAddress, Application.Value("INFO_EMAIL_ADDRESS"), subject, str)
End Sub

Function ValidFormLookup(email)
	Dim rv:			rv = True
	
	If Not ValidData(email, True, 0, 50, "Email", "email") Then rv = False
	
	ValidFormLookup = rv
End Function

Function FormLookupRegistrationToString(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	
	str = str & "<h1>Lookup registrations</h1>"
	str = str & "<p>Wondering if your player is registered for in-house volleyball or not? "
	str = str & "If you provided an email address with your registration, just put that address in the form below "
	str = str & "and we'll shoot you an email with the information for all the players you have registered. </p>"
	
	str = str & FormErrorToString()
	str = str & "<form action=""/lookup_registration.asp" & pg.UrlParamsToString(True) & """ method=""post"" id=""form-lookup-registration"" class=""form"">"
	str = str & "<input type=""hidden"" name=""form_lookup_registration_is_postback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<table><tr><td class=""label"">Your email address</td><td><input type=""text"" name=""email"" value=""" & Server.HTMLEncode(page.LookupEmailAddress) & """ class=""medium"" />"
	str = str & "&nbsp;<input type=""submit"" name=""submit"" value=""Lookup!"" /></td></tr>"
	str = str & "</table></form>"
	
	FormLookupRegistrationToString = str
End Function
%>
<!--#INCLUDE VIRTUAL="/_includes/script/frm_valid_data.asp"-->
<!--#INCLUDE VIRTUAL="/_includes/script/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_includes/constants/constants.asp"-->
<!--#INCLUDE VIRTUAL="/_includes/class/email_sender_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_includes/class/registration_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_includes/class/league_session_cls.asp"-->
<%
Class cPage
	' url params
	Public Action
	Public RegistrationID
	
	' obj
	Public Registration
	
	' not for url
	Public LookupEmailAddress
	
	Public Function UrlParamsToString(isEncoded)
		Dim str, amp
		
		amp = "&"
		If isEncoded Then amp = "&amp;"
		
		If Len(Action) > 0 Then str = str & "act=" & Action & amp
		If Len(RegistrationID) > 0 Then str = str & "rid=" & RegistrationID & amp
		
		If Len(str) > 0 Then 
			str = Left(str, Len(str) - Len(amp))
		Else
			' qstring needs at least one param in case more params are appended ..
			str = str & "noparm=true"
		End If
		str = "?" & str
		
		UrlParamsToString = str
	End Function
	
	Public Function Clone()
		Dim c			: Set c = New cPage
		
		c.Action = Action
		c.RegistrationID = RegistrationID
		Set c.Registration = Registration
		
		Set Clone = c
	End Function
End Class
%>
