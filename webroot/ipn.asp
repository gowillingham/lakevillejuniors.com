<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>

<script runat="server" type="text/vbscript" language="vbscript">

	Call Main()

	Sub Main()
		Dim http			: Set http = Server.CreateObject("Msxml2.ServerXMLHTTP")
		Dim registration	: Set registration = New cRegistration
		Dim myResponse		: myResponse = Request.Form & "&cmd=_notify-validate"
		Dim httpResponse, str, rv

		' repost to paypal for verify
		Call http.Open("POST", Application.Value("PAYPAL_GATEWAY"), False)
		Call http.SetRequestHeader("Content-type", "application/x-www-form-urlencoded")
		Call http.Send(myResponse)
		httpResponse = http.ResponseText
		
		' generate text for logging email
		str = str & "timestamp=" & Now()
		str = str & vbCrLf & vbCrLf & "http.ResponseText=" & http.ResponseText
		str = str & vbCrLf & vbCrLf & "test_ipn=" & Request.Form("test_ipn")
		str = str & vbCrLf & "RegistrationID=" & Request.Form("invoice")
		str = str & vbCrLf & "RegisPaid=" & Request.Form("mc_gross")
		str = str & vbCrLf & "txn_id=" & Request.Form("txn_id")
		str = str & vbCrLf & "payment_status=" & Request.Form("payment_status")
		str = str & vbCrLf & "payment_status_reason=" & Request.Form("payment_status_reason")
		str = str & vbCrLf & "business=" & Request.Form("business")
		str = str & vbCrLf & "mc_gross=" & Request.Form("mc_gross")
		str = str & vbCrLf & "mc_currency=" & Request.Form("mc_currency")
		
		' load the saved/incomplete registration ..
		registration.RegistrationID = Request.Form("invoice")
		If Len(registration.RegistrationID) > 0 Then Call registration.Load()
		
		' test validity of post returned by paypal
		If Not ValidPost(httpResponse, registration.RegisFee, str) Then
			Call SendLogEmail(str)
			Exit Sub
		End If
		
		' generate log email text and save registration to db ..
		If Len(registration.RegistrationID) > 0 Then
		
			str = str & vbCrLf & "registration.RegistrationNumber=" & registration.RegistrationNumber
		
			registration.IsPaymentConfirmed = 1
			registration.RegisPaid = Request.Form("mc_gross")
			registration.PayPalTransactionID = Request.Form("txn_id")
			registration.PayPalPaymentStatus = Request.Form("payment_status")
			registration.PayPalPaymentStatusReason = Request.Form("payment_status_reason")
			If Len(Request.Form("test_ipn")) = 0 Then
				registration.PayPalIsSandbox = 0
			Else
				registration.PayPalIsSandbox = 1
			End If
			
			Call Registration.Save(rv)
			str = str & vbCrLf & "registration.Save(rv)=" & rv
			
			If rv = 0 Then
				Call DoEmailRegistrantConfirmation(registration)
			End If
		End If
		
		Call SendLogEmail(str)

		Set http = Nothing
		Set registration = Nothing
	End Sub	

	Sub SendLogEmail(str)
		Dim email			: Set email = New cEmail
		
		Call email.QuickSend(Application.Value("ADMIN_EMAIL_ADDRESS"), "ipn@lakevillejuniors.com", "[" & Request.ServerVariables("SERVER_NAME") & "] ** IPN Transaction Report **", str)
		Set email = Nothing
	End Sub
		
	Function ValidPost(httpResponse, regisFee, str)
		ValidPost = True
		
		' check for verified from paypal
		If UCase(httpResponse) <> UCase("VERIFIED") Then	
			str = str & vbCrLf & vbCrLf & "Error: Paypal did not return 'VERIFIED'."
			ValidPost = False
			Exit Function
		End If
		
		' check payment status
		If UCase(Request.Form("payment_status")) <> UCase("Completed") Then
			If UCase(Request.Form("payment_status")) = UCase("Pending") Then
				' treat pending transaction as paid ..
				str = str & vbCrLf & vbCrLf & "Error: payment_status = 'Pending'"
			Else
				str = str & vbCrLf & vbCrLf & "Error: payment_status <> 'Completed' or 'Pending'"
				ValidPost = False
			End If
		End If
		
		' spoof - checking for legit email address from my paypal acct
		If UCase(Request.Form("business")) <> UCase(Application.Value("PAYPAL_BUSINESS_ID")) Then
			str = str & vbCrLf & vbCrLf & "Error: Possible spoofed post to /ipn.asp. Business email address does not match a legit email for my paypal acct."
			ValidPost = False
		End If
		
		' spoof - check for price
		If FormatCurrency(Request.Form("mc_gross"), 2) <> FormatCurrency(regisFee, 2) Then
			str = str & vbCrLf & vbCrLf & "Error: Possible spoofed post to /ipn.asp. mc_gross not equal to item price originally passed to PayPal."
			ValidPost = False
		End If
		
		' spoof - check for currency
		If UCase(request.Form("mc_currency")) <> UCase("USD") Then
			str = str & vbCrLf & vbCrLf & "Error: Possible spoofed post to /ipn.asp. mc_currency <> 'USD'."
			ValidPost = False
		End If
	End Function
		
	Sub DoEmailRegistrantConfirmation(registration)
		Dim str 
		Dim email			: Set email = New cEmail
		Dim fromAddress		: fromAddress = Application.Value("INFO_EMAIL_ADDRESS")
		Dim toAddress		: toAddress = registration.Email
		Dim leagueName
		If IsDate(Application.Value("INHOUSE_LEAGUE_FIRST_SESSION_DATE")) Then 
			leagueName = "Lakeville Juniors " & Application.Value("INHOUSE_LEAGUE_VERSION") & " " & Year(Application.Value("INHOUSE_LEAGUE_FIRST_SESSION_DATE")) & " In-House Volleyball League"
		Else
			leagueName = "Lakeville Juniors In-house League for " & Application.Value("INHOUSE_LEAGUE_FIRST_SESSION_DATE")
		End If
		Dim subject			: subject = "** " & leagueName & " Confirmation **"
		
		Dim firstSession	
		If IsDate(Application.Value("INHOUSE_LEAGUE_FIRST_SESSION_DATE")) Then
			firstSession = FormatDateTime(Application.Value("INHOUSE_LEAGUE_FIRST_SESSION_DATE"), vbLongDate)
		Else
			firstSession = Application.Value("INHOUSE_LEAGUE_FIRST_SESSION_DATE")
		End If
		
		str = str & "Hello " & registration.NameFirstParent1 & " " & registration.NameLastParent1
		str = str & vbCrLf & vbCrLf & "Thank your for your registration for the " & leagueName & ". "
		str = str & "This email is to confirm registration for " & registration.NameFirstPlayer & " " & registration.NameLastPlayer & ". "
		str = str & "Please check over the registration details below to make sure that they are correct. "
		str = str & "If there is an error, please reply to this email message with any changes you would wish to make. "
		str = str & vbCrLf & vbCrLf & "Registration Information" 
		str = str & vbCrLf & "--------------------------"
		str = str & vbCrLf & "Regis#: " & registration.RegistrationNumber
		str = str & vbCrLf & registration.NameFirstPlayer & " " & registration.NameLastPlayer
		str = str & vbCrLf & registration.AddressLine1
		If Len(registration.AddressLine2) > 0 Then
			str = str & vbCrLf & registration.AddressLine2
		End If
		str = str & vbCrLf & registration.City & ", " & registration.StateID & " " & registration.Zip
		str = str & vbCrLf & vbCrLf & "Phone: " & registration.Phone
		str = str & vbCrLf & "Email: " & registration.Email
		str = str & vbCrLf & "School: " & registration.School
		str = str & vbCrLf & "Grade: " & registration.Grade
		
		If registration.Session = 1 Then
			str = str & vbCrLf & "Session: Beginner (grade 1-3)"
		ElseIf registration.Session = 2 Then
			str = str & vbCrLf & "Session: Intermediate (grade 3-5)"
		ElseIf registration.Session = 3 Then
			str = str & vbCrLf & "Session: Middle School (grade 6)"
		Else
			str = str & vbCrLf & "Session: Not availableo"
		End If
		
		str = str & vbCrLf & "T-Shirt Size: " & registration.TShirtSize
		If Len(registration.Notes) > 0 Then str = str & vbCrLf & "Special Note: " & registration.Notes
		If registration.IsParentHelper = 1 Then
			str = str & vbCrLf & vbCrLf & "Parent Helper: We have noted that you will be willing to help with a team and will contact you with more information - thanks!"
		End If
		str = str & vbCrLf & "--------------------------"
		str = str & vbCrLf & vbCrLf & "The first league session is " & firstSession & ". "
		str = str & "You will receive more information regarding session times and team assignments as the league start date approaches. "
		str = str & "If you have any questions about your registration, reply to this email or send them to mailto:" & fromAddress & ". " 
		str = str & "Thanks, and see you at volleyball!"
		str = str & vbCrLf & vbCrLf & "Lakeville Junior Volleyball"
		str = str & vbCrLf & "21266 Inspiration Path"
		str = str & vbCrLf & "Lakeville, MN 55044"
		str = str & vbCrLf & "mailto:" & fromAddress
		str = str & vbCrLf & "http://" & Request.ServerVariables("SERVER_NAME")
		str = str & vbCrLf & "952.431.6341" & vbCrLf
		
		' send confirm to registrant
		Call email.QuickSend(toAddress, fromAddress, subject, str)
		' send confirm to admin
		Call email.QuickSend(fromAddress, "inhouse_registration@lakevillejuniors.com", "[" & Request.ServerVariables("SERVER_NAME") & "] ** Notify In-House Registration **", "Timestamp: " & Now() & vbCrLf & vbCrLf & "====Registration Details==========================" & vbCrLf & vbCrLf & str)
	End Sub
</script>

<!--#INCLUDE VIRTUAL="/_includes/constants/constants.asp"-->
<!--#INCLUDE VIRTUAL="/_includes/class/registration_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_includes/class/email_sender_cls.asp"-->
