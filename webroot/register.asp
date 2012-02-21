<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%

Sub OnPageLoad(page)
	page.Action = Request.QueryString("act")
	page.RegistrationID = Request.QueryString("rid")
	
	Set page.Registration = New cRegistration
	page.Registration.RegistrationID = page.RegistrationID
End Sub
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
	<link href="_includes/styles/default.css" rel="stylesheet" type="text/css" />
	<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.2.6/jquery.min.js"></script>
	<title>Lakeville Juniors Online Registration</title>
	
	<script type="text/javascript" language="javascript">
	
		function setSchoolOptions(val) {
			$("#school-dropdown-list option").attr("disabled", "")
			
			if (val == 6) {
				$(".elementary-school").attr("disabled", "disabled")
			}
			else {
				$(".middle-school").attr("disabled", "disabled")
			}
		};
		
		function setSessionOptions(val) {
			$("#session-dropdown-list option").attr("disabled", "disabled")
			
			if (val == 1 || val == 2) {
				$(".beginner-session").attr("disabled", "")
			}
			if (val == 3) {
				$(".beginner-session").attr("disabled", "")
				$(".intermediate-session").attr("disabled", "")
			}
			if (val == 4 || val == 5) {
				$(".intermediate-session").attr("disabled", "")
			}
			if (val == 6) {
				$(".advanced-session").attr("disabled", "")
			}
		};
		
		function setMessages(val) {
			$("#grade-3-message-row").hide()
			$("#grade-6-message-row").hide()
			
			if (val == 3) {
				$("#grade-3-message-row").show()
			}
			else if (val == 6) {
				$("#grade-6-message-row").show()
			};
		};
	
		$(document).ready(function(){
			
			// hide/show session dropdown based on state of grade dropdown ..
			$("#grade-3-message-row").hide()
			$("#grade-6-message-row").hide()
			
			if ($("#grade-dropdown-list").val() > 0) {
				$("#session-dropdown-row").show()
				$("#school-dropdown-row").show()
			}
			else {
				$("#session-dropdown-row").hide()
				$("#school-dropdown-row").hide()
			};
	
			// display message for grade 3 or 6
			if ($("#grade-dropdown-list").val() == 3) {
				$("#grade-3-message-row").show()
			}
			else if ($("#grade-dropdown-list").val() == 6) {
				$("#grade-6-message-row").show()
			};
			
			setMessages($("#grade-dropdown-list").val());
			setSchoolOptions($("#grade-dropdown-list").val());
			setSessionOptions($("#grade-dropdown-list").val());
			
			$("#grade-dropdown-list").change(function(){
				var value = $(this).val();
				
				// reset school, session dropdowns on change ..
				$("#school-dropdown-list").val("")
				$("#session-dropdown-list").val("")
				
				setMessages(value);
				setSchoolOptions(value);
				setSessionOptions(value);

				// show/hide grade 3/6 messages and school/session dropdowns
				if (value > 0) {
					$("#session-dropdown-row").show()
					$("#school-dropdown-row").show()
				}
				else {
					$("#session-dropdown-row").hide()
					$("#school-dropdown-row").hide()
				};
			});
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
	
	If Application.Value("IS_REGISTRATION_ENABLED") = False Then
		Response.Write GetCustomAppMsg("We're Sorry, online registration is not available at this time.", "Please contact <a href=""mailto:" & Application.Value("INFO_EMAIL_ADDRESS")& """ title=""Email Lakeville Juniors"" style=""text-decoration:underline;"">Lakeville Juniors</a> if you would like more information.", "error", "/_images/alert.png") 
		Exit Sub
	End If
	
	Call OnPageLoad(page)

	' cancel event handler ..
	If Request.Form("Submit") = "Cancel" Then
		Select Case page.Action
			Case INSERT_REGISTRATION
				page.Action = ""
			Case Else
				page.Action = DELETE_REGISTRATION
		End Select
		Response.Redirect(Request.ServerVariables("URL") & page.UrlParamsToString(False))
	End If

	Select Case page.Action
		Case INSERT_REGISTRATION	
			If Request.Form("FormRegisterIsPostback") = IS_POSTBACK Then

				Call LoadRegistrationFromForm(page.Registration)
				If IsValidRegister(page) Then

					Call DoInsertRegistration(page.Registration, rv)
					Select Case rv
						Case 0
							page.Action = ACCEPT_PAYMENT: page.RegistrationID = page.Registration.RegistrationID
						Case Else
					End Select
					Response.Redirect(Request.ServerVariables("URL") & page.UrlParamsToString(False)) 
				Else
					str = str & FormRegisterToString(page)
				End If
			Else
				str = str & FormRegisterToString(page)
			End If
		
		Case EDIT_REGISTRATION
			If Request.Form("FormRegisterIsPostback") = IS_POSTBACK Then
				Call page.Registration.Load()
				Call LoadRegistrationFromForm(page.Registration)
				If IsValidRegister(page) Then
					Call DoUpdateRegistration(page.Registration, rv)
					Select Case rv
						Case 0
							page.Action = ACCEPT_PAYMENT: page.RegistrationID = page.Registration.RegistrationID
						Case Else
					End Select
					Response.Redirect(Request.ServerVariables("URL") & page.UrlParamsToString(False)) 
				Else
					str = str & FormRegisterToString(page)
				End If
			Else
				Call page.Registration.Load()
				str = str & FormRegisterToString(page)
			End If

		Case ACCEPT_PAYMENT
			' regis has been inserted/updated, so display PAY_NOW button				
			str = str & FormAcceptPaymentToString(page)
			
		Case CONFIRM_REGISTRATION_COMPLETE
			str = str & ConfirmRegistrationCompleteToString(page)
			
		Case DELETE_REGISTRATION
			' runs when user cancels from paypal or cancels from FormAcceptPayment()
			Call DoDeleteRegistration(page.Registration, rv)
			page.Action = "": page.RegistrationID = ""
			str = str & FormRegisterToString(page)
			
		Case Else
			str = str & FormRegisterToString(page)

	End Select

	Response.Write str
	Set page = Nothing
End Sub

Sub DoInsertRegistration(registration, rv)
	registration.IsPaymentConfirmed = 0
	registration.IsOnlineRegistration = 1
	registration.RegisPaid = 0
	registration.PayPalIsSandbox = 0
	Call registration.Add(rv)
End Sub

Sub DoUpdateRegistration(registration, rv)
	Call registration.Save(rv)
End Sub

Sub DoDeleteRegistration(registration, rv)
	Call registration.Delete(rv)
End Sub

Function ConfirmRegistrationCompleteToString(page)
	Dim str
	
	If Len(page.Registration.RegistrationID) > 0 Then page.Registration.Load()
	Dim helperText			: helperText = "Not this time!"
	If page.Registration.IsParentHelper = 1 Then helperText = "Yes!!"
	
	Dim dateText
	If IsDate(Application.Value("INHOUSE_LEAGUE_FIRST_SESSION_DATE")) Then 
		dateText = FormatDateTime(Application.Value("INHOUSE_LEAGUE_FIRST_SESSION_DATE"), vbLongDate)
	Else
		dateText = Application.Value("INHOUSE_LEAGUE_FIRST_SESSION_DATE")
	End If
	
	str = str & "<h1>Thank You - Your In-House Registration is Complete!</h1>"
	str = str & "<p>We have recieved your registration and payment for In-House Volleyball. "
	str = str & "You will receive confirmation of your payment and registration by email in the next few moments "
	str = str & "(be sure your spam filter is set to allow email from @lakevillejuniors.com). "
	str = str & "Your registration information is included below. "
	str = str & "If it is necessary to change any of your registration information, you may do so by contacting "
	str = str & "<a href=""mailto:" & Application.Value("INFO_EMAIL_ADDRESS") & """ title=""Email"">Lakeville Juniors</a> "
	str = str & "with changes before the session start " & dateText & ". </p>"
	str = str & "<p>You will receive more information regarding session sites, times, and team assignments by email and US mail as the start date approaches. </p>"
	str = str & "<table class=""invoice"" style=""width:100%;font-size:0.9em"">"
	str = str & "<tr><td colspan=""2"" class=""header"">Registration Info for " & Server.HTMLEncode(page.Registration.NameFirstPlayer & " " & page.Registration.NameLastPlayer) & "</td></tr>"
	str = str & "<tr><td class=""label"">Registration #: </td><td>" & page.Registration.RegistrationNumber & "</td></tr>"
	str = str & "<tr><td class=""label"">Player Name: </td><td>" & Server.HTMLEncode(page.Registration.NameLastPlayer & ", " & page.Registration.NameFirstPlayer) & "</td></tr>"
	str = str & "<tr><td class=""label"">Parent Name: </td><td>" & Server.HTMLEncode(page.Registration.NameLastParent1 & ", " & page.Registration.NameFirstParent1) & "</td></tr>"
	str = str & "<tr><td class=""label"">Address: </td><td>"
	str = str & Server.HTMLEncode(page.Registration.AddressLine1)
	If Len(page.Registration.AddressLine2) > 0 Then str = str & "<br />" & Server.HTMLEncode(page.Registration.AddressLine2)
	str = str & "<br />" & Server.HTMLEncode(page.Registration.City & ", MN " & page.Registration.Zip) & "</td></tr>"
	str = str & "<tr><td class=""label"">Email Address: </td><td>" & Server.HTMLEncode(page.Registration.Email) & "</td></tr>"
	str = str & "<tr><td class=""label"">Phone: </td><td>" & Server.HTMLEncode(page.Registration.Phone) & "</td></tr>"
	str = str & "<tr><td class=""label"">School: </td><td>" & Server.HTMLEncode(page.Registration.School) & "</td></tr>"
	str = str & "<tr><td class=""label"">Grade: </td><td>" & Server.HTMLEncode(page.Registration.Grade) & "</td></tr>"
	str = str & "<tr><td class=""label"">T-Shirt Size: </td><td>" & Server.HTMLEncode(page.Registration.TShirtSize) & "</td></tr>"
	str = str & "<tr><td class=""label"">Parent Helper: </td><td>" & Server.HTMLEncode(helperText) & "</td></tr>"
	If page.Registration.IsParentHelper = 1 Then 
		str = str & "<tr><td class=""label"">Helper T-Shirt Size: </td><td>" & Server.HTMLEncode(page.Registration.CoachTShirtSize) & "</td></tr>"
	End If
	If page.Registration.IsPaymentConfirmed = 1 Then 
		str = str & "<tr><td class=""label"">Registration Fee: </td><td>" & FormatCurrency(page.Registration.RegisPaid, 2) & " Paid</td></tr>"
	End If
	If Len(page.Registration.Notes) > 0 Then
		str = str & "<tr><td class=""label"">Notes: </td><td>" & Server.HTMLEncode(page.Registration.Notes) & "</td></tr>"
	End If
	str = str & "<tr><td class=""label"">Start Date: </td><td>" & dateText & "</td></tr>"
	str = str & "</table>" 
	
	ConfirmRegistrationCompleteToString = str
End Function

Function LoadRegistrationFromForm(registration)
	registration.NameFirstPlayer = Trim(Request.Form("NameFirstPlayer"))
	registration.NameLastPlayer = Trim(Request.Form("NameLastPlayer"))
	registration.NameFirstParent1 = Trim(Request.Form("NameFirstParent1"))
	registration.NameLastParent1 = Trim(Request.Form("NameLastParent1"))
	registration.AddressLine1 = Trim(Request.Form("AddressLine1"))
	registration.AddressLine2 = Trim(Request.Form("AddressLine2"))
	registration.City = Trim(Request.Form("City"))
	registration.StateID = Trim(Request.Form("StateID"))
	registration.Zip = Trim(Request.Form("Zip"))
	registration.Email = Trim(Request.Form("Email"))
	registration.EmailRetype = Trim(Request.Form("EmailRetype"))
	registration.Phone = Trim(Request.Form("Phone"))
	registration.TShirtSize = Trim(Request.Form("TShirtSize"))
	registration.School = Trim(Request.Form("School"))
	registration.Grade = Trim(Request.Form("Grade"))
	registration.Session = Trim(Request.Form("session"))
	registration.Notes = Trim(Request.Form("Notes"))

	If Len(Request.Form("IsParentHelper")) > 0 Then
		registration.IsParentHelper = 1 
		registration.CoachTShirtSize = Trim(Request.Form("CoachTShirtSize"))
	Else
		registration.IsParentHelper = 0 
		registration.CoachTShirtSize = ""
	End If
	
	If Len(Request.Form("HasRelease")) > 0 Then
		registration.HasRelease = 1
	Else
		Registration.HasRelease = 0
	End If
End Function

Function IsValidRegister(page)
	Dim rv
	rv = True
	
	If Len(page.Registration.Grade & "") = 0 Then
		AddCustomFrmError("Please provide a grade for the player who is registering. ")
		rv = False
	End If
	If Len(page.Registration.School & "") = 0 Then
		AddCustomFrmError("School is required (you can't choose a school until you have indicated the player's grade). ")
		rv = False
	End If
	If Len(page.Registration.Session & "") = 0 Then
		AddCustomFrmError("Please choose a session (you can't choose a session until you have indicated the player's grade). ")
	End If
	If Not ValidData(page.Registration.NameFirstPlayer, True, 0, 50, "Player First Name", "") Then rv = False
	If Not ValidData(page.Registration.NameLastPlayer, True, 0, 50, "Player Last Name", "") Then rv = False
	If Not ValidData(page.Registration.NameFirstParent1, True, 0, 50, "Parent First Name", "") Then rv = False
	If Not ValidData(page.Registration.NameLastParent1, True, 0, 50, "Player Last Name", "") Then rv = False
	If Not ValidData(page.Registration.AddressLine1, True, 0, 100, "Address Line 1", "") Then rv = False
	If Not ValidData(page.Registration.AddressLine2, False, 0, 50, "Address Line 2", "") Then rv = False
	If Not ValidData(page.Registration.City, True, 0, 50, "City", "") Then rv = False
	If Not ValidData(page.Registration.Zip, True, 0, 10, "Zip", "zip") Then rv = False
	If Not ValidData(page.Registration.Phone, True, 0, 14, "Phone", "phone") Then rv = False
	If Not ValidData(page.Registration.Email, True, 0, 50, "Email", "email") Then rv = False
	If UCase(page.Registration.Email) <> UCase(page.Registration.EmailRetype) Then
		rv = False
		AddCustomFrmError("Email must match Email Retype")
	End If
	If Not ValidData(page.Registration.TShirtSize, True, 0, 0, "T-Shirt Size", "") Then rv = False
	If Not ValidData(page.Registration.Notes, False, 0, 2000, "Notes", "") Then rv = False
	If page.Registration.IsParentHelper = 1 And Len(page.Registration.CoachTShirtSize) = 0 Then
		rv = False
		AddCustomFrmError("If you will be a parent helper, we'll need a t-shirt size for your coaching shirt")
	End If
	If page.Registration.HasRelease = 0 Then 
		rv = False
		AddCustomFrmError("Please indicate that you have read and agree to the liability release. ")
	End If 
	
	IsValidRegister = rv
End Function

Function FormAcceptPaymentToString(page)
	Dim str
	Dim pg					: Set pg = page.Clone()
	
	If Len(page.Registration.RegistrationID) > 0 Then page.Registration.Load()
	Dim helperText			: helperText = "Not this time!"
	If page.Registration.IsParentHelper = 1 Then helperText = "Yes!!"
	
	str = str & "<h1>Confirm Registration</h1>"
	str = str & "<p>Please confirm that your registration information is correct. "
	str = str & "Select <strong>Edit</strong> if you wish to modify any of your registration information. </p>"
	
	str = str & "<div style=""float:right;width:30%;margin:0;text-align:center;"">"
	
	' --- paypal code to display logo ------
	str = str & "<a href=""#"" "
	str = str & "onclick=""javascript:window.open('https://www.paypal.com/us/cgi-bin/webscr?cmd=xpt/cps/popup/OLCWhatIsPayPal-outside','olcwhatispaypal','toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes, width=400, height=350');"">"
	str = str & "<img src=""https://www.paypal.com/en_US/i/bnr/vertical_solution_PPeCheck.gif"" alt=""Solution Graphics"" /></a>"
	' --------------------------------------
	
	str = str & "<p style=""text-align:center;font-size:0.8em;"">Make your Lakeville Juniors payments online by either credit card or your PayPal account.</p>"
	str = str & "</div>"
	str = str & "<table class=""invoice"" style=""width:65%;font-size:0.9em"">"
	str = str & "<tr><td colspan=""2"" class=""header"">Registration Info</td></tr>"
	str = str & "<tr><td class=""label"">Player Name: </td><td>" & Server.HTMLEncode(page.Registration.NameLastPlayer & ", " & page.Registration.NameFirstPlayer) & "</td></tr>"
	str = str & "<tr><td class=""label"">Parent Name: </td><td>" & Server.HTMLEncode(page.Registration.NameLastParent1 & ", " & page.Registration.NameFirstParent1) & "</td></tr>"
	str = str & "<tr><td class=""label"">Address: </td><td>"
	str = str & Server.HTMLEncode(page.Registration.AddressLine1)
	If Len(page.Registration.AddressLine2) > 0 Then str = str & "<br />" & Server.HTMLEncode(page.Registration.AddressLine2)
	str = str & "<br />" & Server.HTMLEncode(page.Registration.City & ", MN " & page.Registration.Zip) & "</td></tr>"
	str = str & "<tr><td class=""label"">Email Address: </td><td>" & Server.HTMLEncode(page.Registration.Email) & "</td></tr>"
	str = str & "<tr><td class=""label"">Phone: </td><td>" & Server.HTMLEncode(page.Registration.Phone) & "</td></tr>"
	str = str & "<tr><td class=""label"">School: </td><td>" & Server.HTMLEncode(page.Registration.School) & "</td></tr>"
	str = str & "<tr><td class=""label"">Grade: </td><td>" & Server.HTMLEncode(page.Registration.Grade) & "</td></tr>"
	str = str & "<tr><td class=""label"">Session: </td><td>" & GetSessionForID(page.Registration.Session) & "</td></tr>"
	str = str & "<tr><td class=""label"">T-Shirt Size: </td><td>" & Server.HTMLEncode(page.Registration.TShirtSize) & "</td></tr>"
	str = str & "<tr><td class=""label"">Parent Helper: </td><td>" & Server.HTMLEncode(helperText) & "</td></tr>"
	If page.Registration.IsParentHelper = 1 Then 
		str = str & "<tr><td class=""label"">Helper T-Shirt Size: </td><td>" & Server.HTMLEncode(page.Registration.CoachTShirtSize) & "</td></tr>"
	End If
	If Len(page.Registration.Notes) > 0 Then
		str = str & "<tr><td class=""label"">Notes: </td><td>" & Server.HTMLEncode(page.Registration.Notes) & "</td></tr>"
	End If
	pg.Action = EDIT_REGISTRATION
	str = str & "<tr><td>&nbsp;</td><td style=""text-align:left;padding-top:20px;"">"
	str = str & "<form style=""display:inline;"" class=""form"" method=""post"" action=""" & request.ServerVariables("URL") & pg.UrlParamsToString(True) & """ name=""formEditRegistration"">"
	str = str & "<input type=""submit"" name=""submit"" value=""Edit"" title=""Edit my Information"" class=""button"" />"
	str = str & "<input type=""hidden"" name=""FormEditRegistrationIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "</form>"
	pg.Action = DELETE_REGISTRATION
	str = str & "&nbsp;<form style=""display:inline;"" class=""form"" method=""post"" action=""" & request.ServerVariables("URL") & pg.UrlParamsToString(True) & """ name=""formDeleteRegistration"">"
	str = str & "<input type=""submit"" name=""submit"" value=""Cancel"" title=""Cancel my Registration"" class=""button"" />"
	str = str & "<input type=""hidden"" name=""FormDeleteRegistrationIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "</form></td></tr></table>" 

	str = str & "<h1>Authorize Payment</h1>"
	str = str & "<p>When you select the <strong>Pay Now</strong> button you will be redirected by secure server to PayPal to make your payment. "
	str = str & "Once redirected, you will have the option to pay by credit card or with a PayPal account. "
	str = str & "You are not required to have a PayPal account. </p>"
	
	str = str & "<table class=""invoice"">"
	str = str & "<tr class=""header""><td>Item</td><td>Quantity</td><td>Total Price</td><td>&nbsp;</td></tr>"
	str = str & "<tr style=""vertical-align:middle;""><td>In-House Volleyball Registration</td><td>1</td><td>" & FormatCurrency(pg.Registration.RegisFee, 2) & "</td>"
	str = str & "<td style=""width:1%;text-align:right;"">"
	str = str & "<form style=""display:inline;"" action=""" & Application.Value("PAYPAL_GATEWAY") & """ method=""post"">"
	str = str & "<input type=""hidden""  name=""cmd"" value=""_xclick"" />"
	str = str & "<input type=""hidden""  name=""business"" value=""" & Application.Value("PAYPAL_BUSINESS_ID") & """ />"
	str = str & "<input type=""hidden""  name=""item_name"" value=""In-House Volleyball Registration"" />"
	str = str & "<input type=""hidden""  name=""invoice"" value=""" & pg.Registration.RegistrationID & """ />"
	str = str & "<input type=""hidden""  name=""item_number"" value=""" & pg.Registration.RegistrationNumber & """ />"
	str = str & "<input type=""hidden""  name=""amount"" value=""" & FormatCurrency(pg.Registration.RegisFee, 2) & """ />"
	str = str & "<input type=""hidden""  name=""no_shipping"" value=""0"" />"
	pg.Action = CONFIRM_REGISTRATION_COMPLETE
	str = str & "<input type=""hidden""  name=""return"" value=""http://" & Request.ServerVariables("SERVER_NAME") & "/register.asp" & pg.UrlParamsToString(True) & """ />"
	pg.Action = EDIT_REGISTRATION
	str = str & "<input type=""hidden""  name=""cancel_return"" value=""http://" & Request.ServerVariables("SERVER_NAME") & "/register.asp" & pg.UrlParamsToString(True) & """ />"
	str = str & "<input type=""hidden""  name=""no_note"" value=""1"" />"
	str = str & "<input type=""hidden""  name=""currency_code"" value=""USD"" />"
	str = str & "<input type=""hidden""  name=""lc"" value=""US"" />"
	str = str & "<input type=""hidden""  name=""bn"" value=""PP-BuyNowBF"" />"
	str = str & "<input type=""image"" src=""https://www.sandbox.paypal.com/en_US/i/btn/btn_paynow_SM.gif"" name=""submit"" alt=""Make payments with PayPal - it's fast, free and secure!"" />"
	str = str & "<img alt="""" border=""0"" src=""https://www.sandbox.paypal.com/en_US/i/scr/pixel.gif"" width=""1"" height=""1"" />"
	' pre-populate credit card fields ..
	str = str & "<input type=""hidden"" name=""first_name"" value=""" & pg.Registration.NameFirstParent1 & """ />"
	str = str & "<input type=""hidden"" name=""last_name"" value=""" & pg.Registration.NameLastParent1 & """ />"
	str = str & "<input type=""hidden"" name=""address1"" value=""" & pg.Registration.AddressLine1 & """ />"
	str = str & "<input type=""hidden"" name=""address2"" value=""" & pg.Registration.AddressLine2 & """ />"
	str = str & "<input type=""hidden"" name=""city"" value=""" & pg.Registration.City & """ />"
	str = str & "<input type=""hidden"" name=""state"" value=""" & pg.Registration.StateID & """ />"
	str = str & "<input type=""hidden"" name=""zip"" value=""" & pg.Registration.zip & """ />"
	str = str & "<input type=""hidden"" name=""night_phone_a"" value=""" & Left(pg.Registration.PhoneRaw, 3) & """ />"
	str = str & "<input type=""hidden"" name=""night_phone_b"" value=""" & Right(Left(pg.Registration.PhoneRaw, 6), 3) & """ />"
	str = str & "<input type=""hidden"" name=""night_phone_c"" value=""" & Right(pg.Registration.PhoneRaw, 4) & """ />"
	
	str = str & "</form></td></tr></table>"
	
	FormAcceptPaymentToString = str
End Function

Function FormRegisterToString(page)
	Dim str
	Dim isChecked		: isChecked = ""
	Dim pg				: Set pg = page.Clone()
	If Len(pg.Action) = 0 Then pg.Action = INSERT_REGISTRATION
	
	str = str & "<h1>Register for Lakeville Juniors In-House Volleyball</h1>"
	
	str = str & "<p>Use the form below to register for In-House Volleyball. "
	str = str & "Information marked with '" & RequiredElementToString("", True) & "' is required for registration. "
	str = str & "After clicking 'Continue', you'll need a credit card (or PayPal account) to make online payment. "
	str = str & "Once your registration is completed, you'll receive confirmation/receipt of your registration by email. "
	str = str & "When all teams have been formed (a few days after the registration deadline), expect confirmation by email with your first practice details and team assignments. </p>"
	str = str & "<p style=""padding-bottom:10px;"">"
	str = str & "If you prefer not to register online, you may download a pdf version of the registration form "
	str = str & "(<a href=""/_files/inhouse_flyer.pdf"" title=""Download Registration"">grade 1-5</a>). </p>"

	' this message shows if javascript is not enabled ..
	str = str & "<noscript>"
	str = str & ApplicationMessageToString("error", "Hey!", "It looks like javascript is turned off in your browser. Javascript must be enabled in your browser in order to register. Please turn on javascript and then refresh the page. ")
	str = str & "</noscript>"
	
	str = str & FormErrorToString()
	str = str & "<form class=""form"" action=""" & Request.ServerVariables("URL") & pg.UrlParamsToString(True) & """ method=""post"" name=""formRegister"">"
	str = str & "<table><tr><td class=""label"">" & RequiredElementToString("Grade", True) & "</td>"
	str = str & "<td>" & GradeDropdownToString(page.Registration.Grade) & "</td></tr>"
	str = str & "<tr id=""grade-3-message-row""><td colspan=""2"">"
	str = str & ApplicationMessageToString("confirm", "Grade 3 players may register for beginner or intermediate. ", "If your grade 3 player has attended two or more previous sessions, they may register for the intemediate in-house session. ")
	str = str & "</td></tr>"
	str = str & "<tr id=""grade-6-message-row""><td colspan=""2"">"
	str = str & ApplicationMessageToString("confirm", "Grade 6 players register for middle school session. ", "All middle school volleyball players (grade 6) should register for the middle school session. ")
	str = str & "</td></tr>"
	str = str & "<tr id=""session-dropdown-row""><td class=""label"">" & RequiredElementToString("Session", True) & "</td>"
	str = str & "<td>" & SessionDropdownToString(page.Registration.Session) & "</td></tr>"
	str = str & "<tr id=""school-dropdown-row""><td class=""label"">" & RequiredElementToString("School", True) & "</td>"
	str = str & "<td>" & SchoolDropdownToString(page.Registration.School) & "</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	
	str = str & "<tr><td class=""label"">" & RequiredElementToString("Player First Name", True) & "</td>"
	str = str & "<td><input type=""text"" name=""NameFirstPlayer"" value=""" & page.Registration.NameFirstPlayer & """ /></td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString("Player Last Name", True) & "</td>"
	str = str & "<td><input type=""text"" name=""NameLastPlayer"" value=""" & page.Registration.NameLastPlayer & """ /></td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString("Parent First Name", True) & "</td>"
	str = str & "<td><input type=""text"" name=""NameFirstParent1"" value=""" & page.Registration.NameFirstParent1 & """ /></td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString("Parent Last Name", True) & "</td>"
	str = str & "<td><input type=""text"" name=""NameLastParent1"" value=""" & page.Registration.NameLastParent1 & """ /></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString("Address Line 1", True) & "</td>"
	str = str & "<td><input class=""medium"" type=""text"" name=""AddressLine1"" value=""" & page.Registration.AddressLine1 & """ /></td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString("Address Line 2", False) & "</td>"
	str = str & "<td><input class=""medium"" type=""text"" name=""AddressLine2"" value=""" & page.Registration.AddressLine2 & """ /></td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString("City", True) & "</td>"
	str = str & "<td><input type=""text"" name=""City"" value=""" & page.Registration.City & """ /></td></tr>"		
	str = str & "<tr><td class=""label"">" & RequiredElementToString("State", True) & "</td>"
	str = str & "<td><select name=""StateID""><option value=""MN"">MN - Minnesota</option></select></td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString("Postal Code", True) & "</td>"
	str = str & "<td><input class=""small"" type=""text"" name=""Zip"" value=""" & page.Registration.Zip & """ /></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString("Email", True) & "</td>"
	str = str & "<td><input class=""medium"" type=""text"" name=""Email"" value=""" & page.Registration.Email & """ /></td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString("Retype Email", True) & "</td>"
	str = str & "<td><input class=""medium"" type=""text"" name=""EmailRetype"" value=""" & page.Registration.EmailRetype & """ /></td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString("Phone", True) & "</td>"
	str = str & "<td><input type=""text"" name=""Phone"" value=""" & page.Registration.Phone & """ /></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString("T-Shirt Size", True) & "</td>"
	str = str & "<td><select name=""TShirtSize""><option value="""">&nbsp;</option>"
	str = str & SelectOption(GetTShirtSizes(), page.Registration.TShirtSize) & "</select></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"

	str = str & "<tr><td class=""label"">" & RequiredElementToString("Any Special Notes for This Registration", False) & "</td>"
	str = str & "<td><textarea class=""medium"" name=""Notes"" rows=""10"" cols=""10"">" & page.Registration.Notes & "</textarea></td></tr>"
	If page.Registration.IsParentHelper = 1 Then isChecked = " checked=""checked"""
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td colspan=""2""><div class=""hint"" style=""margin:0 100px 0px 0;"">"
	str = str & "<img class=""icon"" src=""/_images/add.png"" alt="""" />"
	str = str & "Would you consider being a parent helper for your daughter's team? "
	str = str & "You don't need to know anything about volleyball. We'll make it easy and fun!!</div></td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString("Parent Helper", False) & "</td>"
	str = str & "<td><input class=""checkbox"" type=""checkbox"" name=""IsParentHelper""" & isChecked & " />"
	str = str & "&nbsp;Yes, I'll be a parent helper!!</td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString("Parent/Coach T-Shirt Size", False) & "</td>"
	str = str & "<td><select name=""CoachTShirtSize""><option value="""">&nbsp;</option>"
	str = str & SelectOption(GetTShirtSizes(), page.Registration.CoachTShirtSize) & "</select></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	
	str = str & "<tr><td colspan=""2""><div class=""hint"" style=""background-color:#fff79f;margin:0 100px 0px 0;border:solid 1px #e70000;padding:10px;"">"
	str = str & "<h3 style=""margin-top:0;"">" & RequiredElementToString("Liability Release", True) & "</h3>"
	str = str & "In consideration of your acceptance of my child as a participant in the Lakeville Junior Volleyball program, "
	str = str & "I hereby waive all claims against Lakeville Junior Volleyball, Inc., "
	str = str & "its coaches, board members, organizers, "
	str = str & "persons transporting my child to a from activities, from any claim arising out of injury or harm to my child incidental to, "
	str = str & "connected with, or arising out of the Lakeville Junior Volleyball activities. "
	str = str & "All sports have intrinsic hazards. Participation in the Lakeville Junior Volleyball in-house session implies some risk of injury. "
	str = str & "</div></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">Release of liability</td>"
	str = str & "<td><input class=""checkbox"" type=""checkbox"" name=""HasRelease""" & isChecked & " />"
	str = str & "&nbsp;" & RequiredElementToString("I have read and agree to the liability release.", True) & "</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td>&nbsp;</td>"
	str = str & "<td><input type=""submit"" name=""Submit"" value=""Continue"" class=""button"" />"
	str = str & "&nbsp;<input type=""submit"" name=""Submit"" value=""Cancel"" class=""button"" />"
	str = str & "</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>"
	
	' ----- paypal code to display logo -------------
	str = str & "<a href=""#"" onclick=""javascript:window.open('https://www.paypal.com/us/cgi-bin/webscr?cmd=xpt/cps/popup/OLCWhatIsPayPal-outside','olcwhatispaypal','toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes, width=400, height=350');"">"
	str = str & "<img src=""/_images/horizontal_solution_PPeCheck.gif"" alt=""PayPal Accepted"" style=""margin-top:35px;"" /></a>"	
	' -----------------------------------------------
	
	str = str & "<div style=""width:215px;text-align:center;margin:5px;font-size:0.8em;"">"
	str = str & "Make your Lakeville Juniors payments online by either credit card or your PayPal account.</div>"
	str = str & "</td></tr></table>"
	str = str & "<input type=""hidden"" name=""FormRegisterIsPostback"" value=""" & IS_POSTBACK & """ /></form>"
	
	FormRegisterToString = str
End Function 

Function GetSessionForID(id)
	Dim str
	
	If CStr(id) = CStr(SESSION_BEGINNER) Then str = "Beginner (Grade 1-3)"
	If CStr(id) = CStr(SESSION_INTERMEDIATE) Then str = "Intermediate (Grade 3-5)"
	If CStr(id) = CStr(SESSION_ADVANCED) Then str = "Middle school (Grade 6)"

	GetSessionForID = str
End Function

Function SessionDropdownToString(val)
	Dim str
	
	Dim selected		: selected = ""
	
	str = str & "<select name=""session"" id=""session-dropdown-list"">"
	str = str & "<option value="""">&nbsp;</option>"
	
	selected = ""		: If CStr(val) = CStr(SESSION_BEGINNER) Then selected = " selected=""selected"""
	str = str & "<option class=""beginner-session"" value=""" & SESSION_BEGINNER & """" & selected & ">Beginner (Grade 1-3)</option>"
	
	selected = ""		: If CStr(val) = CStr(SESSION_INTERMEDIATE) Then selected = " selected=""selected"""
	str = str & "<option class=""intermediate-session"" value=""" & SESSION_INTERMEDIATE & """" & selected & ">Intermediate (Grade 3-5)</option>"
	
	selected = ""		: If CStr(val) = CStr(SESSION_ADVANCED) Then selected = " selected=""selected"""
	str = str & "<option class=""advanced-session"" value=""" & SESSION_ADVANCED & """" & selected & ">Middle School (Grade 6)</option>"
	
	str = str & "</select>"
	
	SessionDropdownToString = str
End Function

Function SchoolDropdownToString(id)
	Dim str, i
	Dim arr()			: ReDim arr(1,14)
	
	Dim selected		: selected = ""
	
	arr(0,0) = "ASE"
	arr(1,0) = "All Saints"
	arr(0,1) = "CVE"
	arr(1,1) = "Cherryview"
	arr(0,2) = "CHE"
	arr(1,2) = "Christina Huddleston"
	arr(0,3) = "CLE"
	arr(1,3) = "Crystal Lake"
	arr(0,4) = "EVE"
	arr(1,4) = "Eastview"
	arr(0,5) = "JFK"
	arr(1,5) = "John F Kennedy"
	arr(0,6) = "LME"
	arr(1,6) = "Lake Marion"
	arr(0,7) = "LVE"
	arr(1,7) = "Lakeview"
	arr(0,8) = "OHE"
	arr(1,8) = "Oak Hills"
	arr(0,9) = "OLE"
	arr(1,9) = "Orchard Lake"
	
	arr(0,10) = "ASM"
	arr(1,10) = "All Saints Middle"
	arr(0,11) = "CMS"
	arr(1,11) = "Century Middle"
	arr(0,12) = "KTMS"
	arr(1,12) = "Kenwood Trail Middle"
	arr(0,13) = "MMS"
	arr(1,13) = "McGuire Middle"
	
	arr(0,14) = "NA"
	arr(1,14) = "My school isn't listed"
	
	str = str & "<select id=""school-dropdown-list"" name=""school"">"
	str = str & "<option value="""">&nbsp;</option>"
	
	' elementary schools ..
	str = str & "<optgroup label=""Lakeville elementary school"">"
	For i = 0 To 10
		selected = "":			If CStr(id) = CStr(arr(0,i)) Then selected = " selected=""selected"""
		
		str = str & "<option value=""" & arr(0,i) & """" & selected & " class=""elementary-school"">" & arr(1,i) & "</option>"
	Next
	str = str & "</optgroup>"
	
	' middle schools ..
	str = str & "<optgroup label=""Lakeville middle school"">"
	For i = 10 To 13
		selected = "":			If CStr(id) = CStr(arr(0,i)) Then selected = " selected=""selected"""
		
		str = str & "<option value=""" & arr(0,i) & """" & selected & " class=""middle-school"">" & arr(1,i) & "</option>"
	Next
	str = str & "</optgroup>"
	
	' other schools ..
	str = str & "<optgroup label=""Other school"">"
	For i = 14 To 14
		selected = "":			If CStr(id) = CStr(arr(0,i)) Then selected = " selected=""selected"""
		
		str = str & "<option value=""" & arr(0,i) & """" & selected & ">" & arr(1,i) & "</option>"
	Next
	str = str & "</optgroup>"
	
	str = str & "</select>"
	
	SchoolDropdownToString = str
End Function

Function GradeDropdownToString(id)
	Dim str, i
	Dim arr			: ReDim arr(1,5)
	
	Dim selected	: selected = ""

	arr(0,0) = "1"
	arr(1,0) = "1"
	arr(0,1) = "2"
	arr(1,1) = "2"
	arr(0,2) = "3"
	arr(1,2) = "3"
	arr(0,3) = "4"
	arr(1,3) = "4"
	arr(0,4) = "5"
	arr(1,4) = "5"
'	arr(0,5) = "6"
'	arr(1,5) = "6"
	
	str = str & "<select id=""grade-dropdown-list"" name=""grade"">"
	str = str & "<option value="""">&nbsp;</option>"
	For i = 0 To UBound(arr,2)
		selected = "":			If CStr(id) = CStr(arr(0,i)) Then selected = " selected=""selected"""
		
		str = str & "<option value=""" & arr(0,i) & """" & selected & ">" & arr(1,i) & "</option>"
	Next
	str = str & "</select>"
	
	GradeDropdownToString = str
End Function

Function GetTShirtSizes()
	Dim arr()
	ReDim arr(1,5)
	
	arr(0,0) = "YM"
	arr(1,0) = "Youth Medium"
	arr(0,1) = "YL"
	arr(1,1) = "Youth Large"
	arr(0,2) = "S"
	arr(1,2) = "Adult Small"
	arr(0,3) = "M"
	arr(1,3) = "Adult Medium"
	arr(0,4) = "L"
	arr(1,4) = "Adult Large"
	arr(0,5) = "XL"
	arr(1,5) = "Adult X-Large"
	
	GetTShirtSizes = arr
End Function
%>
<!--#INCLUDE VIRTUAL="/_includes/script/frm_select_option.asp"-->
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
