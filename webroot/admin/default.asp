<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
	<link href="/_includes/styles/default.css" rel="stylesheet" type="text/css" />
	<link href="/_includes/jquery/tablesorter/themes/blue/style.css" rel="stylesheet" type="text/css" />	
		
	<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.2.6/jquery.min.js"></script>
	<script language="javascript" type="text/javascript" src="/_includes/script/jquery/tablesorter/jquery.tablesorter.min.js"></script>

	<script language="javascript" type="text/javascript">
		$(document).ready(function(){
			// set up tablesorter plugin
			$.tablesorter.defaults.widgets = ['zebra']; 
			$("#registrations").tablesorter();
			
			$("table tr:even").addClass("odd");		});
	</script>
	<title>About Lakeville Juniors</title>
</head>
	<body>
		<%Call Main()%>
	</body>
</html>
<%
Sub Main()
	Dim str
	
	str = RegistrationGridToString()
	
	Response.Write str
End Sub

Function RegistrationGridToString()
	Dim str, i
	
	Dim registration		: Set registration = New cRegistration
	Dim list				: list = registration.List()
	
	Dim isOnline			: isOnline = ""
	Dim isConfirmed			: isConfirmed = ""
	
	' 0-RegistrationID 1-RegistrationNumber 2-NameFirstPlayer 3-NameLastPlayer 4-NameFirstParent
	' 5-NameLastParent 6-AddressLine1 7-AddressLine2 8-City 9-StateID 10-Zip
	' 11-Phone 12-Email 13-school 14-TShirtSize 15-Grade 16-IsParentHelper 17-Notes 18-dateCreated
	' 19-DateModified 20-IsOnlineRegistration 21-RegisFee 22-RegisPaid 23-CoachTShirtSize
	' 24-Team 25-IsHeadCoach 26-IsPaymentConfirmed 27-PayPalTransactionId 28-PayPalIsSandbox
	' 29-PayPalPaymentStatus 30-PayPalPaymentStatusReason 31-Session

	If Not IsArray(list) Then
		RegistrationGridToString = "<p>no registrations ..</p>"
		Exit Function
	End If
	
	str = str & "<h1>" & Application.Value("INHOUSE_LEAGUE_VERSION") & " " & Year(Now()) & " in-house registrations (" & UBound(list,2) + 1 & ")</h1>"
	str = str & "<p>(click on a header to sort ..)</p>"
	
	str = str & "<div class=""grid"">"
	str = str & "<table id=""registrations""><thead><tr>"
	
	str = str & "<th>&nbsp;</th>"
	str = str & "<th>Player</th>"
	str = str & "<th>Parent</th>"
	str = str & "<th>Email</th>"
	str = str & "<th>Phone</th>"
	str = str & "<th>School</th>"
	str = str & "<th>Size</th>"
	str = str & "<th>Session</th>"
	str = str & "<th>Online</th>"
	str = str & "<th>PP Status</th>"
	str = str & "<th>Paid</th>"
	str = str & "<th>Confirmed</th>"
	str = str & "<th>Date</th>"
	str = str & "<th>Notes</th>"
	
	str = str & "</tr></thead>"
	str = str & "<tbody>"
	For i = 0 To UBound(list,2)
		isConfirmed = "": If List(26,i) = 1 Then isConfirmed = "yes"
		isOnline = "": If list(20,i) = 1 Then isOnline = "yes"
		str = str & "<tr><td>" & (i + 1) & "</td>"
		str = str & "<td class=""no-break"">" & Server.HTMLEncode(list(3,i) & ", " & list(2,i)) & "</td>"
		str = str & "<td class=""no-break"">" & server.HTMLEncode(list(5,i) & ", " & list(4,i)) & "</td>"
		str = str & "<td>"		
		If Len(list(12,i)) > 0 Then 
			str = str & Server.HTMLEncode(list(12,i & ""))
		End If
		str = str & "</td>"
		str = str & "<td class=""no-break"">" & list(11,i) & "</td>"
		str = str & "<td>" & list(13,i) & "</td>"
		str = str & "<td>" & list(14,i) & "</td>"
		str = str & "<td>" & list(31,i) & "</td>"
		str = str & "<td>" & isOnline & "</td>"
		str = str & "<td>" & list(29,i) & "</td>"
		str = str & "<td class=""currency"">" & FormatCurrency(list(22,i), 2, True) & "</td>"
		str = str & "<td>" & isConfirmed & "</td>"
		str = str & "<td class=""no-break"">" & list(18,i) & "</td>"
		str = str & "<td class=""no-break"">" & Server.HTMLEncode(list(17,i) & "") & "</td>"
		str = str & "</tr>"
	Next
	str = str & "</tbody></table></div>"
	
	RegistrationGridToString = str
End Function
%>

<!--#INCLUDE VIRTUAL="/_includes/constants/constants.asp"-->
<!--#INCLUDE VIRTUAL="/_includes/class/registration_cls.asp"-->

