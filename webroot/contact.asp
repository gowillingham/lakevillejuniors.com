<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
	<link href="_includes/styles/default.css" rel="stylesheet" type="text/css" />
	<title>Lakeville Juniors Contact Us</title>
</head>
	<body>
		<!--#INCLUDE VIRTUAL="/_includes/maincontent.asp"-->
	</body>
</html>
<%Sub Main()%>
	<h1>Contact Lakeville Juniors</h1>
	<p>
		<div><b>Lakeville Junior Volleyball</b></div>
		<div>21266 Inspiration Path</div>
		<div>Lakeville, MN 55044</div>
		<div style="padding:5px 0 0 0;"><a href="mailto:<%=Application.Value("INFO_EMAIL_ADDRESS")%>" title="Contact Lakeville Juniors"><%=Application.Value("INFO_EMAIL_ADDRESS")%></a></div>
	</p>
<%End Sub%>
