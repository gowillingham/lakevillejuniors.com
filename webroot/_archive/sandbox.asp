<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
    <title>Untitled Page</title>
	<link href="/_includes/styles/default.css" rel="stylesheet" type="text/css" />
	
	<script src="http://www.google.com/jsapi" type="text/javascript" language="javascript"></script>
	<script language="javascript" type="text/javascript">
		google.load("prototype", "1");
		google.load("scriptaculous", "1");
	</script>
</head>
<body>


<h1>Sandbox.asp</h1>

<form action="/_archive/sandbox.asp" method="post" id="form_test">
	<table>
		<tr>
			<td>First Name</td>
			<td><input type="text" name="first_name" id="first_name" class="text_input" /></td>
		</tr>
		<tr>
			<td>Last Name</td>
			<td><input type="text" name="last_name" id="last_name" class="text_input" /></td>
		</tr>
		<tr>
			<td>Email</td>
			<td><input type="text" name="email" id="email" class="text_input" /></td>
		</tr>
		<tr>
			<td>Phone</td>
			<td><input type="text" name="phone" id="phone" class="text_input" /></td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td><input type="submit" name="Submit" value="Save" class="submit_button" /></td>
		</tr>
	</table>
</form>
</body>
</html>
