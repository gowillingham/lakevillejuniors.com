<html>
	<head>
		<title>Sandbox</title>
	</head>
	<body>
		<h1>Test IPN Post</h1>
		<p>
			<form action="/ipn.asp" method="post">
				<table>
					<tr><td>guid</td>
					<td><input type="text" name="invoice" style="width:300px;"/></td></tr>
					<tr><td>mc_gross</td>
					<td><input type="text" name="mc_gross" value="60" /></td></tr>
					<tr><td>test_ipn</td>
					<td><input type="text" name="test_ipn" value="1" /></td></tr>
					<tr><td>txn_id</td>
					<td><input type="text" name="txn_id" value="paypal transaction id" /></td></tr>
					<tr><td>payment_status</td>
					<td><input type="text" name="payment_status" value="Completed" /></td></tr>
					<tr><td>payment_status_reason</td>
					<td><input type="text" name="payment_status_reason" value="eCheck" /></td></tr>
					<tr><td>business</td>
					<td><input type="text" name="business" value="willin_1199747222_biz@lakevillejuniors.com" /></td></tr>
					<tr><td>mc_currency</td>
					<td><input type="text" name="mc_currency" value="USD" /></td></tr>
					<tr><td>&nbsp;</td><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td>
					<td><input type="submit" name="submit" value="Post to IPN" /></td></tr>
				</table>
			</form>
		</p>
	</body>
</html>