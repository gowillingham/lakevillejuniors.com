	<table id="Layout">
		<tr id="HeaderRow">
			<td id="HeaderLeft"><a class="HeaderLink" href="/" title="Home Page">lakevillejuniors.com</a></td>
			<td id="HeaderMid">&nbsp;</td>
			<td id="HeaderRight"></td>
		</tr>					
		<tr>
			<td id="TopMenu" colspan="2"><!--#INCLUDE VIRTUAL="/_includes/topnav.asp"--></td>
			<td id="TopMenuRight"><%=FormatDateTime(Now(), vbLongDate) %></td>
		</tr>
		<tr>
			<td colspan="3" id="Content">
				<table id="ContentTable">
					<tr>
						<td id="LeftPane"><!--#INCLUDE VIRTUAL="/_includes/leftpane.asp"--></td>
						<td id="ContentPane" style="vertical-align:top;">
							<%Call Main() %>
						</td>
						<td id="RightPane"></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td id="FooterRow">
				<table id="FooterTable">
					<tr>
						<td id="FooterLeft"><!--#INCLUDE VIRTUAL="/_includes/footer.asp"--></td>
						<td id="FooterMid">&nbsp;</td>
						<td id="FooterRight">&nbsp;</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
