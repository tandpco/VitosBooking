<%
Option Explicit
Response.buffer = TRUE

If Session("SecurityID") = "" Then
	Response.Redirect("/default.asp")
End If

Dim gsReturnURL

If Request("CustomerID").Count > 0 Then
	If Not IsNumeric(Request("CustomerID")) Then
		Response.Redirect("../main.asp")
	End If
Else
	Response.Redirect("../main.asp")
End If

If Request("AddressID").Count > 0 Then
	If Not IsNumeric(Request("AddressID")) Then
		Response.Redirect("../main.asp")
	End If
Else
	Response.Redirect("../main.asp")
End If

If Request("OrderID").Count > 0 Then
	If Not IsNumeric(Request("OrderID")) Then
		Response.Redirect("../main.asp")
	End If
End If


' Highlight the currently selected tab.
Dim currentTab
currentTab = "notes"
%>
<!-- #Include Virtual="include2/globals.asp" -->
<!-- #Include Virtual="include2/math.asp" -->
<!-- #Include Virtual="include2/db-connect.asp" -->
<!-- #Include Virtual="include2/customer.asp" -->
<!-- #Include Virtual="include2/address.asp" -->
<%
Dim gnCustomerID, gnAddressID, gnOrderID
Dim gnStoreID, gsAddress1, gsAddress2, gsCity, gsState, gsPostalCode, gsAddressNotes
Dim gsCustomerAddressDescription, gsCustomerAddressNotes

gnCustomerID = CLng(Request("CustomerID"))
gnAddressID = CLng(Request("AddressID"))
If Request("OrderID").Count > 0 Then
	gnOrderID = CLng(Request("OrderID"))
Else
	gnOrderID = 0
End If

If Request("ReturnURL").Count > 0 Then
	gsReturnURL = Request("ReturnURL")
Else
	If LCase(Right(Request.ServerVariables("HTTP_REFERER"), 14)) = "viewdrives.asp" Then
		gsReturnURL = "../viewdrives.asp"
	Else
		gsReturnURL = "editcustomer.asp?o=" & gnOrderID & "&c=" & gnCustomerID & "&a=" & gnAddressID
	End If
End If

If Request("action") = "savenotes" Then
	gsAddressNotes = Request("SaveAddressNotes")
	gsCustomerAddressNotes = Request("SaveCustomerNotes")
	
	If Not UpdateAddressNotes(gnAddressID, gsAddressNotes) Then
		Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
	End If
	
	If Not UpdateCustomerAddressNotes(gnCustomerID, gnAddressID, gsCustomerAddressNotes) Then
		Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
	End If
	
	Response.Redirect(gsReturnURL)
End If

If Not GetAddressDetails(gnAddressID, gnStoreID, gsAddress1, gsAddress2, gsCity, gsState, gsPostalCode, gsAddressNotes) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If Not GetCustomerAddressDetails(gnCustomerID, gnAddressID, gsCustomerAddressDescription, gsCustomerAddressNotes) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="en-us" http-equiv="Content-Language" />
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>Vito's Point of Sale</title>
<link rel="stylesheet" href="/css/vitos.css" type="text/css" />
<script src="http://code.jquery.com/jquery-latest.js"></script>

    <link rel="stylesheet" href="/Scripts/keyboard/css/jsKeyboard.css" type="text/css" media="screen"/>
    <script type="text/javascript" src="/Scripts/keyboard/jsKeyboard.js"></script>
<!-- #Include Virtual="include2/clock-server.asp" -->
<script type="text/javascript">
var ie4=document.all;

var gsCurrentField = "addressnotes";

function resetRedirect() {}

function clearCurrentField() {
	var loNotes;
	
	loNotes = ie4? eval("document.all." + gsCurrentField) : document.getElementById(gsCurrentField);
	loNotes.value = "";
	
	resetRedirect();
}
function setCurrentField(psCurrentField) {
	var loField;
	
	loField = ie4? eval("document.all." + gsCurrentField) : document.getElementById(gsCurrentField);
	loField.style.backgroundColor = "#CCCCCC";
	
	loField = ie4? eval("document.all." + psCurrentField) : document.getElementById(psCurrentField);
	loField.style.backgroundColor = "#FFFFFF";
	
	gsCurrentField = psCurrentField;
	
	resetRedirect();
}
function toggleDivs(psHideDiv, psShowDiv) {
	var loHideDiv, loShowDiv;
	
	loHideDiv = ie4? eval("document.all." + psHideDiv) : document.getElementById(psHideDiv);
	loShowDiv = ie4? eval("document.all." + psShowDiv) : document.getElementById(psShowDiv);
	
	loHideDiv.style.visibility = "hidden";
	loShowDiv.style.visibility = "visible";
	
	resetRedirect();
}

function orderNotes() {
	$("#orderlinenotes").hide()
	$("#ordernotes").show()
}

function saveNotes() {
	$("#SaveAddressNotes").val($("#addressnotes").val());
	$("#SaveCustomerNotes").val($("#customeraddressnotes").val());
	$("#formNotes").submit();
}
$(function(){
  jsKeyboard.init("virtualKeyboard");
  jsKeyboard.show();
});

</script>
</head>

<body onload="clockInit(clockLocalStartTime, clockServerStartTime); clockOnLoad();" onunload="clockOnUnload()">

<div id="mainwindow" style="position: absolute; top: 0px; left: 0px; width=1010px; height: 768px; overflow: hidden;">
<table cellspacing="0" cellpadding="0" width="1010" height="764" border="1">
	<tr>
		<td valign="top" width="1010" height="764">
		<table cellspacing="0" cellpadding="5" width="1010">
			<!-- #Include Virtual="ordering/top-header.asp" -->
			<tr height="733">
				<td valign="top" width="1010">
					<div id="content-wrapper" style="top:10px">
					<div id="content" style="position: relative; width: 1010px; height: 723px; overflow: auto;">
						<div id="orderlinenotes" style="position: absolute; left: 0px; top: 0px; width: 1010px; height: 723px; background-color: #fbf3c5;">
							<form id="formNotes" name="formNotes" method="post" action="addressnotes.asp">
								<input type="hidden" id="action" name="action" value="savenotes" />
								<input type="hidden" id="ReturnURL" name="ReturnURL" value="<%=gsReturnURL%>" />
								<input type="hidden" id="CustomerID" name="CustomerID" value="<%=gnCustomerID%>" />
								<input type="hidden" id="AddressID" name="AddressID" value="<%=gnAddressID%>" />
								<input type="hidden" id="SaveAddressNotes" name="SaveAddressNotes" value="<%=gsAddressNotes%>" />
								<input type="hidden" id="SaveCustomerNotes" name="SaveCustomerNotes" value="<%=gsCustomerAddressNotes%>" />
								<p align="center"><strong>Editing Notes For <%=gsAddress1%>&nbsp;<%=gsAddress2%>, <%=gsCity%>, <%=gsState%>&nbsp;<%=gsPostalCode%></strong></p>
							</form>
							<table width="925" align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="3">
										<table align="center" cellpadding="0" cellspacing="0">
											<tr>
												<td><strong>General Address Notes:</strong><br/>
												<textarea id="addressnotes" name="addressnotes" style="width: 450px; height: 60px;" onfocus="setCurrentField('addressnotes');"><%=Server.HTMLEncode(gsAddressNotes)%></textarea></td>
												<td>&nbsp;</td>
												<td><strong>Customer Specific Address Notes:</strong><br/>
												<textarea id="customeraddressnotes" name="customeraddressnotes" style="width: 450px; height: 60px; background-color: #cccccc;" onfocus="setCurrentField('customeraddressnotes');"><%=Server.HTMLEncode(gsCustomerAddressNotes)%></textarea></td>
											</tr>
										</table>
									</td>
								</tr>
							</table>
							<table width="925" align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td><button style="width: 100px;" onclick="window.location = '<%=gsReturnURL%>'">Cancel</button><button style="width: 100px;" onclick="clearCurrentField()">Clear</button><button style="width: 100px;" onclick="saveNotes()">Done</button></td>
									<td align="right">
							      <% If Session("CustomerID") > 1 AND Session("AddressID") > 1 Then %>
							      <button style="width: 300px;" onclick="window.location='/ordering/unitselect.asp?t=<%=Session("OrderTypeID")%>&c=<%=Session("CustomerID")%>&a=<%=Session("AddressID")%>&amp;ordernotes=1'">Edit Order Notes</button>
							      <% End If %>
							    </td>
								</tr>
							</table>
						</div>
            <div id="virtualKeyboard"></div>
					</div>
					</div>
				</td>
			</tr>
		</table>
		</td>
	</tr>
</table>
</div>

</body>

</html>
<!-- #Include Virtual="include2/db-disconnect.asp" -->
