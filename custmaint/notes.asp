<%
Option Explicit
Response.buffer = TRUE

If Session("SecurityID") = "" Then
	Response.Redirect("/default.asp")
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
Dim gsCustomerAddressDescription, gsCustomerAddressNotes,gsCustomerNotes,gsOrderNotes,gsReturnURL
gsReturnURL = "/ordering/unitselect.asp?t="&Session("OrderTypeID")

gnCustomerID = CLng(Request("CustomerID"))
gnAddressID = CLng(Request("AddressID"))
If Request("OrderID").Count > 0 Then
	gnOrderID = CLng(Request("OrderID"))
Else
	gnOrderID = 0
End If
gsAddressNotes = Session("AddressNotes")
gsCustomerNotes = Session("CustomerNotes")
gsCustomerAddressNotes = Session("CustomerAddressNotes")
gsOrderNotes = Session("OrderNotes")

If Request("action") = "savenotes" Then
	Session("AddressNotes") = Request("addressnotes")
	Session("CustomerNotes") = Request("customernotes")
	Session("CustomerAddressNotes") = Request("customeraddressnotes")
	Session("OrderNotes") = Request("ordernotes")
	gsAddressNotes = Session("AddressNotes")
	gsCustomerNotes = Session("CustomerNotes")
	gsCustomerAddressNotes = Session("CustomerAddressNotes")
	gsOrderNotes = Session("OrderNotes")
	Response.Redirect(gsReturnURL)
End If

If gnAddressID > 0 Then
	Call GetAddressDetails(gnAddressID, gnStoreID, gsAddress1, gsAddress2, gsCity, gsState, gsPostalCode, gsAddressNotes)
End If
If gnCustomerID > 0 Then
	Call GetCustomerAddressDetails(gnCustomerID, gnAddressID, gsCustomerAddressDescription, gsCustomerAddressNotes) 
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
	
	loNotes = document.getElementById(gsCurrentField);
	loNotes.value = "";
	
	resetRedirect();
}
function setCurrentField(psCurrentField) {
	var loField;
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
	$("#customerNotes").hide()
	$("#orderNotes").show()
	$("#addressNotes").hide()
	$(".buttons").removeClass('active')
	$(".buttons.orderNotes").addClass('active')
}

function customerNotes() {
	$("#customerNotes").show()
	$("#orderNotes").hide()
	$("#addressNotes").hide()
	$(".buttons").removeClass('active')
	$(".buttons.customerNotes").addClass('active')
}

function addressNotes() {
	$("#customerNotes").hide()
	$("#orderNotes").hide()
	$(".buttons").removeClass('active')
	$(".buttons.addressNotes").addClass('active')
	$("#addressNotes").show()
}

function saveNotes() {
	$("#SaveAddressNotes").val($("#addressnotesField").val());
	$("#SaveCustomerAddressNotes").val($("#customeraddressnotesField").val());
	$("#SaveOrderNotes").val($("#orderNotesField").val());
	$("#SaveCustomerNotes").val($("#customerNotesField").val());
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
					<button style="width: 300px;margin-left:50px" class="active buttons addressNotes" onclick="addressNotes()">Address Notes</button>
					<button style="width: 300px;" class="buttons customerNotes" onclick="customerNotes()">Customer Notes</button>
					<button style="width: 300px;" class="buttons orderNotes" onclick="orderNotes()">One Time Notes</button>
					<form style="display:none" action="notes.asp?action=savenotes" id="formNotes" method="post">
						<input type="hidden" name="addressnotes" id="SaveAddressNotes">
						<input type="hidden" name="customeraddressnotes" id="SaveCustomerAddressNotes">
						<input type="hidden" name="customernotes" id="SaveCustomerNotes">
						<input type="hidden" name="ordernotes" id="SaveOrderNotes">
					</form>
					<div id="content" style="position: relative; width: 1010px; height: 723px; overflow: auto;">
						<div id="addressNotes">
							<table width="925" align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="3">
										<table align="center" cellpadding="0" cellspacing="0">
											<tr>
												<td><strong>General Address Notes:</strong><br/>
												<textarea id="addressnotesField" name="addressnotes" style="width: 450px; height: 60px;" onfocus="setCurrentField('addressnotesField');"><%=Server.HTMLEncode(gsAddressNotes)%></textarea></td>
												<td>&nbsp;</td>
												<td><strong>Customer Specific Address Notes:</strong><br/>
												<textarea id="customeraddressnotesField" name="customeraddressnotes" style="width: 450px; height: 60px" onfocus="setCurrentField('customeraddressnotesField');"><%=Server.HTMLEncode(gsCustomerAddressNotes)%></textarea></td>
											</tr>
										</table>
									</td>
								</tr>
							</table>
						</div>
						<div id="customerNotes" style="display:none">
							<table width="925" align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="3">
										<table align="center" cellpadding="0" cellspacing="0">
											<tr>
												<td><strong>Customer Notes:</strong><br/>
												<textarea id="customerNotesField" name="customernotes" style="width: 900px; height: 60px;" onfocus="setCurrentField('customerNotesField');"><%=Server.HTMLEncode(gsCustomerNotes)%></textarea></td>
											</tr>
										</table>
									</td>
								</tr>
							</table>
						</div>
						<div id="orderNotes" style="display:none">
							<table width="925" align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="3">
										<table align="center" cellpadding="0" cellspacing="0">
											<tr>
												<td><strong>One Time Notes:</strong><br/>
												<textarea id="orderNotesField" name="ordernotes" style="width: 900px; height: 60px;" onfocus="setCurrentField('orderNotesField');"><%=Server.HTMLEncode(gsOrderNotes)%></textarea></td>
											</tr>
										</table>
									</td>
								</tr>
							</table>
						</div>
						<table width="925" align="center" cellpadding="0" cellspacing="0">
							<tr>
								<td><button style="width: 100px;background-color:#c83100" onclick="window.location = '<%=gsReturnURL%>'">Cancel</button><button style="width: 100px;" onclick="clearCurrentField()">Clear</button><button style="width: 100px;float:right;background-color:green" onclick="saveNotes()">Done</button></td>
								<td align="right">
						      <% If Session("CustomerID") > 1 AND Session("AddressID") > 1 Then %>
						      <button style="width: 300px;" onclick="window.location='/ordering/unitselect.asp?t=<%=Session("OrderTypeID")%>&c=<%=Session("CustomerID")%>&a=<%=Session("AddressID")%>&amp;ordernotes=1'">Edit Order Notes</button>
						      <% End If %>
						    </td>
							</tr>
						</table>
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
