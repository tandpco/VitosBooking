<%
Option Explicit
Response.buffer = TRUE

If Session("SecurityID") = "" Then
	Response.Redirect("/default.asp")
End If

If Request("o").Count = 0 Then
	Response.Redirect("neworder.asp")
Else
	If Not IsNumeric(Request("o")) Then
		Response.Redirect("neworder.asp")
	End If
End If
%>
<!-- #Include Virtual="include2/globals.asp" -->
<!-- #Include Virtual="include2/math.asp" -->
<!-- #Include Virtual="include2/db-connect.asp" -->
<!-- #Include Virtual="include2/order.asp" -->
<!-- #Include Virtual="include2/customer.asp" -->
<!-- #Include Virtual="include2/address.asp" -->
<!-- #Include Virtual="include2/menu.asp" -->
<!-- #Include Virtual="include2/store.asp" -->
<!-- #Include Virtual="include2/pricing.asp" -->
<!-- #Include Virtual="include2/inventory.asp" -->
<!-- #Include Virtual="include2/printing.asp" -->
<!-- #Include Virtual="include2/coupons.asp" -->
<!-- #Include Virtual="include2/employee.asp" -->
<!-- #Include Virtual="include2/clearorder.asp" -->
<%
Dim gnOrderID, gnSessionID, gsIPAddress, gnEmpID, gsRefID, gdtTransactionDate, gdtSubmitDate, gdtReleaseDate, gdtExpectedDate, gnStoreID, gnCustomerID, gsCustomerName, gsCustomerPhone, gnAddressID, gnOrderTypeID, gbIsPaid, gnPaymentTypeID, gsPaymentReference, gnAccountID, gdDeliveryCharge, gdDriverMoney, gdTax, gdTax2, gdTip, gnOrderStatusID, gsOrderNotes
Dim gbQuickMode
Dim gsOrderTypeDescription
Dim gsEMail, gsFirstName, gsLastName, gdtBirthdate, gnPrimaryAddressID, gsHomePhone, gsCellPhone, gsWorkPhone, gsFAXPhone, gbIsEMailList, gbIsTextList
Dim gsAddress1, gsAddress2, gsCity, gsState, gsPostalCode, gsAddressNotes
Dim gsAddressDescription, gsCustomerNotes
Dim ganUnitIDs(), gasDescriptions(), gasShortDescriptions()
Dim ganOrderLineIDs(), gasOrderLineDescriptions(), ganQuantity(), gadCost(), gadDiscount()
Dim gdOrderTotal
Dim i, gnTmp
Dim ganOrderIDs(), gsLocalErrorMsg
Dim gbNeedPrinterAlert
Dim gasVoidReasons()

' 2012-10-01 TAM: Don't release hold orders during ordering process in case of stuck hold order
'If Not ReleaseHoldOrders(Session("StoreID"), Session("TransactionDate"), ganOrderIDs) Then
'	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
'End If
'
'If ganOrderIDs(0) > 0 Then
'	For i = 0 To UBound(ganOrderIDs)
'		If Not PrintOrder(Session("StoreID"), ganOrderIDs(i), TRUE) Then
'			ResetHoldOrder ganOrderIDs(i)
'			gbNeedPrinterAlert = TRUE
'			gsLocalErrorMsg = "PRINT FAILURE, CANNOT RELEASE HOLD ORDERS!"
'		End If
'	Next
'Else
	gbNeedPrinterAlert = FALSE
'End If

gnOrderID = CLng(Request("o"))
Session("OrderID") = gnOrderID

If GetOrderDetails(gnOrderID, gnSessionID, gsIPAddress, gnEmpID, gsRefID, gdtTransactionDate, gdtSubmitDate, gdtReleaseDate, gdtExpectedDate, gnStoreID, gnCustomerID, gsCustomerName, gsCustomerPhone, gnAddressID, gnOrderTypeID, gbIsPaid, gnPaymentTypeID, gsPaymentReference, gnAccountID, gdDeliveryCharge, gdDriverMoney, gdTax, gdTax2, gdTip, gnOrderStatusID, gsOrderNotes) Then
'	If gnStoreID <> Session("StoreID") Then
'		Response.Redirect("otherstore.asp?t=" & gnOrderTypeID & "&s=" & gnStoreID & "&c=" & gnCustomerID & "&a=" & gnAddressID)
'	End If
	
	If gnOrderStatusID <> 2 And DateValue(gdtTransactionDate) <> DateValue(Session("TransactionDate")) Then
		Response.Redirect("/error.asp?err=" & Server.URLEncode("Order is not From Today"))
	End If
	
	If gnOrderStatusID > 10 Then
		Response.Redirect("/error.asp?err=" & Server.URLEncode("Order is Already Voided"))
	End If

' Don't replace SessionID or IPAddress
'	Session("SessionID") = gnSessionID
'	Session("IPAddress") = gsIPAddress
	Session("SessionID") = Session.SessionID
	Session("IPAddress") = Request.ServerVariables("REMOTE_ADDR")
	gnSessionID = Session("SessionID")
	gsIPAddress = Session("IPAddress")
	Session("RefID") = gsRefID
	Session("SubmitDate") = gdtSubmitDate
	Session("ReleaseDate") = gdtReleaseDate
	Session("ExpectedDate") = gdtExpectedDate
	Session("CustomerID") = gnCustomerID
	Session("CustomerName") = gsCustomerName
	Session("CustomerPhone") = gsCustomerPhone
	Session("AddressID") = gnAddressID
	Session("OrderTypeID") = gnOrderTypeID
	Session("IsPaid") = gbIsPaid
	Session("PaymentTypeID") = gnPaymentTypeID
	Session("PaymentReference") = gsPaymentReference
	Session("AccountID") = gnAccountID
	Session("DeliveryCharge") = gdDeliveryCharge
	Session("DriveMoney") = gdDriverMoney
	Session("Tax") = gdTax
	Session("Tax2") = gdTax2
	Session("Tip") = gdTip
	Session("OrderStatusID") = gnOrderStatusID
	Session("OrderNotes") = gsOrderNotes
	
	gsOrderTypeDescription = GetOrderTypeDescription(gnOrderTypeID)
	Session("OrderTypeDescription") = gsOrderTypeDescription
	
	If GetCustomerDetails(gnCustomerID, gsEMail, gsFirstName, gsLastName, gdtBirthdate, gnPrimaryAddressID, gsHomePhone, gsCellPhone, gsWorkPhone, gsFAXPhone, gbIsEMailList, gbIsTextList) Then
		Session("EMail") = gsEMail
		Session("FirstName") = gsFirstName
		Session("LastName") = gsLastName
		Session("Birthdate") = gdtBirthdate
		Session("PrimaryAddressID") = gnPrimaryAddressID
		Session("HomePhone") = gsHomePhone
		Session("CellPhone") = gsCellPhone
		Session("WorkPhone") = gsWorkPhone
		Session("FAXPhone") = gsFAXPhone
		Session("IsEmailList") = gbIsEMailList
		Session("IsTextList") = gbIsTextList
	Else
		Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
	End If
	
	If GetAddressDetails(gnAddressID, gnTmp, gsAddress1, gsAddress2, gsCity, gsState, gsPostalCode, gsAddressNotes) Then
		Session("Address1") = gsAddress1
		Session("Address2") = gsAddress2
		Session("City") = gsCity
		Session("State") = gsState
		Session("PostalCode") = gsPostalCode
		Session("AddressNotes") = gsAddressNotes
	Else
		Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
	End If
	
' 2012-10-01 TAM: Remove this for now because of bug where address is not associated with the customer
'	If GetCustomerAddressDetails(gnCustomerID, gnAddressID, gsAddressDescription, gsCustomerNotes) Then
'		Session("AddressDescription") = gsAddressDescription
'		Session("CustomerNotes") = gsCustomerNotes
'	Else
'		Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
'	End If
' 2013-10-03 TAM: The following probably fixes the above bug but since we're voiding anyways it's moot and we'll leave this commented out
'	If gnAddressID = 1 Then
'		gsAddressDescription = ""
'		gsCustomerNotes = ""
'	Else
'		If Not GetCustomerAddressDetails(gnCustomerID, gnAddressID, gsAddressDescription, gsCustomerNotes) Then
'			Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
'		End If
'	End If
'	Session("AddressDescription") = gsAddressDescription
'	Session("CustomerNotes") = gsCustomerNotes
Else
	If Len(gsDBErrorMessage) > 0 Then
		Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
	Else
		Response.Redirect("/error.asp?err=" & Server.URLEncode("Invalid Order Specified"))
	End If
End If

If Not GetOrderLines(gnOrderID, ganOrderLineIDs, gasOrderLineDescriptions, ganQuantity, gadCost, gadDiscount) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If ganOrderLineIDs(0) = 0 Then
	Session("OrderLineCount") = 0
Else
	Session("OrderLineCount") = (UBound(ganOrderLineIDs) + 1)
End If

gdOrderTotal = 0.00
If ganOrderLineIDs(0) <> 0 Then
	For i = 0 To UBound(ganOrderLineIDs)
		gdOrderTotal = gdOrderTotal + (ganQuantity(i) * (gadCost(i) - gadDiscount(i)))
	Next
End If
gdOrderTotal = gdOrderTotal + gdDeliveryCharge + gdTax + gdTax2 + gdTip
Session("OrderTotal") = gdOrderTotal

If Not GetVoidReasons(gasVoidReasons) Then
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
<!-- #Include Virtual="include2/clock-server.asp" -->
<script type="text/javascript">
<!--
var ie4=document.all;
var gbVoidReasonInLowerCase = false;

function resetRedirect() {
//	var loRedirectDiv;
//	
//	loRedirectDiv = ie4? eval("document.all.redirect") : document.getElementById("redirect");
//	loRedirectDiv.innerHTML = <%=gnRedirectTime%>;
}

function toggleDivs(psHideDiv, psShowDiv) {
	var loHideDiv, loShowDiv;
	
	loHideDiv = ie4? eval("document.all." + psHideDiv) : document.getElementById(psHideDiv);
	loShowDiv = ie4? eval("document.all." + psShowDiv) : document.getElementById(psShowDiv);
	
	loHideDiv.style.visibility = "hidden";
	loShowDiv.style.visibility = "visible";
	
	resetRedirect();
}

function gotoVoidReason() {
	var loNotes, loDiv;
	
	loNotes = ie4? eval("document.all.voidreason") : document.getElementById('voidreason');
	loNotes.value = "";
	
	loDiv = ie4? eval("document.all.voiddiv") : document.getElementById("voiddiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.voidreasondiv") : document.getElementById("voidreasondiv");
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function addToVoidReason(psDigit) {
	var loNotes, lsNotes;
	
	loNotes = ie4? eval("document.all.voidreason") : document.getElementById('voidreason');
	lsNotes = loNotes.value;
	
	if (psDigit.length > 1) {
		if (lsNotes.length > 0) {
			lsNotes = lsNotes + " ";
		}
	}
	
	if (gbVoidReasonInLowerCase) {
		lsNotes += psDigit.toLowerCase();
	}
	else {
		lsNotes += psDigit;
	}
	
	if (lsNotes.length > 255) {
		lsNotes = lsNotes.substr(0, 255);
	}
	
	loNotes.value = lsNotes;
	
	resetRedirect();
}

function backspaceVoidReason() {
	var loNotes, lsNotes;
	
	loNotes = ie4? eval("document.all.voidreason") : document.getElementById('voidreason');
	lsNotes = loNotes.value;
	if (lsNotes.length > 0) {
		lsNotes = lsNotes.substr(0, (lsNotes.length - 1));
		loNotes.value = lsNotes;
	}
	
	resetRedirect();
}

function clearVoidReason() {
	var loNotes;
	
	loNotes = ie4? eval("document.all.voidreason") : document.getElementById('voidreason');
	loNotes.value = "";
	
	resetRedirect();
}

function properVoidReason() {
	var loNotes, i, lsText, lbDoUpper;
	
	loNotes = ie4? eval("document.all.voidreason") : document.getElementById('voidreason');
	if (loNotes.value.length > 0) {
		lsText = loNotes.value.substr(0, 1).toUpperCase();
		lbDoUpper = false;
		for (i = 1; i < loNotes.value.length; i++) {
			if (loNotes.value.substr(i, 1) == " ") {
				lsText += " ";
				lbDoUpper = true;
			}
			else {
				if (lbDoUpper) {
					lsText += loNotes.value.substr(i, 1).toUpperCase();
					lbDoUpper = false;
				}
				else {
					lsText += loNotes.value.substr(i, 1).toLowerCase();
				}
			}
		}
		loNotes.value = lsText;
	}
	
	resetRedirect();
}

function shiftVoidReason() {
	var loObj;
	
	if (gbVoidReasonInLowerCase) {
		loObj = ie4? eval("document.all.key-q") : document.getElementById('key-q');
		loObj.innerHTML = "Q";
		loObj = ie4? eval("document.all.key-w") : document.getElementById('key-w');
		loObj.innerHTML = "W";
		loObj = ie4? eval("document.all.key-e") : document.getElementById('key-e');
		loObj.innerHTML = "E";
		loObj = ie4? eval("document.all.key-r") : document.getElementById('key-r');
		loObj.innerHTML = "R";
		loObj = ie4? eval("document.all.key-t") : document.getElementById('key-t');
		loObj.innerHTML = "T";
		loObj = ie4? eval("document.all.key-y") : document.getElementById('key-y');
		loObj.innerHTML = "Y";
		loObj = ie4? eval("document.all.key-u") : document.getElementById('key-u');
		loObj.innerHTML = "U";
		loObj = ie4? eval("document.all.key-i") : document.getElementById('key-i');
		loObj.innerHTML = "I";
		loObj = ie4? eval("document.all.key-o") : document.getElementById('key-o');
		loObj.innerHTML = "O";
		loObj = ie4? eval("document.all.key-p") : document.getElementById('key-p');
		loObj.innerHTML = "P";
		loObj = ie4? eval("document.all.key-a") : document.getElementById('key-a');
		loObj.innerHTML = "A";
		loObj = ie4? eval("document.all.key-s") : document.getElementById('key-s');
		loObj.innerHTML = "S";
		loObj = ie4? eval("document.all.key-d") : document.getElementById('key-d');
		loObj.innerHTML = "D";
		loObj = ie4? eval("document.all.key-f") : document.getElementById('key-f');
		loObj.innerHTML = "F";
		loObj = ie4? eval("document.all.key-g") : document.getElementById('key-g');
		loObj.innerHTML = "G";
		loObj = ie4? eval("document.all.key-h") : document.getElementById('key-h');
		loObj.innerHTML = "H";
		loObj = ie4? eval("document.all.key-j") : document.getElementById('key-j');
		loObj.innerHTML = "J";
		loObj = ie4? eval("document.all.key-k") : document.getElementById('key-k');
		loObj.innerHTML = "K";
		loObj = ie4? eval("document.all.key-l") : document.getElementById('key-l');
		loObj.innerHTML = "L";
		loObj = ie4? eval("document.all.key-z") : document.getElementById('key-z');
		loObj.innerHTML = "Z";
		loObj = ie4? eval("document.all.key-x") : document.getElementById('key-x');
		loObj.innerHTML = "X";
		loObj = ie4? eval("document.all.key-c") : document.getElementById('key-c');
		loObj.innerHTML = "C";
		loObj = ie4? eval("document.all.key-v") : document.getElementById('key-v');
		loObj.innerHTML = "V";
		loObj = ie4? eval("document.all.key-b") : document.getElementById('key-b');
		loObj.innerHTML = "B";
		loObj = ie4? eval("document.all.key-n") : document.getElementById('key-n');
		loObj.innerHTML = "N";
		loObj = ie4? eval("document.all.key-m") : document.getElementById('key-m');
		loObj.innerHTML = "M";
	}
	else {
		loObj = ie4? eval("document.all.key-q") : document.getElementById('key-q');
		loObj.innerHTML = "q";
		loObj = ie4? eval("document.all.key-w") : document.getElementById('key-w');
		loObj.innerHTML = "w";
		loObj = ie4? eval("document.all.key-e") : document.getElementById('key-e');
		loObj.innerHTML = "e";
		loObj = ie4? eval("document.all.key-r") : document.getElementById('key-r');
		loObj.innerHTML = "r";
		loObj = ie4? eval("document.all.key-t") : document.getElementById('key-t');
		loObj.innerHTML = "t";
		loObj = ie4? eval("document.all.key-y") : document.getElementById('key-y');
		loObj.innerHTML = "y";
		loObj = ie4? eval("document.all.key-u") : document.getElementById('key-u');
		loObj.innerHTML = "u";
		loObj = ie4? eval("document.all.key-i") : document.getElementById('key-i');
		loObj.innerHTML = "i";
		loObj = ie4? eval("document.all.key-o") : document.getElementById('key-o');
		loObj.innerHTML = "o";
		loObj = ie4? eval("document.all.key-p") : document.getElementById('key-p');
		loObj.innerHTML = "p";
		loObj = ie4? eval("document.all.key-a") : document.getElementById('key-a');
		loObj.innerHTML = "a";
		loObj = ie4? eval("document.all.key-s") : document.getElementById('key-s');
		loObj.innerHTML = "s";
		loObj = ie4? eval("document.all.key-d") : document.getElementById('key-d');
		loObj.innerHTML = "d";
		loObj = ie4? eval("document.all.key-f") : document.getElementById('key-f');
		loObj.innerHTML = "f";
		loObj = ie4? eval("document.all.key-g") : document.getElementById('key-g');
		loObj.innerHTML = "g";
		loObj = ie4? eval("document.all.key-h") : document.getElementById('key-h');
		loObj.innerHTML = "h";
		loObj = ie4? eval("document.all.key-j") : document.getElementById('key-j');
		loObj.innerHTML = "j";
		loObj = ie4? eval("document.all.key-k") : document.getElementById('key-k');
		loObj.innerHTML = "k";
		loObj = ie4? eval("document.all.key-l") : document.getElementById('key-l');
		loObj.innerHTML = "l";
		loObj = ie4? eval("document.all.key-z") : document.getElementById('key-z');
		loObj.innerHTML = "z";
		loObj = ie4? eval("document.all.key-x") : document.getElementById('key-x');
		loObj.innerHTML = "x";
		loObj = ie4? eval("document.all.key-c") : document.getElementById('key-c');
		loObj.innerHTML = "c";
		loObj = ie4? eval("document.all.key-v") : document.getElementById('key-v');
		loObj.innerHTML = "v";
		loObj = ie4? eval("document.all.key-b") : document.getElementById('key-b');
		loObj.innerHTML = "b";
		loObj = ie4? eval("document.all.key-n") : document.getElementById('key-n');
		loObj.innerHTML = "n";
		loObj = ie4? eval("document.all.key-m") : document.getElementById('key-m');
		loObj.innerHTML = "m";
	}
	
	gbVoidReasonInLowerCase = !gbVoidReasonInLowerCase;
	
	resetRedirect();
}

function saveVoidReason() {
	var loNotes;
	
	loNotes = ie4? eval("document.all.voidreason") : document.getElementById('voidreason');
	gsOrderNotes = loNotes.value;
	if (gsOrderNotes.length > 0) {
<%
If Request("Inet") = "yes" Then
%>
		window.location = "neworder.asp?cancel=yes&Inet=yes&VoidReason=" + encodeURIComponent(gsOrderNotes);
<%
Else
%>
		window.location = "neworder.asp?cancel=yes&VoidReason=" + encodeURIComponent(gsOrderNotes);
<%
End If
%>
	}
}
//-->
</script>
</head>

<body onload="clockInit(clockLocalStartTime, clockServerStartTime); clockOnLoad();" onunload="clockOnUnload()">

<div id="mainwindow" style="position: absolute; top: 0px; left: 0px; width=1010px; height: 768px; overflow: hidden;">
<table cellspacing="0" cellpadding="0" width="1010" height="764" border="1">
	<tr>
		<td valign="top" width="1010" height="764">
		<table cellspacing="0" cellpadding="5" width="1010">
			<tr height="31">
				<td valign="top" width="1010">
					<div align="center">
<%
If gbTestMode Then
	If gbDevMode Then
%>
						<strong>DEV SYSTEM
<%
	Else
%>
						<strong>TEST SYSTEM
<%
	End If
End If
%>
						Store <%=Session("StoreID")%></strong> |
						<b><%=Session("name")%></b> |
						<span id="ClockDate"><%=clockDateString(gDate)%></span> |
						<span id="ClockTime" onclick="clockToggleSeconds()"><%=clockTimeString(Hour(gDate), Minute(gDate), Second(gDate))%></span>
					</div>
				</td>
			</tr>
			<tr height="733">
				<td valign="top" width="1010">
					<div id="content" style="position: relative; width: 1010px; height: 723px; overflow: auto;">
						<div id="voiddiv" style="position: absolute; left: 0px; top: 0px; width: 1010px; height: 723px; background-color: #fbf3c5;">
							<table cellpadding="0" cellspacing="0" width="1010" height="723">
								<tr>
									<td valign="top" width="680">
										<p align="center"><strong>Void Completed Order</strong></p>
										<p align="center"><strong><%=gsOrderTypeDescription%>&nbsp;Order #&nbsp;<%=gnOrderID%></strong></p>
										<div style="height: 375px; padding: 10px;">
										<%=gsCustomerName%><br/>
<%
If gnCustomerID = 1 Then
%>
										&nbsp;<br/>
<%
	If Len(gsCustomerPhone) > 0 Then
%>
										Phone: <%=gsCustomerPhone%><br/>
										&nbsp;<br/>
<%
	End If
Else
	If Len(gsAddress2) = 0 Then
		If Len(gsAddress1) > 0 Then
%>
										<%=gsAddress1%><br/>
<%
		End If
	Else
%>
										<%=gsAddress1%> #<%=gsAddress2%><br/>
<%
	End If
	
	If Len(gsCity) > 0 Or Len(gsState) > 0 Or Len(gsPostalCode) > 0 Then
%>
										<%=gsCity%>, <%=gsState%>&nbsp;<%=gsPostalCode%><br/>
<%
	End If
%>
										&nbsp;<br/>
<%
	If Len(gsAddressNotes) > 0 Then
%>
										Address Notes: <%=gsAddressNotes%><br/>
										&nbsp;<br/>
<%
	End If
	
	If Len(gsCustomerPhone) > 0 Then
%>
										Phone: <%=gsCustomerPhone%><br/>
										&nbsp;<br/>
<%
	End If
	
	If Len(gsCustomerNotes) > 0 Then
%>
										Customer Notes: <%=gsCustomerNotes%><br/>
										&nbsp;<br/>
<%
	End If
End If

If Len(gsOrderNotes) > 0 Then
%>
										Order Notes: <%=gsOrderNotes%><br/>
										&nbsp;<br/>
<%
End If
%>
										</div>
										<p align="center"><strong>Are you sure you want to void this completed order?</strong></p>
										<p align="center"><button style="width: 337px;" onclick="gotoVoidReason();">Void This Order</button></p>
										<p align="center">&nbsp;</p>
										<p align="center"><button style="width: 337px;" onclick="window.location = 'neworder.asp';">Cancel</button></p>
									</td>
									<td align="right" valign="top" width="330">
										<div style="position: relative; width: 320px; height: 691px; text-align: left; background-color: #FFFFFF;">
<%
If ganOrderLineIDs(0) <> 0 Then
	For i = 0 To UBound(ganOrderLineIDs)
		If i Mod 4 = 0 Then
			If i > 0 Then
%>
												<button style="width: 320px; color: #FFFFFF; background-color: #FF0000;" onclick="toggleDivs('itemdiv<%=Int(i/4)-1%>', 'itemdiv<%=Int(i/4)%>')">Page <%=Int(i/4)%> of <%=Int(UBound(ganOrderLineIDs)/4)+1%><br/>(Next)</button>
											</div>
											<div id="itemdiv<%=Int(i/4)%>" style="position: absolute; top: 0px; visibility: hidden;">
<%
			Else
%>
											<div id="itemdiv<%=Int(i/4)%>" style="position: absolute; top: 0px;">
<%
			End If
		End If
%>
											<div style="height: 142px; padding: 5px; overflow: auto;">
												<div id="linediv<%=i%>"><%=gasOrderLineDescriptions(i)%></div>
											</div>
<%
	Next
	
	' Add hidden divs here
	If ((UBound(ganOrderLineIDs) + 1) Mod 4) > 0 Then
		For i = ((UBound(ganOrderLineIDs) + 1) Mod 4) To 3
%>
											<div style="height: 152px;">&nbsp;</div>
<%
		Next
	End If
	
	If UBound(ganOrderLineIDs) > 3 Then
%>
												<button style="width: 320px; color: #FFFFFF; background-color: #FF0000;" onclick="toggleDivs('itemdiv<%=Int(UBound(ganOrderLineIDs)/4)%>', 'itemdiv0')">Page <%=Int(UBound(ganOrderLineIDs)/4)+1%> of <%=Int(UBound(ganOrderLineIDs)/4)+1%><br/>(Next)</button>
<%
	End If
End If
%>
											</div>
										</div>
										<div style="width: 320px; text-align: center; background-color: #FFFFFF;">Tax: <%=FormatCurrency(gdTax + gdTax2)%>&nbsp; Delivery: <%=FormatCurrency(gdDeliveryCharge)%>&nbsp; Total: <%=FormatCurrency(gdOrderTotal)%></div>
									</td>
								</tr>
							</table>
						</div>
						<div id="voidreasondiv" style="position: absolute; left: 0px; top: 0px; width: 1010px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
							<table align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="9"><div align="center">
										<strong>WHY IS THIS ORDER BEING VOIDED</strong></div></td>
								</tr>
								<tr>
									<td colspan="9"><div align="center">
										<textarea id="voidreason" style="width: 930px; height: 60px;"></textarea></div></td>
								</tr>
								<tr>
<%
If Len(gasVoidReasons(0)) > 0 Then
	For i = 0 To UBound(gasVoidReasons)
%>
									<td width="11%"><button style="width: 100px;" onclick="addToVoidReason('<%=gasVoidReasons(i)%>')"><%=gasVoidReasons(i)%></button></td>
<%
		If i = 5 Then
			Exit For
		End If
	Next
	
	If UBound(gasVoidReasons) < 5 Then
		For i = 4 To UBound(gasVoidReasons) Step -1
%>
									<td width="11%">&nbsp;</td>
<%
		Next
	End If
Else
%>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
<%
End If
%>
									<td width="11%"><button style="width: 100px;" onclick="window.location = 'neworder.asp';">Cancel</button></td>
									<td width="11%"><button style="width: 100px;" onclick="clearVoidReason();">Clear</button></td>
									<td width="11%"><button style="width: 100px;" onclick="saveVoidReason();">Done</button></td>
								</tr>
								<tr>
<%
If UBound(gasVoidReasons) > 5 Then
	For i = 6 To UBound(gasVoidReasons)
%>
									<td width="11%"><button style="width: 100px;" onclick="addToVoidReason('<%=gasVoidReasons(i)%>')"><%=gasVoidReasons(i)%></button></td>
<%
		If i = 14 Then
			Exit For
		End If
	Next
	
	If UBound(gasVoidReasons) < 14 Then
		For i = 13 To UBound(gasVoidReasons) Step -1
%>
									<td width="11%">&nbsp;</td>
<%
		Next
	End If
Else
%>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
<%
End If
%>
								</tr>
							</table>
							<table align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td><div align="center">
										<button onclick="addToVoidReason('+')">+</button><button onclick="addToVoidReason('!')">!</button><button onclick="addToVoidReason('@')">@</button><button onclick="addToVoidReason('#')">#</button><button onclick="addToVoidReason('$')">$</button><button onclick="addToVoidReason('%')">%</button><button onclick="addToVoidReason('^')">^</button><button onclick="addToVoidReason('&')">&amp;</button><button onclick="addToVoidReason('*')">*</button><button onclick="addToVoidReason('(')">(</button><button onclick="addToVoidReason(')')">)</button><button onclick="addToVoidReason(':')">:</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="addToVoidReason('=')">=</button><button onclick="addToVoidReason('1')">1</button><button onclick="addToVoidReason('2')">2</button><button onclick="addToVoidReason('3')">3</button><button onclick="addToVoidReason('4')">4</button><button onclick="addToVoidReason('5')">5</button><button onclick="addToVoidReason('6')">6</button><button onclick="addToVoidReason('7')">7</button><button onclick="addToVoidReason('8')">8</button><button onclick="addToVoidReason('9')">9</button><button onclick="addToVoidReason('0')">0</button><button onclick="addToVoidReason('?')">?</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="addToVoidReason('\'')">'</button><button name="key-q" id="key-q" onclick="addToVoidReason('Q')">Q</button><button name="key-w" id="key-w" onclick="addToVoidReason('W')">W</button><button name="key-e" id="key-e" onclick="addToVoidReason('E')">E</button><button name="key-r" id="key-r" onclick="addToVoidReason('R')">R</button><button name="key-t" id="key-t" onclick="addToVoidReason('T')">T</button><button name="key-y" id="key-y" onclick="addToVoidReason('Y')">Y</button><button name="key-u" id="key-u" onclick="addToVoidReason('U')">U</button><button name="key-i" id="key-i" onclick="addToVoidReason('I')">I</button><button name="key-o" id="key-o" onclick="addToVoidReason('O')">O</button><button name="key-p" id="key-p" onclick="addToVoidReason('P')">P</button><button onclick="addToVoidReason('/')">/</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="addToVoidReason('.')">.</button><button name="key-a" id="key-a" onclick="addToVoidReason('A')">A</button><button name="key-s" id="key-s" onclick="addToVoidReason('S')">S</button><button name="key-d" id="key-d" onclick="addToVoidReason('D')">D</button><button name="key-f" id="key-f" onclick="addToVoidReason('F')">F</button><button name="key-g" id="key-g" onclick="addToVoidReason('G')">G</button><button name="key-h" id="key-h" onclick="addToVoidReason('H')">H</button><button name="key-j" id="key-j" onclick="addToVoidReason('J')">J</button><button name="key-k" id="key-k" onclick="addToVoidReason('K')">K</button><button name="key-l" id="key-l" onclick="addToVoidReason('L')">L</button><button onclick="addToVoidReason(',')">,</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="shiftVoidReason()">Shift</button><button onclick="addToVoidReason('<')">&lt;</button><button name="key-z" id="key-z" onclick="addToVoidReason('Z')">Z</button><button name="key-x" id="key-x" onclick="addToVoidReason('X')">X</button><button name="key-c" id="key-c" onclick="addToVoidReason('C')">C</button><button name="key-v" id="key-v" onclick="addToVoidReason('V')">V</button><button name="key-b" id="key-b" onclick="addToVoidReason('B')">B</button><button name="key-n" id="key-n" onclick="addToVoidReason('N')">N</button><button name="key-m" id="key-m" onclick="addToVoidReason('M')">M</button><button onclick="addToVoidReason('>')">&gt;</button><button onclick="properVoidReason()">Proper</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button style="width: 150px;" onclick="addToVoidReason('1/2')">1/2</button>&nbsp;<button style="width: 600px;" onclick="addToVoidReason(' ')">Space</button>&nbsp;<button style="width: 150px;" onclick="backspaceVoidReason()">Bksp</button></div></td>
								</tr>
							</table>
						</div>
					</div>
				</td>
			</tr>
		</table>
		</td>
	</tr>
</table>
</div>

<%
If gbNeedPrinterAlert Then
%>
<script type="text/javascript">
<!--
alert("<%=gsLocalErrorMsg%>\nCHECK PRINTER!");
//-->
</script>
<%
End If
%>
</body>

</html>
<!-- #Include Virtual="include2/db-disconnect.asp" -->
