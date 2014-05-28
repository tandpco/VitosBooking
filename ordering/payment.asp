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
Dim ganEmpIDs(), ganEmployeeIDs(), gasCardIDs()
Dim ganAccountIDs(), gasAccountNames(), gabAccountOnHolds()
Dim ganAllAccountIDs(), gasAllAccountNames(), gabAllAccountOnHolds()
Dim gbNeedPrinterAlert
Dim lnTop, lnLeft
Dim ganCollegeDebitAccountIDs(), gasCollegeDebitAccountNames(), gabCollegeDebitAccountOnHolds()

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

If Request("EditReason").Count > 0 Then
	If Len(Trim(Request("EditReason"))) > 0 Then
		Session("EditReason") = Trim(Request("EditReason"))
	End If
End If

gnOrderID = CLng(Request("o"))
Session("OrderID") = gnOrderID
If GetOrderDetails(gnOrderID, gnSessionID, gsIPAddress, gnEmpID, gsRefID, gdtTransactionDate, gdtSubmitDate, gdtReleaseDate, gdtExpectedDate, gnStoreID, gnCustomerID, gsCustomerName, gsCustomerPhone, gnAddressID, gnOrderTypeID, gbIsPaid, gnPaymentTypeID, gsPaymentReference, gnAccountID, gdDeliveryCharge, gdDriverMoney, gdTax, gdTax2, gdTip, gnOrderStatusID, gsOrderNotes) Then
'	If gnStoreID <> Session("StoreID") Then
'		Response.Redirect("otherstore.asp?t=" & gnOrderTypeID & "&s=" & gnStoreID & "&c=" & gnCustomerID & "&a=" & gnAddressID)
'	End If
	
	If gnOrderStatusID <> 2 And DateValue(gdtTransactionDate) <> DateValue(Session("TransactionDate")) Then
		Response.Redirect("/error.asp?err=" & Server.URLEncode("Order is not From Today"))
	End If
	
' 2013-08-20 TAM: Allow edit order even if complete and paid
'	If gnOrderStatusID >= 10 Then
'		Response.Redirect("/error.asp?err=" & Server.URLEncode("Order is Complete"))
'	End If
'	
'	If gbIsPaid Then
'		Response.Redirect("/error.asp?err=" & Server.URLEncode("Order Has Been Paid"))
'	End If

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
	
	If gnAddressID = 1 Then
		gsAddressDescription = ""
		gsCustomerNotes = ""
	Else
		If Not GetCustomerAddressDetails(gnCustomerID, gnAddressID, gsAddressDescription, gsCustomerNotes) Then
			Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
		End If
	End If
	Session("AddressDescription") = gsAddressDescription
	Session("CustomerNotes") = gsCustomerNotes
	
	' 2014-01-02 TAM: Bug fix for using pay now with unreleased orders
	If gnOrderStatusID = 1 And InStr(LCase(Request.ServerVariables("HTTP_REFERER")), "ticket.asp") > 0 Then
		Session("OrderEdited") = TRUE
		Session("NewOrder") = TRUE
	End If
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

If Not GetAllStoreManagers(gnStoreID, ganEmpIDs, ganEmployeeIDs, gasCardIDs) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

'If Not GetCustomerStoreAccounts(gnCustomerID, gnStoreID, ganAccountIDs, gasAccountNames, gabAccountOnHolds) Then
'	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
'End If
If Not GetStoreAccounts(gnStoreID, ganAccountIDs, gasAccountNames, gabAccountOnHolds) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If
If Not GetAllAccounts(gnStoreID, ganAllAccountIDs, gasAllAccountNames, gabAllAccountOnHolds) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If
If Not GetStoreCollegeDebitAccounts(gnStoreID, ganCollegeDebitAccountIDs, gasCollegeDebitAccountNames, gabCollegeDebitAccountOnHolds) Then
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
var gnOrderID = <%=gnOrderID%>;
var gdOrderTotal = <%=gdOrderTotal%>;
var gnPaymentTypeID = <%=gnPaymentTypeID%>;
var gsPaymentReference = "<%=gsPaymentReference%>";
var gdTenderCash = 0.00;
var gdTenderCheck = 0.00;
var gdTenderCreditCard = 0.00;
var gdTenderOnAccount = 0.00;
var gdTenderBalance = <%=gdOrderTotal%>;
var gdTipAmount = 0.00;
var gbDoPrint = true;
var gbGoAcct = false;

var i;

i = <%=UBound(ganEmpIDs) + 1%>;
var ganEmpIDs = new Array(i);
var ganEmployeeIDs = new Array(i);
var gasCardIDs = new Array(i);
<%
For i = 0 To UBound(ganEmpIDs)
%>
	ganEmpIDs[<%=i%>] = <%=ganEmpIDs(i)%>;
	ganEmployeeIDs[<%=i%>] = <%=ganEmployeeIDs(i)%>;
	gasCardIDs[<%=i%>] = "<%=gasCardIDs(i)%>";
<%
Next
%>

if (gnPaymentTypeID == 4) {
	gsPaymentReference = "<%=gnAccountID%>";
}

if (gnPaymentTypeID == 0) {
	gnPaymentTypeID = 1;
}

function resetRedirect() {
//	var loRedirectDiv;
//	
//	loRedirectDiv = ie4? eval("document.all.redirect") : document.getElementById("redirect");
//	loRedirectDiv.innerHTML = <%=gnRedirectTime%>;
}

function disableEnterKey() {
	var loText, loDiv;
	
	if (event.keyCode == 13) {
		event.cancelBubble = true;
		event.returnValue = false;
		return false;
	}
}

function toggleDivs(psHideDiv, psShowDiv) {
	var loHideDiv, loShowDiv;
	
	loHideDiv = ie4? eval("document.all." + psHideDiv) : document.getElementById(psHideDiv);
	loShowDiv = ie4? eval("document.all." + psShowDiv) : document.getElementById(psShowDiv);
	
	loHideDiv.style.visibility = "hidden";
	loShowDiv.style.visibility = "visible";
	
	resetRedirect();
}

function FormatCurrency(amount)
{
	var i = parseFloat(amount);
	if(isNaN(i)) { i = 0.00; }
	var minus = '';
	if(i < 0) { minus = '-'; }
	i = Math.abs(i);
	i = parseInt((i + .005) * 100);
	i = i / 100;
	s = new String(i);
	if(s.indexOf('.') < 0) { s += '.00'; }
	if(s.indexOf('.') == (s.length - 2)) { s += '0'; }
	s = '$' + minus + s;
	return s;
}

function getCash() {
	var loCash, loDiv;

	loCash = ie4? eval("document.all.cash") : document.getElementById('cash');
	loCash.value = "";
	
<%
If Session("SecurityID") > 1 And ((Left(Request.ServerVariables("remote_addr"), 9) = "10.0.254." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.1." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.2.") OR Session("Swipe")) Then
%>
	loDiv = ie4? eval("document.all.tenderdiv") : document.getElementById('tenderdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.checkdiv") : document.getElementById('checkdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountdiv") : document.getElementById('accountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountprintdiv") : document.getElementById('accountprintdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tipdiv") : document.getElementById('tipdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById('confirmreprint');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.allaccountdiv") : document.getElementById('allaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.collegedebitaccountdiv") : document.getElementById('collegedebitaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.cashdiv") : document.getElementById('cashdiv');
	loDiv.style.visibility = "visible";
<%
Else
%>
	setCash()
<%
End If
%>
	
	resetRedirect();
}

function addToCash(psDigit) {
	var loText, lsText, loRE;
	
	loText = ie4? eval("document.all.cash") : document.getElementById('cash');
	loRE = /\./i;
	lsText = loText.value.replace(loRE, "");
	while (lsText.substr(0, 1) == "0") {
		lsText = lsText.substr(1);
	}
	lsText += psDigit;
	switch (lsText.length) {
		case 0:
			loText.value = lsText;
			break;
		case 1:
			loText.value = "0.0" + lsText;
			break;
		case 2:
			loText.value = "0." + lsText;
			break;
		default:
			loText.value = lsText.substr(0, (lsText.length - 2)) + "." + lsText.substr((lsText.length - 2));
			break;
	}
	
	resetRedirect();
}

function setCashTo(psDigits) {
	var loCash;
	
	loCash = ie4? eval("document.all.cash") : document.getElementById('cash');
	loCash.value = psDigits;
	
	setCash();
}

function backspaceCash() {
	var loText, lsText, loRE;
	
	loText = ie4? eval("document.all.cash") : document.getElementById('cash');
	loRE = /\./i;
	lsText = loText.value.replace(loRE, "");
	while (lsText.substr(0, 1) == "0") {
		lsText = lsText.substr(1);
	}
	if (lsText.length > 0) {
		lsText = lsText.substr(0, (lsText.length - 1));
		switch (lsText.length) {
			case 0:
				loText.value = lsText;
				break;
			case 1:
				loText.value = "0.0" + lsText;
				break;
			case 2:
				loText.value = "0." + lsText;
				break;
			default:
				loText.value = lsText.substr(0, (lsText.length - 2)) + "." + lsText.substr((lsText.length - 2));
				break;
		}
	}
	
	resetRedirect();
}

function cancelCash() {
	var loCash, loDiv;
	
	loCash = ie4? eval("document.all.cash") : document.getElementById('cash');
	loCash.value = "";
	
	loDiv = ie4? eval("document.all.cashdiv") : document.getElementById('cashdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.checkdiv") : document.getElementById('checkdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountdiv") : document.getElementById('accountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountprintdiv") : document.getElementById('accountprintdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tipdiv") : document.getElementById('tipdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById('confirmreprint');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.allaccountdiv") : document.getElementById('allaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.collegedebitaccountdiv") : document.getElementById('collegedebitaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tenderdiv") : document.getElementById('tenderdiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function setCash() {
	var loCash, ldCash, loDiv;
	
	loCash = ie4? eval("document.all.cash") : document.getElementById('cash');
	if (loCash.value == "") {
		ldCash = 0.00;
	}
	else {
		ldCash = new Number(loCash.value);
	}
	gnPaymentTypeID = 1;
	gdTenderCash = Math.round(ldCash * Math.pow(10,2))/Math.pow(10,2);
	gdTenderCheck = 0.00;
	gdTenderCreditCard = 0.00;
	gdTenderOnAccount = 0.00;
	gdTenderBalance = <%=gdOrderTotal%> - gdTenderCash;
	gdTenderBalance = Math.round(gdTenderBalance * Math.pow(10,2))/Math.pow(10,2);
	gsPaymentReference = "";
	gdTipAmount = 0.00;
	
	loDiv = ie4? eval("document.all.tendertip") : document.getElementById('tendertip');
	loDiv.innerHTML = "&nbsp;";
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tendercheck") : document.getElementById('tendercheck');
	loDiv.innerHTML = "&nbsp;";
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tendercreditcard") : document.getElementById('tendercreditcard');
	loDiv.innerHTML = "&nbsp;";
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tenderonaccount") : document.getElementById('tenderonaccount');
	loDiv.innerHTML = "&nbsp;";
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tendercash") : document.getElementById('tendercash');
	loDiv.innerHTML = "Cash: " + FormatCurrency(gdTenderCash);
	loDiv.style.visibility = "visible";
	loDiv = ie4? eval("document.all.tenderbalance") : document.getElementById('tenderbalance');
	if (gdTenderBalance < 0) {
		loDiv.innerHTML = "Change Due: " + FormatCurrency(gdTenderBalance);
	}
	else {
		loDiv.innerHTML = "Balance Due: " + FormatCurrency(gdTenderBalance);
	}
	
	loDiv = ie4? eval("document.all.cashdiv") : document.getElementById('cashdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.checkdiv") : document.getElementById('checkdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountdiv") : document.getElementById('accountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountprintdiv") : document.getElementById('accountprintdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tipdiv") : document.getElementById('tipdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById('confirmreprint');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.allaccountdiv") : document.getElementById('allaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.collegedebitaccountdiv") : document.getElementById('collegedebitaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tenderdiv") : document.getElementById('tenderdiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function getCheck() {
	var loCheck, loDiv;

	loCheck = ie4? eval("document.all.check") : document.getElementById('check');
	loCheck.value = "";
	
	loDiv = ie4? eval("document.all.tenderdiv") : document.getElementById('tenderdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.cashdiv") : document.getElementById('cashdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountdiv") : document.getElementById('accountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountprintdiv") : document.getElementById('accountprintdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tipdiv") : document.getElementById('tipdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById('confirmreprint');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.allaccountdiv") : document.getElementById('allaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.collegedebitaccountdiv") : document.getElementById('collegedebitaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.checkdiv") : document.getElementById('checkdiv');
	loDiv.style.visibility = "visible";
		
	resetRedirect();
}

function addToCheck(psDigit) {
	var loText, lsText, loRE;
		
	loText = ie4? eval("document.all.check") : document.getElementById('check');
	loRE = /\./i;
	lsText = loText.value.replace(loRE, "");
	while (lsText.substr(0, 1) == "0") {
		lsText = lsText.substr(1);
	}
	lsText += psDigit;
	switch (lsText.length) {
		case 0:
			loText.value = lsText;
			break;
		case 1:
			loText.value = "0.0" + lsText;
			break;
		case 2:
			loText.value = "0." + lsText;
			break;
		default:
			loText.value = lsText.substr(0, (lsText.length - 2)) + "." + lsText.substr((lsText.length - 2));
			break;
	}
	
	resetRedirect();
}

function setCheckTo(psDigits) {
	var loCheck;
	
	loCheck = ie4? eval("document.all.check") : document.getElementById('check');
	loCheck.value = psDigits;
	
	setCheck();
}

function backspaceCheck() {
	var loText, lsText, loRE;
	
	loText = ie4? eval("document.all.check") : document.getElementById('check');
	loRE = /\./i;
	lsText = loText.value.replace(loRE, "");
	while (lsText.substr(0, 1) == "0") {
		lsText = lsText.substr(1);
	}
	if (lsText.length > 0) {
		lsText = lsText.substr(0, (lsText.length - 1));
		switch (lsText.length) {
			case 0:
				loText.value = lsText;
				break;
			case 1:
				loText.value = "0.0" + lsText;
				break;
			case 2:
				loText.value = "0." + lsText;
				break;
			default:
				loText.value = lsText.substr(0, (lsText.length - 2)) + "." + lsText.substr((lsText.length - 2));
				break;
		}
	}
	
	resetRedirect();
}

function cancelCheck() {
	var loCheck, loDiv;
	
	loCheck = ie4? eval("document.all.check") : document.getElementById('check');
	loCheck.value = "";
	
	loDiv = ie4? eval("document.all.cashdiv") : document.getElementById('cashdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.checkdiv") : document.getElementById('checkdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountdiv") : document.getElementById('accountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountprintdiv") : document.getElementById('accountprintdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tipdiv") : document.getElementById('tipdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById('confirmreprint');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.allaccountdiv") : document.getElementById('allaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.collegedebitaccountdiv") : document.getElementById('collegedebitaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tenderdiv") : document.getElementById('tenderdiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function setCheck() {
	var loCheck, ldCheck, loDiv;
	
	loCheck = ie4? eval("document.all.check") : document.getElementById('check');
	if (loCheck.value == "") {
		ldCheck = 0.00;
	}
	else {
		ldCheck = new Number(loCheck.value);
	}
	gnPaymentTypeID = 2;
	gdTenderCheck = Math.round(ldCheck * Math.pow(10,2))/Math.pow(10,2);
	gdTenderCash = 0.00;
	gdTenderCreditCard = 0.00;
	gdTenderOnAccount = 0.00;
	gdTenderBalance = <%=gdOrderTotal%> - gdTenderCheck;
	gdTenderBalance = Math.round(gdTenderBalance * Math.pow(10,2))/Math.pow(10,2);
	gsPaymentReference = "";
	gdTipAmount = 0.00;
	
	loDiv = ie4? eval("document.all.tendertip") : document.getElementById('tendertip');
	loDiv.innerHTML = "&nbsp;";
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tendercash") : document.getElementById('tendercash');
	loDiv.innerHTML = "&nbsp;";
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tendercreditcard") : document.getElementById('tendercreditcard');
	loDiv.innerHTML = "&nbsp;";
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tenderonaccount") : document.getElementById('tenderonaccount');
	loDiv.innerHTML = "&nbsp;";
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tendercheck") : document.getElementById('tendercheck');
	loDiv.innerHTML = "Check: " + FormatCurrency(gdTenderCheck);
	loDiv.style.visibility = "visible";
	loDiv = ie4? eval("document.all.tenderbalance") : document.getElementById('tenderbalance');
	if (gdTenderBalance < 0) {
		loDiv.innerHTML = "Change Due: " + FormatCurrency(gdTenderBalance);
	}
	else {
		loDiv.innerHTML = "Balance Due: " + FormatCurrency(gdTenderBalance);
	}
	
	loDiv = ie4? eval("document.all.cashdiv") : document.getElementById('cashdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.checkdiv") : document.getElementById('checkdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountdiv") : document.getElementById('accountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountprintdiv") : document.getElementById('accountprintdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tipdiv") : document.getElementById('tipdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById('confirmreprint');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.allaccountdiv") : document.getElementById('allaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.collegedebitaccountdiv") : document.getElementById('collegedebitaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tenderdiv") : document.getElementById('tenderdiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function getManager(pbGoAcct) {
	var loManager, loDiv;
	
	gbGoAcct = pbGoAcct;
	
	loManager = ie4? eval("document.all.manager") : document.getElementById('manager');
	loManager.value = "";
	
	loDiv = ie4? eval("document.all.tenderdiv") : document.getElementById('tenderdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.cashdiv") : document.getElementById('cashdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.checkdiv") : document.getElementById('checkdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountdiv") : document.getElementById('accountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountprintdiv") : document.getElementById('accountprintdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tipdiv") : document.getElementById('tipdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById('confirmreprint');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.allaccountdiv") : document.getElementById('allaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.collegedebitaccountdiv") : document.getElementById('collegedebitaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
	loDiv.style.visibility = "visible";
	
	loManager.focus();
	
	resetRedirect();
}

function addToManager(psDigit) {
	var loManager, lsManager;
	
	loManager = ie4? eval("document.all.manager") : document.getElementById('manager');
	lsManager = loManager.value;
	lsManager += psDigit;
	loManager.value = lsManager;
	
	resetRedirect();
}

function cancelManager() {
	var loManager, loDiv;
	
	loManager = ie4? eval("document.all.manager") : document.getElementById('manager');
	loManager.value = "";
	
	loDiv = ie4? eval("document.all.cashdiv") : document.getElementById('cashdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.checkdiv") : document.getElementById('checkdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountdiv") : document.getElementById('accountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountprintdiv") : document.getElementById('accountprintdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tipdiv") : document.getElementById('tipdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById('confirmreprint');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.allaccountdiv") : document.getElementById('allaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.collegedebitaccountdiv") : document.getElementById('collegedebitaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tenderdiv") : document.getElementById('tenderdiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function checkManagerEnterKey() {
	if (event.keyCode == 13) {
		setManager();
	}
}

function setManager() {
	var loManager, i, lbFound;
	
	lbFound = false;
	
	loManager = ie4? eval("document.all.manager") : document.getElementById('manager');
	if (loManager.value != "") {
		for (i = 0; i < ganEmpIDs.length; i++) {
			if (gasCardIDs[i] == loManager.value) {
				lbFound = true;
				i = ganEmpIDs.length;
			}
		}
		
		if (lbFound) {
			if (gbGoAcct) {
				getAccount();
			}
			else {
				getCheck();
			}
		}
		else {
			loManager.value = "";
			loManager.focus();
		}
	}
	
	resetRedirect();
}

function getAccount() {
	var loDiv;
	
	if (gnPaymentTypeID == 4) {
		loDiv = ie4? eval("document.all.tenderdiv") : document.getElementById('tenderdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.cashdiv") : document.getElementById('cashdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.checkdiv") : document.getElementById('checkdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.accountdiv") : document.getElementById('accountdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.tipdiv") : document.getElementById('tipdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById('confirmreprint');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.allaccountdiv") : document.getElementById('allaccountdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.collegedebitaccountdiv") : document.getElementById('collegedebitaccountdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.accountprintdiv") : document.getElementById('accountprintdiv');
		loDiv.style.visibility = "visible";
	}
	else {
		loDiv = ie4? eval("document.all.tenderdiv") : document.getElementById('tenderdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.cashdiv") : document.getElementById('cashdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.checkdiv") : document.getElementById('checkdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.accountprintdiv") : document.getElementById('accountprintdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.tipdiv") : document.getElementById('tipdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById('confirmreprint');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.allaccountdiv") : document.getElementById('allaccountdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.collegedebitaccountdiv") : document.getElementById('collegedebitaccountdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.accountdiv") : document.getElementById('accountdiv');
		loDiv.style.visibility = "visible";
	}
	
	resetRedirect();
}

function getCollegeDebitAccount() {
	var loDiv;
	
	if (gnPaymentTypeID == 4) {
		loDiv = ie4? eval("document.all.tenderdiv") : document.getElementById('tenderdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.cashdiv") : document.getElementById('cashdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.checkdiv") : document.getElementById('checkdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.accountdiv") : document.getElementById('accountdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.tipdiv") : document.getElementById('tipdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById('confirmreprint');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.allaccountdiv") : document.getElementById('allaccountdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.collegedebitaccountdiv") : document.getElementById('collegedebitaccountdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.accountprintdiv") : document.getElementById('accountprintdiv');
		loDiv.style.visibility = "visible";
	}
	else {
		loDiv = ie4? eval("document.all.tenderdiv") : document.getElementById('tenderdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.cashdiv") : document.getElementById('cashdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.checkdiv") : document.getElementById('checkdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.accountprintdiv") : document.getElementById('accountprintdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.tipdiv") : document.getElementById('tipdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById('confirmreprint');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.allaccountdiv") : document.getElementById('allaccountdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.accountdiv") : document.getElementById('accountdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.collegedebitaccountdiv") : document.getElementById('collegedebitaccountdiv');
		loDiv.style.visibility = "visible";
	}
	
	resetRedirect();
}

function cancelAccount() {
	var loDiv;
	
	gsPaymentReference = "";
	
	loDiv = ie4? eval("document.all.cashdiv") : document.getElementById('cashdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.checkdiv") : document.getElementById('checkdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountdiv") : document.getElementById('accountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountprintdiv") : document.getElementById('accountprintdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tipdiv") : document.getElementById('tipdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById('confirmreprint');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.allaccountdiv") : document.getElementById('allaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.collegedebitaccountdiv") : document.getElementById('collegedebitaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tenderdiv") : document.getElementById('tenderdiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function viewAllAccounts() {
	var loDiv;
	
	loDiv = ie4? eval("document.all.tenderdiv") : document.getElementById('tenderdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.cashdiv") : document.getElementById('cashdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.checkdiv") : document.getElementById('checkdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountprintdiv") : document.getElementById('accountprintdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tipdiv") : document.getElementById('tipdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById('confirmreprint');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountdiv") : document.getElementById('accountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.collegedebitaccountdiv") : document.getElementById('collegedebitaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.allaccountdiv") : document.getElementById('allaccountdiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function setAccount(pnAccountID) {
	var loCheck, ldCheck, loDiv;
	
	gsPaymentReference = pnAccountID.toString();
	
	loDiv = ie4? eval("document.all.cashdiv") : document.getElementById('cashdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.checkdiv") : document.getElementById('checkdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountdiv") : document.getElementById('accountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tenderdiv") : document.getElementById('tenderdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tipdiv") : document.getElementById('tipdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById('confirmreprint');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.allaccountdiv") : document.getElementById('allaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.collegedebitaccountdiv") : document.getElementById('collegedebitaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountprintdiv") : document.getElementById('accountprintdiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function cancelAccountPrint() {
	var loDiv;
	
	loDiv = ie4? eval("document.all.cashdiv") : document.getElementById('cashdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.checkdiv") : document.getElementById('checkdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tenderdiv") : document.getElementById('tenderdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountprintdiv") : document.getElementById('accountprintdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tipdiv") : document.getElementById('tipdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById('confirmreprint');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.allaccountdiv") : document.getElementById('allaccountdiv');
	loDiv.style.visibility = "hidden";
<%
If Not (Session("SecurityID") > 1 And Session("Swipe")) And ganCollegeDebitAccountIDs(0) <> 0 Then
%>
	loDiv = ie4? eval("document.all.accountdiv") : document.getElementById('accountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.collegedebitaccountdiv") : document.getElementById('collegedebitaccountdiv');
	loDiv.style.visibility = "visible";
<%
Else
%>
	loDiv = ie4? eval("document.all.collegedebitaccountdiv") : document.getElementById('collegedebitaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountdiv") : document.getElementById('accountdiv');
	loDiv.style.visibility = "visible";
<%
End If
%>
	
	resetRedirect();
}

function setAccountPrint(pbPrintNow) {
	var loCheck, ldCheck, loDiv;
	
	gnPaymentTypeID = 4;
	gdTenderCheck = 0.00;
	gdTenderCash = 0.00;
	gdTenderCreditCard = 0.00;
	if (pbPrintNow) {
		gdTenderOnAccount = <%=gdOrderTotal%>;
		gdTenderBalance = 0.00;
	}
	else {
		gdTenderOnAccount = 0.00;
		gdTenderBalance = <%=gdOrderTotal%>;
		gsTip = 0.00;
	}
	
	loDiv = ie4? eval("document.all.tendertip") : document.getElementById('tendertip');
	loDiv.innerHTML = "&nbsp;";
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tendercash") : document.getElementById('tendercash');
	loDiv.innerHTML = "&nbsp;";
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tendercreditcard") : document.getElementById('tendercreditcard');
	loDiv.innerHTML = "&nbsp;";
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tendercheck") : document.getElementById('tendercheck');
	loDiv.innerHTML = "&nbsp;";
	loDiv.style.visibility = "hidden";
	
	if (pbPrintNow) {
		getTip();
	}
	else {
		loDiv = ie4? eval("document.all.tenderonaccount") : document.getElementById('tenderonaccount');
		loDiv.innerHTML = "On Account " + gsPaymentReference + ": " + FormatCurrency(gdTenderOnAccount + gdTipAmount);
		loDiv.style.visibility = "visible";
		loDiv = ie4? eval("document.all.tenderbalance") : document.getElementById('tenderbalance');
		if (gdTenderBalance < 0) {
			loDiv.innerHTML = "Change Due: " + FormatCurrency(gdTenderBalance);
		}
		else {
			loDiv.innerHTML = "Balance Due: " + FormatCurrency(gdTenderBalance);
		}
		
		loDiv = ie4? eval("document.all.cashdiv") : document.getElementById('cashdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.checkdiv") : document.getElementById('checkdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.accountdiv") : document.getElementById('accountdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.accountprintdiv") : document.getElementById('accountprintdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.tipdiv") : document.getElementById('tipdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById('confirmreprint');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.allaccountdiv") : document.getElementById('allaccountdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.collegedebitaccountdiv") : document.getElementById('collegedebitaccountdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.tenderdiv") : document.getElementById('tenderdiv');
		loDiv.style.visibility = "visible";
		
		resetRedirect();
	}
}

function getTip() {
	var loTip, loDiv;

	loTip = ie4? eval("document.all.tip") : document.getElementById('tip');
	loTip.value = "";
	
	loDiv = ie4? eval("document.all.tenderdiv") : document.getElementById('tenderdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.cashdiv") : document.getElementById('cashdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountdiv") : document.getElementById('accountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountprintdiv") : document.getElementById('accountprintdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.checkdiv") : document.getElementById('checkdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById('confirmreprint');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.allaccountdiv") : document.getElementById('allaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.collegedebitaccountdiv") : document.getElementById('collegedebitaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tipdiv") : document.getElementById('tipdiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function addToTip(psDigit) {
	var loText, lsText, loRE;
	
	loText = ie4? eval("document.all.tip") : document.getElementById('tip');
	loRE = /\./i;
	lsText = loText.value.replace(loRE, "");
	while (lsText.substr(0, 1) == "0") {
		lsText = lsText.substr(1);
	}
	lsText += psDigit;
	switch (lsText.length) {
		case 0:
			loText.value = lsText;
			break;
		case 1:
			loText.value = "0.0" + lsText;
			break;
		case 2:
			loText.value = "0." + lsText;
			break;
		default:
			loText.value = lsText.substr(0, (lsText.length - 2)) + "." + lsText.substr((lsText.length - 2));
			break;
	}
	
	resetRedirect();
}

function setTipTo(psDigits) {
	var loTip;
	
	loTip = ie4? eval("document.all.tip") : document.getElementById('tip');
	loTip.value = psDigits;
	
	setTip();
}

function backspaceTip() {
	var loText, lsText, loRE;
	
	loText = ie4? eval("document.all.tip") : document.getElementById('tip');
	loRE = /\./i;
	lsText = loText.value.replace(loRE, "");
	while (lsText.substr(0, 1) == "0") {
		lsText = lsText.substr(1);
	}
	if (lsText.length > 0) {
		lsText = lsText.substr(0, (lsText.length - 1));
		switch (lsText.length) {
			case 0:
				loText.value = lsText;
				break;
			case 1:
				loText.value = "0.0" + lsText;
				break;
			case 2:
				loText.value = "0." + lsText;
				break;
			default:
				loText.value = lsText.substr(0, (lsText.length - 2)) + "." + lsText.substr((lsText.length - 2));
				break;
		}
	}
	
	resetRedirect();
}

function cancelTip() {
	var loTip, loDiv;
	
	loTip = ie4? eval("document.all.tip") : document.getElementById('tip');
	loTip.value = "";
	
	loDiv = ie4? eval("document.all.cashdiv") : document.getElementById('cashdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.checkdiv") : document.getElementById('checkdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountdiv") : document.getElementById('accountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountprintdiv") : document.getElementById('accountprintdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tipdiv") : document.getElementById('tipdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById('confirmreprint');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.allaccountdiv") : document.getElementById('allaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.collegedebitaccountdiv") : document.getElementById('collegedebitaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tenderdiv") : document.getElementById('tenderdiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function setTip() {
	var loTip, ldTip, loDiv;
	
	loTip = ie4? eval("document.all.tip") : document.getElementById('tip');
	if (loTip.value == "") {
		ldTip = 0.00;
	}
	else {
		ldTip = new Number(loTip.value);
	}
	gdTipAmount = Math.round(ldTip * Math.pow(10,2))/Math.pow(10,2);
	
	loDiv = ie4? eval("document.all.tendertip") : document.getElementById('tendertip');
	if (gdTipAmount == 0) {
		loDiv.innerHTML = "&nbsp;";
	}
	else {
		loDiv.innerHTML = "Tip: " + FormatCurrency(gdTipAmount);
		loDiv.style.visibility = "visible";
	}
	
	if (gnPaymentTypeID == 4) {
		loDiv = ie4? eval("document.all.tenderonaccount") : document.getElementById('tenderonaccount');
		loDiv.innerHTML = "On Account " + gsPaymentReference + ": " + FormatCurrency(gdTenderOnAccount + gdTipAmount);
		loDiv.style.visibility = "visible";
	}
	else {
		loDiv = ie4? eval("document.all.tendercreditcard") : document.getElementById('tendercreditcard');
		loDiv.innerHTML = "Credit Card: " + FormatCurrency(gdTenderCreditCard + gdTipAmount);
		loDiv.style.visibility = "visible";
	}
	loDiv = ie4? eval("document.all.tenderbalance") : document.getElementById('tenderbalance');
	if (gdTenderBalance < 0) {
		loDiv.innerHTML = "Change Due: " + FormatCurrency(gdTenderBalance);
	}
	else {
		loDiv.innerHTML = "Balance Due: " + FormatCurrency(gdTenderBalance);
	}
	
	loDiv = ie4? eval("document.all.cashdiv") : document.getElementById('cashdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.checkdiv") : document.getElementById('checkdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountdiv") : document.getElementById('accountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountprintdiv") : document.getElementById('accountprintdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tipdiv") : document.getElementById('tipdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById('confirmreprint');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.allaccountdiv") : document.getElementById('allaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.collegedebitaccountdiv") : document.getElementById('collegedebitaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tenderdiv") : document.getElementById('tenderdiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function gotoConfirmReprint() {
	var loDiv;
	
	loDiv = ie4? eval("document.all.cashdiv") : document.getElementById('cashdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.checkdiv") : document.getElementById('checkdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountdiv") : document.getElementById('accountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountprintdiv") : document.getElementById('accountprintdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tipdiv") : document.getElementById('tipdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tenderdiv") : document.getElementById('tenderdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.allaccountdiv") : document.getElementById('allaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.collegedebitaccountdiv") : document.getElementById('collegedebitaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById('confirmreprint');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function cancelConfirmReprint() {
	var loDiv;
	
	loDiv = ie4? eval("document.all.cashdiv") : document.getElementById('cashdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.checkdiv") : document.getElementById('checkdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountdiv") : document.getElementById('accountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.accountprintdiv") : document.getElementById('accountprintdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tipdiv") : document.getElementById('tipdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById('confirmreprint');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.allaccountdiv") : document.getElementById('allaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.collegedebitaccountdiv") : document.getElementById('collegedebitaccountdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tenderdiv") : document.getElementById('tenderdiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function goChange() {
	var lsLocation;
	
	if (gdTenderBalance <= 0) {
		lsLocation = "change.asp?o=" + gnOrderID.toString() + "&v=" + gnPaymentTypeID.toString() + "&s="
		switch (gnPaymentTypeID) {
			case 1:
				lsLocation = lsLocation + gdTenderCash.toString();
				break;
			case 2:
				lsLocation = lsLocation + gdTenderCheck.toString();
				break;
			case 3:
				lsLocation = lsLocation + gdTenderCreditCard.toString();
				break;
			case 4:
				lsLocation = lsLocation + gdTenderOnAccount.toString();
				break;
		}
		lsLocation = lsLocation + "&j=" + gdTipAmount.toString();
		if (gsPaymentReference.length > 0) {
			lsLocation = lsLocation + "&r=" + encodeURIComponent(gsPaymentReference);
		}
		if (!gbDoPrint) {
			lsLocation = lsLocation + "&q=yes"
		}
		
		window.location = lsLocation;
	}
	else {
		if ((gnPaymentTypeID != 0) && (gdOrderTotal == gdTenderBalance)) {
			lsLocation = "neworder.asp?o=" + gnOrderID.toString() + "&a=<%=gnAddressID%>&c=<%=gnCustomerID%>&v=" + gnPaymentTypeID.toString();
			if (gsPaymentReference.length > 0) {
				lsLocation = lsLocation + "&r=" + encodeURIComponent(gsPaymentReference);
			}
			if (!gbDoPrint) {
				lsLocation = lsLocation + "&q=yes"
			}
			
			window.location = lsLocation;
		}
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
					<table cellpadding="0" cellspacing="0" width="1010" height="723">
						<tr>
							<td valign="top" width="680">
								<div id="content" style="position: relative; width: 680px; height: 723px; overflow: auto;">
									<div id="tenderdiv" style="position: absolute; left: 0px; top: 0px; width: 680px; height: 723px; background-color: #fbf3c5;">
										<p align="center"><strong>Tender Payment</strong></p>
										<p align="center"><strong>Order #&nbsp;<%=gnOrderID%></strong></p>
										<div style="height: 380px; padding: 10px;">
										<%=gsCustomerName%><br/>
<%
If gnCustomerID = 1 Then
%>
										&nbsp;<br/>
<%
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
										<div style="position: relative;">
										<div style="position: absolute; top: 0px; left: 0px;"><button style="width: 337px;" onclick="getCash();">Cash</button></div>
<%
If IsStoreChecksOK(gnStoreID) Then
	If gnCustomerID > 1 Then
		If IsCustomerCheckOK(gnCustomerID) Then
			If DateDiff("d", GetLastCustomerOrderDate(gnCustomerID), Now()) <= gnCheckAcceptMaxDaysSinceOrdered Then

%>
										<div style="position: absolute; top: 0px; left: 341px;"><button style="width: 337px;" onclick="getCheck();">Check</button></div>
<%
			Else
				If Session("SecurityID") > 1 And ((Left(Request.ServerVariables("remote_addr"), 9) = "10.0.254." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.1." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.2.") OR Session("Swipe")) Then
%>
										<div style="position: absolute; top: 0px; left: 341px;"><button style="width: 337px; background-color: #C0C0C0;" onclick="getCheck();">NO CHECKS<br/>Customer Has Not Order Within <%=gnCheckAcceptMaxDaysSinceOrdered%> Days<br/>Press For Manager Override</button></div>
<%
				Else
%>
										<div style="position: absolute; top: 0px; left: 341px;"><button style="width: 337px; background-color: #C0C0C0;" onclick="getManager(false);">NO CHECKS<br/>Customer Has Not Order Within <%=gnCheckAcceptMaxDaysSinceOrdered%> Days<br/>Press For Manager Override</button></div>
<%
				End If
			End If
		Else
			If Session("SecurityID") > 1 And ((Left(Request.ServerVariables("remote_addr"), 9) = "10.0.254." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.1." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.2.") OR Session("Swipe")) Then
%>
										<div style="position: absolute; top: 0px; left: 341px;"><button style="width: 337px; background-color: #C0C0C0;" onclick="getCheck();">NO CHECKS<br/>CUSTOMER IS FLAGGED</button></div>
<%
			Else
%>
										<div style="position: absolute; top: 0px; left: 341px;"><button style="width: 337px; background-color: #C0C0C0;" onclick="getManager(false);">NO CHECKS<br/>CUSTOMER IS FLAGGED</button></div>
<%
			End If
		End If
	Else
		If Session("SecurityID") > 1 And ((Left(Request.ServerVariables("remote_addr"), 9) = "10.0.254." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.1." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.2.") OR Session("Swipe")) Then
%>
										<div style="position: absolute; top: 0px; left: 341px;"><button style="width: 337px; background-color: #C0C0C0;" onclick="getCheck();">Checks Not Accepted Without Customer</button></div>
<%
		Else
%>
										<div style="position: absolute; top: 0px; left: 341px;"><button style="width: 337px; background-color: #C0C0C0;" onclick="getManager(false);">Checks Not Accepted Without Customer</button></div>
<%
		End If
	End If
Else
%>
										<div style="position: absolute; top: 0px; left: 341px;"><button style="width: 337px; background-color: #C0C0C0;" onclick="getManager(false);">This Store Does Not Accept Checks</button></div>
<%
End If
%>
										<div style="position: absolute; top: 79px; left: 0px;"><button style="width: 337px;" onclick="window.location = 'creditcard.asp?o=<%=gnOrderID%>'">Credit Card</button></div>
<%
If gnCustomerID = 1 Or ganAllAccountIDs(0) = 0 Then
	If ganCollegeDebitAccountIDs(0) = 0 Then
%>
										<div style="position: absolute; top: 79px; left: 341px;"><button style="width: 337px; background-color: #C0C0C0;">&nbsp;</button></div>
<%
	Else
%>
										<div style="position: absolute; top: 79px; left: 341px;"><button style="width: 337px;" onclick="getCollegeDebitAccount();">On Account / College Debit</button></div>
<%
	End If
Else
	If Session("SecurityID") > 1 And Session("Swipe") Then
%>
										<div style="position: absolute; top: 79px; left: 341px;"><button style="width: 337px;" onclick="getAccount();">On Account / College Debit</button></div>
<%
	Else
		If ganCollegeDebitAccountIDs(0) = 0 Then
			If Session("SecurityID") > 1 And ((Left(Request.ServerVariables("remote_addr"), 9) = "10.0.254." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.1." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.2.") OR Session("Swipe")) Then
%>
										<div style="position: absolute; top: 79px; left: 341px;"><button style="width: 337px;" onclick="getAccount();">On Account / College Debit</button></div>
<%
			Else
%>
										<div style="position: absolute; top: 79px; left: 341px;"><button style="width: 337px;" onclick="getManager(true);">On Account / College Debit</button></div>
<%
			End If
		Else
%>
										<div style="position: absolute; top: 79px; left: 341px;"><button style="width: 337px;" onclick="getCollegeDebitAccount();">On Account / College Debit</button></div>
<%
		End If
	End If
End If

If InStr(LCase(Request.ServerVariables("HTTP_REFERER")), "ticket.asp") > 0 Then
%>
										<div style="position: absolute; top: 158px; left: 0px;"><button style="width: 337px;" onclick="window.location = '../ticket.asp?OrderID=<%=gnOrderID%>'">Cancel</button></div>
<%
Else
%>
										<div style="position: absolute; top: 158px; left: 0px;"><button style="width: 337px;" onclick="window.location = 'unitselect.asp?o=<%=gnOrderID%>'">Cancel</button></div>
<%
End If

If Session("NewOrder") Then
%>
										<div style="position: absolute; top: 158px; left: 341px;"><button style="width: 337px;" onclick="goChange();">Done</button></div>
<%
Else
	If Session("OrderEdited") Then
%>
										<div style="position: absolute; top: 158px; left: 341px;"><button style="width: 337px;" onclick="gotoConfirmReprint();">Done</button></div>
<%
	Else
%>
										<div style="position: absolute; top: 158px; left: 341px;"><button style="width: 337px;" onclick="goChange();">Done</button></div>
<%
	End If
End If
%>
										</div>
									</div>
									<div id="cashdiv" style="position: absolute; left: 0px; top: 0px; width: 680px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
										<table align="center" cellpadding="0" cellspacing="0">
											<tr>
												<td colspan="3"><div align="center">
													<strong>ENTER AMOUNT OF CASH</strong></div></td>
											</tr>
											<tr>
												<td colspan="3"><div align="center">
													<input type="text" id="cash" style="width: 200px" autocomplete="off" onkeydown="disableEnterKey();" /></div></td>
											</tr>
											<tr>
												<td><button onclick="addToCash('1')">1</button></td>
												<td><button onclick="addToCash('2')">2</button></td>
												<td><button onclick="addToCash('3')">3</button></td>
											</tr>
											<tr>
												<td><button onclick="addToCash('4')">4</button></td>
												<td><button onclick="addToCash('5')">5</button></td>
												<td><button onclick="addToCash('6')">6</button></td>
											</tr>
											<tr>
												<td><button onclick="addToCash('7')">7</button></td>
												<td><button onclick="addToCash('8')">8</button></td>
												<td><button onclick="addToCash('9')">9</button></td>
											</tr>
											<tr>
												<td><button onclick="cancelCash()">Cancel</button></td>
												<td><button onclick="addToCash('0')">0</button></td>
												<td><button onclick="backspaceCash()">Bksp</button></td>
											</tr>
											<tr>
												<td colspan="3"><button style="width: 235px;" onclick="setCash()">OK</button></td>
											</tr>
											<tr>
												<td colspan="3">&nbsp;</td>
											</tr>
											<tr>
												<td><button onclick="setCashTo('<%=Replace(FormatCurrency(gdOrderTotal), "$", "")%>')"><%=FormatCurrency(gdOrderTotal)%></button></td>
												<td><button onclick="setCashTo('1.00')">$1</button></td>
												<td><button onclick="setCashTo('5.00')">$5</button></td>
											</tr>
											<tr>
												<td><button onclick="setCashTo('10.00')">$10</button></td>
												<td><button onclick="setCashTo('15.00')">$15</button></td>
												<td><button onclick="setCashTo('20.00')">$20</button></td>
											</tr>
											<tr>
												<td><button onclick="setCashTo('25.00')">$25</button></td>
												<td><button onclick="setCashTo('30.00')">$30</button></td>
												<td><button onclick="setCashTo('50.00')">$50</button></td>
											</tr>
										</table>
									</div>
									<div id="checkdiv" style="position: absolute; left: 0px; top: 0px; width: 680px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
										<table align="center" cellpadding="0" cellspacing="0">
											<tr>
												<td colspan="3"><div align="center">
													<strong>ENTER AMOUNT OF CHECK</strong></div></td>
											</tr>
											<tr>
												<td colspan="3"><div align="center">
													<input type="text" id="check" style="width: 200px" autocomplete="off" onkeydown="disableEnterKey();" /></div></td>
											</tr>
											<tr>
												<td><button onclick="addToCheck('1')">1</button></td>
												<td><button onclick="addToCheck('2')">2</button></td>
												<td><button onclick="addToCheck('3')">3</button></td>
											</tr>
											<tr>
												<td><button onclick="addToCheck('4')">4</button></td>
												<td><button onclick="addToCheck('5')">5</button></td>
												<td><button onclick="addToCheck('6')">6</button></td>
											</tr>
											<tr>
												<td><button onclick="addToCheck('7')">7</button></td>
												<td><button onclick="addToCheck('8')">8</button></td>
												<td><button onclick="addToCheck('9')">9</button></td>
											</tr>
											<tr>
												<td><button onclick="cancelCheck()">Cancel</button></td>
												<td><button onclick="addToCheck('0')">0</button></td>
												<td><button onclick="backspaceCheck()">Bksp</button></td>
											</tr>
											<tr>
												<td colspan="3"><button style="width: 235px;" onclick="setCheck()">OK</button></td>
											</tr>
											<tr>
												<td colspan="3">&nbsp;</td>
											</tr>
											<tr>
												<td><button onclick="setCheckTo('<%=Replace(FormatCurrency(gdOrderTotal), "$", "")%>')"><%=FormatCurrency(gdOrderTotal)%></button></td>
												<td><button onclick="setCheckTo('1.00')">$1</button></td>
												<td><button onclick="setCheckTo('5.00')">$5</button></td>
											</tr>
											<tr>
												<td><button onclick="setCheckTo('10.00')">$10</button></td>
												<td><button onclick="setCheckTo('15.00')">$15</button></td>
												<td><button onclick="setCheckTo('20.00')">$20</button></td>
											</tr>
											<tr>
												<td><button onclick="setCheckTo('25.00')">$25</button></td>
												<td><button onclick="setCheckTo('30.00')">$30</button></td>
												<td><button onclick="setCheckTo('50.00')">$50</button></td>
											</tr>
										</table>
									</div>
									<div id="managerdiv" style="position: absolute; left: 0px; top: 0px; width: 680px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
										<table align="center" cellpadding="0" cellspacing="0">
											<tr>
												<td colspan="3"><div align="center">
													<strong>MANAGER OVERRIDE REQUIRED</strong></div></td>
											</tr>
											<tr>
												<td colspan="3"><div align="center">
													<input type="password" id="manager" style="width: 200px" autocomplete="off" onkeyup="checkManagerEnterKey();" /></div></td>
											</tr>
											<tr>
												<td><button onclick="cancelManager()">Cancel</button></td>
												<td>&nbsp;</td>
												<td align="right"><button onclick="setManager()">OK</button></td>
											</tr>
										</table>
									</div>
									<div id="accountdiv" style="position: absolute; left: 0px; top: 0px; width: 680px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
										<p align="center"><strong>SELECT ACCOUNT TO PLACE ON</strong></p>
										<div style="position: relative; width: 670px; height: 558px;">
<%
If ganAccountIDs(0) <> 0 Then
	lnTop = 0
	lnLeft = 0
	For i = 0 To UBound(ganAccountIDs)
		If i Mod 13 = 0 Then
			If i > 0 And UBound(ganAccountIDs) > 13 Then
%>
												<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 337px; height: 75px;" onclick="toggleDivs('accountdiv<%=Int(i/13)-1%>', 'accountdiv<%=Int(i/13)%>')">(Next)</button></div>
											</div>
											<div id="accountdiv<%=Int(i/13)%>" style="position: absolute; top: 0px; left: 0px; width: 670px; visibility: hidden;">
<%
			Else
				If i = 0 Then
%>
											<div id="accountdiv<%=Int(i/13)%>" style="position: absolute; top: 0px; left: 0px; width: 670px;">
<%
				End If
			End If
			
			lnTop = 0
			lnLeft = 0
		End If
		
		If gabAccountOnHolds(i) Then
%>
												<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 337px; height: 75px; background-color: #C0C0C0;" onclick="">Account #<%=ganAccountIDs(i)%><br /><%=gasAccountNames(i)%><br />** ON HOLD **</button></div>
<%
		Else
%>
												<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 337px; height: 75px;" onclick="setAccount(<%=ganAccountIDs(i)%>);">Account #<%=ganAccountIDs(i)%><br /><%=gasAccountNames(i)%></button></div>
<%
		End If
		
		If lnLeft = 338 Then
			lnTop = lnTop + 76
			lnLeft = 0
		Else
			lnLeft = lnLeft + 338
		End If
	Next
	
	' Add hidden buttons here
	If ((UBound(ganAccountIDs) + 1) Mod 13) > 0 And UBound(ganAccountIDs) <> 13 Then
		For i = ((UBound(ganAccountIDs) + 1) Mod 13) To 12
%>
												<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 337px; height: 75px; background-color: #C0C0C0;">&nbsp;</button></div>
<%
			If lnLeft = 338 Then
				lnTop = lnTop + 76
				lnLeft = 0
			Else
				lnLeft = lnLeft + 338
			End If
		Next
	End If
	
	If UBound(ganAccountIDs) > 13 Then
%>
												<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 337px; height: 75px;" onclick="toggleDivs('accountdiv<%=Int(UBound(ganAccountIDs)/13)%>', 'accountdiv0')">(Next)</button></div>
<%
	Else
		If UBound(ganAccountIDs) <> 13 Then
%>
												<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 337px; height: 75px; background-color: #C0C0C0;">&nbsp;</button></div>
<%
		End If
	End If
%>
											</div>
<%
End If
%>
										</div>
										<div style="position: relative;">
										<div style="position: absolute; top: 0px; left: 0px;"><button style="width: 337px;" onclick="viewAllAccounts();">View All Accounts</button></div>
										<div style="position: absolute; top: 0px; left: 338px;"><button style="width: 337px;" onclick="cancelAccount();">Cancel</button></div>
										</div>
									</div>
									<div id="allaccountdiv" style="position: absolute; left: 0px; top: 0px; width: 680px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
										<p align="center"><strong>SELECT ACCOUNT TO PLACE ON</strong></p>
										<div style="position: relative; width: 670px; height: 558px;">
<%
If ganAllAccountIDs(0) <> 0 Then
	lnTop = 0
	lnLeft = 0
	For i = 0 To UBound(ganAllAccountIDs)
		If i Mod 13 = 0 Then
			If i > 0 And UBound(ganAllAccountIDs) > 13 Then
%>
												<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 337px; height: 75px;" onclick="toggleDivs('allaccountdiv<%=Int(i/13)-1%>', 'allaccountdiv<%=Int(i/13)%>')">(Next)</button></div>
											</div>
											<div id="allaccountdiv<%=Int(i/13)%>" style="position: absolute; top: 0px; left: 0px; width: 670px; visibility: hidden;">
<%
			Else
				If i = 0 Then
%>
											<div id="allaccountdiv<%=Int(i/13)%>" style="position: absolute; top: 0px; left: 0px; width: 670px;">
<%
				End If
			End If
			
			lnTop = 0
			lnLeft = 0
		End If
		
		If gabAllAccountOnHolds(i) Then
%>
												<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 337px; height: 75px; background-color: #C0C0C0;" onclick="">Account #<%=ganAllAccountIDs(i)%><br /><%=gasAllAccountNames(i)%><br />** ON HOLD **</button></div>
<%
		Else
%>
												<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 337px; height: 75px;" onclick="setAccount(<%=ganAllAccountIDs(i)%>);">Account #<%=ganAllAccountIDs(i)%><br /><%=gasAllAccountNames(i)%></button></div>
<%
		End If
		
		If lnLeft = 338 Then
			lnTop = lnTop + 76
			lnLeft = 0
		Else
			lnLeft = lnLeft + 338
		End If
	Next
	
	' Add hidden buttons here
	If ((UBound(ganAllAccountIDs) + 1) Mod 13) > 0 And UBound(ganAllAccountIDs) <> 13 Then
		For i = ((UBound(ganAllAccountIDs) + 1) Mod 13) To 12
%>
												<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 337px; height: 75px; background-color: #C0C0C0;">&nbsp;</button></div>
<%
			If lnLeft = 338 Then
				lnTop = lnTop + 76
				lnLeft = 0
			Else
				lnLeft = lnLeft + 338
			End If
		Next
	End If
	
	If UBound(ganAllAccountIDs) > 13 Then
%>
												<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 337px; height: 75px;" onclick="toggleDivs('allaccountdiv<%=Int(UBound(ganAllAccountIDs)/13)%>', 'allaccountdiv0')">(Next)</button></div>
<%
	Else
		If UBound(ganAllAccountIDs) <> 13 Then
%>
												<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 337px; height: 75px; background-color: #C0C0C0;">&nbsp;</button></div>
<%
		End If
	End If
%>
											</div>
<%
End If
%>
										</div>
										<div style="position: relative;">
										<div style="position: absolute; top: 0px; left: 0px;"><button style="width: 337px;" onclick="getAccount();">View Store Accounts</button></div>
										<div style="position: absolute; top: 0px; left: 338px;"><button style="width: 337px;" onclick="cancelAccount();">Cancel</button></div>
										</div>
									</div>
									<div id="collegedebitaccountdiv" style="position: absolute; left: 0px; top: 0px; width: 680px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
										<p align="center"><strong>SELECT ACCOUNT TO PLACE ON</strong></p>
										<div style="position: relative; width: 670px; height: 558px;">
<%
If ganCollegeDebitAccountIDs(0) <> 0 Then
	lnTop = 0
	lnLeft = 0
	For i = 0 To UBound(ganCollegeDebitAccountIDs)
		If i Mod 13 = 0 Then
			If i > 0 And UBound(ganCollegeDebitAccountIDs) > 13 Then
%>
												<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 337px; height: 75px;" onclick="toggleDivs('collegedebitaccountdiv<%=Int(i/13)-1%>', 'collegedebitaccountdiv<%=Int(i/13)%>')">(Next)</button></div>
											</div>
											<div id="collegedebitaccountdiv<%=Int(i/13)%>" style="position: absolute; top: 0px; left: 0px; width: 670px; visibility: hidden;">
<%
			Else
				If i = 0 Then
%>
											<div id="collegedebitaccountdiv<%=Int(i/13)%>" style="position: absolute; top: 0px; left: 0px; width: 670px;">
<%
				End If
			End If
			
			lnTop = 0
			lnLeft = 0
		End If
		
		If gabCollegeDebitAccountOnHolds(i) Then
%>
												<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 337px; height: 75px; background-color: #C0C0C0;" onclick="">Account #<%=ganCollegeDebitAccountIDs(i)%><br /><%=gasCollegeDebitAccountNames(i)%><br />** ON HOLD **</button></div>
<%
		Else
%>
												<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 337px; height: 75px;" onclick="setAccount(<%=ganCollegeDebitAccountIDs(i)%>);">Account #<%=ganCollegeDebitAccountIDs(i)%><br /><%=gasCollegeDebitAccountNames(i)%></button></div>
<%
		End If
		
		If lnLeft = 338 Then
			lnTop = lnTop + 76
			lnLeft = 0
		Else
			lnLeft = lnLeft + 338
		End If
	Next
	
	' Add hidden buttons here
	If ((UBound(ganCollegeDebitAccountIDs) + 1) Mod 13) > 0 And UBound(ganCollegeDebitAccountIDs) <> 13 Then
		For i = ((UBound(ganCollegeDebitAccountIDs) + 1) Mod 13) To 12
%>
												<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 337px; height: 75px; background-color: #C0C0C0;">&nbsp;</button></div>
<%
			If lnLeft = 338 Then
				lnTop = lnTop + 76
				lnLeft = 0
			Else
				lnLeft = lnLeft + 338
			End If
		Next
	End If
	
	If UBound(ganCollegeDebitAccountIDs) > 13 Then
%>
												<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 337px; height: 75px;" onclick="toggleDivs('collegedebitaccountdiv<%=Int(UBound(ganCollegeDebitAccountIDs)/13)%>', 'collegedebitaccountdiv0')">(Next)</button></div>
<%
	Else
		If UBound(ganCollegeDebitAccountIDs) <> 13 Then
%>
												<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 337px; height: 75px; background-color: #C0C0C0;">&nbsp;</button></div>
<%
		End If
	End If
%>
											</div>
<%
End If
%>
										</div>
										<div style="position: relative;">
										<div style="position: absolute; top: 0px; left: 0px;">
<%
If gnCustomerID > 1 Then
	If Session("SecurityID") > 1 And ((Left(Request.ServerVariables("remote_addr"), 9) = "10.0.254." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.1." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.2.") OR Session("Swipe")) Then
%>
											<button style="width: 337px;" onclick="getAccount();">Other Account</button>
<%
	Else
%>
											<button style="width: 337px;" onclick="getManager(true);">Other Account</button>
<%
	End If
End If
%>
										</div>
										<div style="position: absolute; top: 0px; left: 338px;"><button style="width: 337px;" onclick="cancelAccount();">Cancel</button></div>
										</div>
									</div>
									<div id="accountprintdiv" style="position: absolute; left: 0px; top: 0px; width: 680px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
										<div align="center"><strong>PLACE ON ACCOUNT</strong></div><br/>
										<button style="width: 680px;" onclick="setAccountPrint(true);">Print Signature Copy Now</button>
										<button style="width: 680px;" onclick="setAccountPrint(false);">Print Signature Copy Later</button>
										<button style="width: 680px;" onclick="cancelAccountPrint();">Change Account</button>
									</div>
									<div id="tipdiv" style="position: absolute; left: 0px; top: 0px; width: 680px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
										<table align="center" cellpadding="0" cellspacing="0">
											<tr>
												<td colspan="3"><div align="center">
													<strong>ENTER TIP AMOUNT</strong></div></td>
											</tr>
											<tr>
												<td colspan="3"><div align="center">
													<input type="text" id="tip" style="width: 200px" autocomplete="off" onkeydown="disableEnterKey();" /></div></td>
											</tr>
											<tr>
												<td><button onclick="addToTip('1')">1</button></td>
												<td><button onclick="addToTip('2')">2</button></td>
												<td><button onclick="addToTip('3')">3</button></td>
											</tr>
											<tr>
												<td><button onclick="addToTip('4')">4</button></td>
												<td><button onclick="addToTip('5')">5</button></td>
												<td><button onclick="addToTip('6')">6</button></td>
											</tr>
											<tr>
												<td><button onclick="addToTip('7')">7</button></td>
												<td><button onclick="addToTip('8')">8</button></td>
												<td><button onclick="addToTip('9')">9</button></td>
											</tr>
											<tr>
												<td>&nbsp;</td>
												<td><button onclick="addToTip('0')">0</button></td>
												<td><button onclick="backspaceTip()">Bksp</button></td>
											</tr>
											<tr>
												<td colspan="3"><button style="width: 235px;" onclick="setTip()">OK</button></td>
											</tr>
										</table>
									</div>
									<div id="confirmreprint" style="position: absolute; top: 0px; left: 0px; width: 680px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
										<p align="center"><strong>Do you want to reprint this order?</strong><br/><br/>
										<button onclick="goChange();">Reprint</button>&nbsp;&nbsp;&nbsp;&nbsp;<button onclick="gbDoPrint = false; goChange();">No Reprint</button>&nbsp;&nbsp;&nbsp;&nbsp;<button onclick="cancelConfirmReprint();">Cancel</button>
										</p>
									</div>
								</div>
							</td>
							<td align="right" valign="top" width="330">
								<div style="position: relative; width: 320px; height: 539px; text-align: left; background-color: #FFFFFF;">
<%
If ganOrderLineIDs(0) <> 0 Then
	For i = 0 To UBound(ganOrderLineIDs)
		If i Mod 3 = 0 Then
			If i > 0 Then
%>
										<button style="width: 320px; color: #FFFFFF; background-color: #FF0000;" onclick="toggleDivs('itemdiv<%=Int(i/3)-1%>', 'itemdiv<%=Int(i/3)%>')">Page <%=Int(i/3)%> of <%=Int(UBound(ganOrderLineIDs)/3)+1%><br/>(Next)</button>
									</div>
									<div id="itemdiv<%=Int(i/3)%>" style="position: absolute; top: 0px; visibility: hidden;">
<%
			Else
%>
									<div id="itemdiv<%=Int(i/3)%>" style="position: absolute; top: 0px;">
<%
			End If
		End If
%>
										<div style="height: 142px; padding: 5px; overflow: auto;">
											<div><%=gasOrderLineDescriptions(i)%></div>
										</div>
<%
	Next
	
	' Add hidden divs here
	If ((UBound(ganOrderLineIDs) + 1) Mod 3) > 0 Then
		For i = ((UBound(ganOrderLineIDs) + 1) Mod 3) To 2
%>
										<div style="height: 152px;">&nbsp;</div>
<%
		Next
	End If
	
	If UBound(ganOrderLineIDs) > 2 Then
%>
										<button style="width: 320px; color: #FFFFFF; background-color: #FF0000;" onclick="toggleDivs('itemdiv<%=Int(UBound(ganOrderLineIDs)/3)%>', 'itemdiv0')">Page <%=Int(UBound(ganOrderLineIDs)/3)+1%> of <%=Int(UBound(ganOrderLineIDs)/3)+1%><br/>(Next)</button>
<%
	End If
End If
%>
									</div>
								</div>
								<div style="width: 320px; text-align: center; background-color: #FFFFFF;">Tax: <%=FormatCurrency(gdTax + gdTax2)%>&nbsp; Delivery: <%=FormatCurrency(gdDeliveryCharge)%>&nbsp; Total: <%=FormatCurrency(gdOrderTotal)%></div>
								<div style="width: 320px; height: 152px; text-align: center; background-color: #FFFFFF;">
									<div id="tendertip" style="padding: 5px 0px 0px 0px; visibility: hidden;">&nbsp;</div>
<%
If gnPaymentTypeID = 0 Or gnPaymentTypeID = 1 Then
%>
									<div id="tendercash" style="padding: 5px 0px 0px 0px;">Cash: $0.00</div>
<%
Else
%>
									<div id="tendercash" style="padding: 5px 0px 0px 0px; visibility: hidden;">&nbsp;</div>
<%
End If

If gnPaymentTypeID = 2 Then
%>
									<div id="tendercheck" style="padding: 5px 0px 0px 0px;">Check: $0.00</div>
<%
Else
%>
									<div id="tendercheck" style="padding: 5px 0px 0px 0px; visibility: hidden;">&nbsp;</div>
<%
End If

If gnPaymentTypeID = 3 Then
%>
									<div id="tendercreditcard" style="padding: 5px 0px 0px 0px;">Credit Card: $0.00</div>
<%
Else
%>
									<div id="tendercreditcard" style="padding: 5px 0px 0px 0px; visibility: hidden;">&nbsp;</div>
<%
End If

If gnPaymentTypeID = 4 Then
%>
									<div id="tenderonaccount" style="padding: 5px 0px 0px 0px;">On Account <%=gnAccountID%>: $0.00</div>
<%
Else
%>
									<div id="tenderonaccount" style="padding: 5px 0px 0px 0px; visibility: hidden;">&nbsp;</div>
<%
End If
%>
									<div id="tenderbalance" style="padding: 5px 0px 0px 0px;">Balance Due: <%=FormatCurrency(gdOrderTotal)%></div>
								</div>
							</td>
						</tr>
					</table>
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
