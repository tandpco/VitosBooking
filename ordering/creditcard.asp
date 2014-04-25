<%
Option Explicit
Response.buffer = TRUE

If Session("SecurityID") = "" Then
	Response.Redirect("/default.asp")
End If

If Request("o").Count = 0 Then
	If Request("PayInAmount").Count = 0 Then
		If Request("Action") = "Apply" Then
			If Request("OpenCount").Count = 0 Then
				Response.Redirect("neworder.asp")
			Else
				If Not IsNumeric(Request("OpenCount")) Then
					Response.Redirect("neworder.asp")
				End If
			End If
		Else
			Response.Redirect("neworder.asp")
		End If
	Else
		If Not IsNumeric(Request("PayInAmount")) And Not IsNumeric(Request("PayInMethod")) And Not IsNumeric(Request("PayInCategory")) Then
			Response.Redirect("neworder.asp")
		End If
	End If
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
Dim gsPayInOutAccountNumber, gnPayAmount, gnPayInOutCategory, gsPayInOutCheckNumber, gsPayInFrom, gsPayMemo, gnPayInOutMethod
Dim gbNeedPrinterAlert
Dim gnOpenCount

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

If Request("o").Count <> 0 Then
	gnOrderID = CLng(Request("o"))
	Session("OrderID") = gnOrderID
	If GetOrderDetails(gnOrderID, gnSessionID, gsIPAddress, gnEmpID, gsRefID, gdtTransactionDate, gdtSubmitDate, gdtReleaseDate, gdtExpectedDate, gnStoreID, gnCustomerID, gsCustomerName, gsCustomerPhone, gnAddressID, gnOrderTypeID, gbIsPaid, gnPaymentTypeID, gsPaymentReference, gnAccountID, gdDeliveryCharge, gdDriverMoney, gdTax, gdTax2, gdTip, gnOrderStatusID, gsOrderNotes) Then
'		If gnStoreID <> Session("StoreID") Then
'			Response.Redirect("otherstore.asp?t=" & gnOrderTypeID & "&s=" & gnStoreID & "&c=" & gnCustomerID & "&a=" & gnAddressID)
'		End If
		
		If gnOrderStatusID <> 2 And DateValue(gdtTransactionDate) <> DateValue(Session("TransactionDate")) Then
			Response.Redirect("/error.asp?err=" & Server.URLEncode("Order is not From Today"))
		End If
		
' 2013-08-20 TAM: Allow edit order even if complete and paid
'		If gnOrderStatusID >= 10 Then
'			Response.Redirect("/error.asp?err=" & Server.URLEncode("Order is Complete"))
'		End If
'		
'		If gbIsPaid Then
'			Response.Redirect("/error.asp?err=" & Server.URLEncode("Order Has Been Paid"))
'		End If
	
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
		
		If Request("OrderNotes").Count > 0 Then
			gsOrderNotes = Request("OrderNotes")
			If SetOrderNotes(gnOrderID, gsOrderNotes) Then
				Session("OrderNotes") = gsOrderNotes
				Session("OrderEdited") = TRUE
			Else
				Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
			End If
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
Else
	gnStoreID = Session("StoreID")
	gnOrderID = -1
	gnOrderTypeID = -1
	gdDeliveryCharge = 0.00
	gdTax = 0.00
	gdTax2 = 0.00
	gdTip = 0.00
	
	If Request("Action") = "Apply" Then
		gnOpenCount = CLng(Request.Form("OpenCount"))
		gnAccountID = CLng(Request.Form("AccountID"))
		gnPaymentTypeID = CLng(Request.Form("PaymentTypeID"))
		
		gdOrderTotal = 0.00
		For i = 1 To gnOpenCount
			If CDbl(Request.Form("Paying-" & i)) > 0 Then
				gdOrderTotal = gdOrderTotal + CDbl(Request.Form("Paying-" & i))
			End If
		Next
	Else
		gsPayInOutAccountNumber = Request.Form("PayInAccountNumber")
		gnPayAmount = CDbl(Request.Form("PayInAmount"))
		gnPayInOutCategory = CLng(Request.Form("PayInCategory"))
		gsPayInOutCheckNumber = Request.Form("PayInCheckNumber")
		gsPayInFrom = Request.Form("PayInFrom")
		gsPayMemo = Request.Form("PayInMemo")
		gnPayInOutMethod = CLng(Request.Form("PayInMethod"))
		
		gnPaymentTypeID = gnPayInOutMethod
		gdOrderTotal = gnPayAmount
	End If
End If

If Not GetAllStoreManagers(gnStoreID, ganEmpIDs, ganEmployeeIDs, gasCardIDs) Then
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
var gsOrderNotes = "<%=gsOrderNotes%>";
var gnOrderID = <%=gnOrderID%>;
var gdOrderTotal = <%=gdOrderTotal%>;
var gnOrderTypeID = <%=gnOrderTypeID%>;
var gnPaymentTypeID = <%=gnPaymentTypeID%>;
var gsPaymentReference = "<%=gsPaymentReference%>";
var gdTenderCash = 0.00;
var gdTenderCheck = 0.00;
var gdTenderCreditCard = 0.00;
var gdTenderOnAccount = 0.00;
var gdTenderBalance = <%=gdOrderTotal%>;
var gdTipAmount = 0.00;
var gbDoPrint = true;
var gbDoSignaturePrint = true;
var gbPromptPrint = false;
var gbFormSubmitted = false;

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

if (gnOrderTypeID == 2) {
	gbPromptPrint = true;
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

function setCardNumFocus() {
	var loText;
	
	loText = ie4? eval("document.all.cardnum") : document.getElementById('cardnum');
	loText.focus();
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

function resetStatus() {
	var loText, loSpan;
	
	loText = ie4? eval("document.all.cardnum") : document.getElementById('cardnum');
	loText.style.color = "#000000";
	loText = ie4? eval("document.all.expdate") : document.getElementById('expdate');
	loText.style.color = "#000000";
	loSpan = ie4? eval("document.all.Status") : document.getElementById('Status');
	loSpan.innerHTML = "&nbsp;";
}

function addToCardNum(psDigit) {
	var loCardNum, lsCardNum;
	
	loCardNum = ie4? eval("document.all.cardnum") : document.getElementById('cardnum');
	lsCardNum = loCardNum.value;
	lsCardNum += psDigit;
	loCardNum.value = lsCardNum;
	
	resetStatus();
	
	resetRedirect();
}

function backspaceCardNum() {
	var loText, lsText;
	
	loText = ie4? eval("document.all.cardnum") : document.getElementById('cardnum');
	lsText = loText.value;
	if (lsText.length > 0) {
		lsText = lsText.substr(0, (lsText.length - 1));
		loText.value = lsText;
	}
	
	resetStatus();
	
	resetRedirect();
}

function checkSwipeEnterKey() {
	var loText, loDiv;
	
	if (event.keyCode == 13) {
		event.cancelBubble = true;
		event.returnValue = false;
		
		loText = ie4? eval("document.all.cardnum") : document.getElementById('cardnum');
		if (loText.value.length < 17) {
			return false;
		}
		else {
			if (loText.value.indexOf("?;") == -1) {
				return false;
			}
		}
		
		setCredit();
	}
}

function addToExpDate(psDigit) {
	var loText, lsText;
	
	loText = ie4? eval("document.all.expdate") : document.getElementById('expdate');
	lsText = loText.value;
	lsText += psDigit;
	loText.value = lsText;
	
	resetStatus();
	
	resetRedirect();
}

function backspaceExpDate() {
	var loText, lsText;
	
	loText = ie4? eval("document.all.expdate") : document.getElementById('expdate');
	lsText = loText.value;
	if (lsText.length > 0) {
		lsText = lsText.substr(0, (lsText.length - 1));
		loText.value = lsText;
	}
	
	resetStatus();
	
	resetRedirect();
}

function setCredit() {
	var loText, loDiv, loSpan, lsTrack1, lsTrack2;
	
	loText = ie4? eval("document.all.cardnum") : document.getElementById('cardnum');
	if (loText.value.length < 17) {
		if (loText.value.length < 15) {
			loText.style.color = "#FF0000";
			loSpan = ie4? eval("document.all.Status") : document.getElementById('Status');
			loSpan.innerHTML = "<strong>INVALID CARD NUMBER - NOT ENOUGH DIGITS</strong>";
			return false;
		}
		if (isNaN(Number(loText.value))) {
			loText.style.color = "#FF0000";
			loSpan = ie4? eval("document.all.Status") : document.getElementById('Status');
			loSpan.innerHTML = "<strong>INVALID CARD NUMBER</strong>";
			return false;
		}
		loText = ie4? eval("document.all.expdate") : document.getElementById('expdate');
		if (loText.value.length != 4) {
			loText.style.color = "#FF0000";
			loSpan = ie4? eval("document.all.Status") : document.getElementById('Status');
			loSpan.innerHTML = "<strong>INVALID EXPIRATION - NOT ENOUGH DIGITS</strong>";
			return false;
		}
		if (isNaN(Number(loText.value))) {
			loText.style.color = "#FF0000";
			loSpan = ie4? eval("document.all.Status") : document.getElementById('Status');
			loSpan.innerHTML = "<strong>INVALID EXPIRATION</strong>";
			return false;
		}
	}
	else {
		if (loText.value.indexOf("?;") == -1) {
			loText.style.color = "#FF0000";
			loSpan = ie4? eval("document.all.Status") : document.getElementById('Status');
			loSpan.innerHTML = "<strong>INVALID CARD NUMBER - TOO MANY DIGITS</strong>";
			return false;
		}
		else {
			if (loText.value.indexOf("?;") < 2) {
				loText.value = "";
				loSpan = ie4? eval("document.all.Status") : document.getElementById('Status');
				loSpan.innerHTML = "<strong>BAD DATA DETECTED - PLEASE SWIPE CARD AGAIN</strong>";
				return false;
			}
			else {
				lsTrack1 = loText.value.substr(0, loText.value.indexOf("?;"));
				lsTrack2 = loText.value.substr((loText.value.indexOf("?;") + 2));
				if ((lsTrack2.length < 2) || (lsTrack2.indexOf("=") < 2)) {
					loText.value = "";
					loSpan = ie4? eval("document.all.Status") : document.getElementById('Status');
					loSpan.innerHTML = "<strong>BAD DATA DETECTED - PLEASE SWIPE CARD AGAIN</strong>";
					return false;
				}
			}
		}
	}
	
	gnPaymentTypeID = 3;
	gdTenderCheck = 0.00;
	gdTenderCash = 0.00;
	gdTenderCreditCard = <%=gdOrderTotal%>;
	gdTenderOnAccount = 0.00;
	gdTenderBalance = <%=gdOrderTotal%> - gdTenderCreditCard;
	gdTenderBalance = Math.round(gdTenderBalance * Math.pow(10,2))/Math.pow(10,2);
	gsPaymentReference = "";
	
	loDiv = ie4? eval("document.all.tendercash") : document.getElementById('tendercash');
	loDiv.innerHTML = "&nbsp;";
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tenderonaccount") : document.getElementById('tenderonaccount');
	loDiv.innerHTML = "&nbsp;";
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tendercheck") : document.getElementById('tendercheck');
	loDiv.innerHTML = "&nbsp;";
	loDiv.style.visibility = "hidden";
	
	if (gnOrderID == -1)
	{
		setTip();
	}
	else
	{
		getTip();
	}
}

function getTip() {
	var loTip, loDiv;

	loTip = ie4? eval("document.all.tip") : document.getElementById('tip');
	loTip.value = "";
	
	loDiv = ie4? eval("document.all.tenderdiv") : document.getElementById('tenderdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.creditdiv") : document.getElementById('creditdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById('confirmreprint');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.printnow") : document.getElementById('printnow');
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
	var loText, lsText;
	
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
	
	loDiv = ie4? eval("document.all.tendercreditcard") : document.getElementById('tendercreditcard');
	loDiv.innerHTML = "Credit Card: " + FormatCurrency(gdTenderCreditCard + gdTipAmount);
	loDiv.style.visibility = "visible";
	loDiv = ie4? eval("document.all.tenderbalance") : document.getElementById('tenderbalance');
	if (gdTenderBalance < 0) {
		loDiv.innerHTML = "Change Due: " + FormatCurrency(gdTenderBalance);
	}
	else {
		loDiv.innerHTML = "Balance Due: " + FormatCurrency(gdTenderBalance);
	}
	
	loDiv = ie4? eval("document.all.creditdiv") : document.getElementById('creditdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tipdiv") : document.getElementById('tipdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById('confirmreprint');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.printnow") : document.getElementById('printnow');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tenderdiv") : document.getElementById('tenderdiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function cancelCreditCard() {
	if (gnOrderID == -1)
	{
		window.location = "../main.asp";
	}
	else
	{
		window.location = "payment.asp?o=<%=gnOrderID%>";
	}
}

function gotoConfirmReprint() {
	var loDiv;
	
	loDiv = ie4? eval("document.all.creditdiv") : document.getElementById('creditdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tipdiv") : document.getElementById('tipdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tenderdiv") : document.getElementById('tenderdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.printnow") : document.getElementById('printnow');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById('confirmreprint');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function cancelConfirmReprint() {
	var loDiv;
	
	loDiv = ie4? eval("document.all.creditdiv") : document.getElementById('creditdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tipdiv") : document.getElementById('tipdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById('confirmreprint');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.printnow") : document.getElementById('printnow');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tenderdiv") : document.getElementById('tenderdiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function cancelPrintNow() {
	var loDiv;
	
	loDiv = ie4? eval("document.all.creditdiv") : document.getElementById('creditdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tipdiv") : document.getElementById('tipdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById('confirmreprint');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.printnow") : document.getElementById('printnow');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.tenderdiv") : document.getElementById('tenderdiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function goChange() {
	var loText, lsText, loForm;
	
	if (gbPromptPrint) {
		loDiv = ie4? eval("document.all.creditdiv") : document.getElementById('creditdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.tipdiv") : document.getElementById('tipdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.tenderdiv") : document.getElementById('tenderdiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById('confirmreprint');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.printnow") : document.getElementById('printnow');
		loDiv.style.visibility = "visible";
	}
	else {
		if (gbFormSubmitted) {
			loForm = ie4? eval("document.all.creditcardform") : document.getElementById('creditcardform');
			loForm.action = "";
		}
		else {
			gbFormSubmitted = true;
			
			loText = ie4? eval("document.all.cardnum") : document.getElementById('cardnum');
			lsText = loText.value;
			loText = ie4? eval("document.all.xcardnum") : document.getElementById('xcardnum');
			loText.value = lsText;
			loText = ie4? eval("document.all.expdate") : document.getElementById('expdate');
			lsText = loText.value;
			loText = ie4? eval("document.all.xexpdate") : document.getElementById('xexpdate');
			loText.value = lsText;
			loText = ie4? eval("document.all.tip") : document.getElementById('tip');
			lsText = loText.value;
			loText = ie4? eval("document.all.xtip") : document.getElementById('xtip');
			loText.value = lsText;
			if (!gbDoPrint)
			{
				loText = ie4? eval("document.all.q") : document.getElementById('q');
				loText.value = "yes";
			}
			if (!gbDoSignaturePrint)
			{
				loText = ie4? eval("document.all.signatureprint") : document.getElementById('signatureprint');
				loText.value = "no";
			}
			
			loForm = ie4? eval("document.all.creditcardform") : document.getElementById('creditcardform');
			loForm.action = "change.asp";
			loForm.submit();
		}
	}
}
//-->
</script>
</head>

<body onload="clockInit(clockLocalStartTime, clockServerStartTime); clockOnLoad(); setCardNumFocus();" onunload="clockOnUnload()">

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
								<form method="post" action="" name="creditcardform" id="creditcardform">
<%
If gnOrderID = -1 Then
	If Request("Action") = "Apply" Then
%>
							         <input type="hidden" name="Action" value="Apply" />
							         <input type="hidden" name="AccountID" value="<%=gnAccountID%>" />
							         <input type="hidden" name="PaymentTypeID" value="<%=gnPaymentTypeID%>" />
							         <input type="hidden" name="OpenCount" value="<%=gnOpenCount%>" />
<%
		For i = 1 to gnOpenCount
%>
							         <input type="hidden" name="OrderID-<%=i%>" value="<%=Request.Form("OrderID-" & i)%>" />
							         <input type="hidden" name="Discount-<%=i%>" value="<%=Request.Form("Discount-" & i)%>" />
							         <input type="hidden" name="Paying-<%=i%>" value="<%=Request.Form("Paying-" & i)%>" />
<%
		Next
	Else
%>
									<input type="hidden" name="PayInFrom" id="PayInFrom" value="<%=gsPayInFrom%>"/>
									<input type="hidden" name="PayInMethod" id="PayInMethod" value="<%=gnPayInOutMethod%>"/>
									<input type="hidden" name="PayInAmount" id="PayInAmount" value="<%=gnPayAmount%>"/>
									<input type="hidden" name="PayInCheckNumber" id="PayInCheckNumber" value="<%=gsPayInOutCheckNumber%>"/>
									<input type="hidden" name="PayInAccountNumber" id="PayInAccountNumber" value="<%=gsPayInOutAccountNumber%>"/>
									<input type="hidden" name="PayInCategory" id="PayInCategory" value="<%=gnPayInOutCategory%>"/>
									<input type="hidden" name="PayInMemo" id="PayInMemo" value="<%=gsPayMemo%>"/>
<%
	End If
Else
%>
									<input type="hidden" name="o" value="<%=gnOrderID%>" />
<%
End If
%>
									<input type="hidden" name="v" value="3" />
									<input type="hidden" name="s" value="<%=gdOrderTotal%>" />
									<input type="hidden" name="j" value="0.00" />
									<input type="hidden" name="xcardnum" id="xcardnum" value="" />
									<input type="hidden" name="xexpdate" id="xexpdate" value="" />
									<input type="hidden" name="xtip" id="xtip" value="" />
									<input type="hidden" name="q" id="q" value="no" />
									<input type="hidden" name="signatureprint" id="signatureprint" value="yes" />
								<div id="content" style="position: relative; width: 680px; height: 723px; overflow: auto;">
									<div id="tenderdiv" style="position: absolute; left: 0px; top: 0px; width: 680px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
										<p align="center"><strong>Tender Payment</strong></p>
<%
If gnOrderID = -1 Then
	If Request("Action") = "Apply" Then
%>
										<p align="center"><strong>Pay On Account</strong></p>
										<div style="height: 538px; padding: 10px;">
											<table width="100%">
												<tr>
													<td align="right" valign="top" width="100">From:</td>
													<td valign="top">&nbsp;</td>
													<td valign="top"><%=GetAccountName(Request.Form("AccountID"))%></td>
												</tr>
											</table>
										</div>
<%
	Else
%>
										<p align="center"><strong>Pay In</strong></p>
										<div style="height: 538px; padding: 10px;">
											<table width="100%">
												<tr>
													<td align="right" valign="top" width="100">From:</td>
													<td valign="top">&nbsp;</td>
													<td valign="top"><%=Request("PayInFrom")%></td>
												</tr>
												<tr>
													<td align="right" valign="top" width="100">Memo:</td>
													<td valign="top">&nbsp;</td>
													<td valign="top"><%=Request("PayInMemo")%></td>
												</tr>
											</table>
										</div>
<%
	End If
Else
%>
										<p align="center"><strong>Order #&nbsp;<%=gnOrderID%></strong></p>
										<div style="height: 538px; padding: 10px;">
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
<%
End If
%>
										<div style="position: relative;">
										<div style="position: absolute; top: 0px; left: 0px;"><button style="width: 337px;" onclick="cancelCreditCard(); return false;">Cancel</button></div>
<%
If Session("NewOrder") Then
%>
										<div style="position: absolute; top: 0px; left: 341px;"><button style="width: 337px;" onclick="goChange(); return false;">Process Card</button></div>
<%
Else
	If Session("OrderEdited") Then
%>
										<div style="position: absolute; top: 0px; left: 341px;"><button style="width: 337px;" onclick="gotoConfirmReprint(); return false;">Done</button></div>
<%
	Else
%>
										<div style="position: absolute; top: 0px; left: 341px;"><button style="width: 337px;" onclick="goChange(); return false;">Process Card</button></div>
<%
	End If
End If
%>
										</div>
									</div>
									<div id="creditdiv" style="position: absolute; left: 0px; top: 0px; width: 680px; height: 723px; background-color: #fbf3c5;">
										<table align="center" cellpadding="0" cellspacing="0">
											<tr>
												<td colspan="3"><div align="center">
													<strong>ENTER CARD NUMBER</strong></div></td>
												<td width="25">&nbsp;</td>
												<td colspan="3"><div align="center">
													<strong>ENTER EXP DATE AS MMYY</strong></div></td>
											</tr>
											<tr>
												<td colspan="3"><div align="center">
													<input type="text" id="cardnum" style="width: 200px" autocomplete="off" onkeydown="checkSwipeEnterKey();" /></div></td>
												<td>&nbsp;</td>
												<td colspan="3"><div align="center">
													<input type="text" id="expdate" style="width: 200px" autocomplete="off" onkeydown="disableEnterKey();" /></div></td>
											</tr>
											<tr>
												<td><button onclick="addToCardNum('1'); return false;">1</button></td>
												<td><button onclick="addToCardNum('2'); return false;">2</button></td>
												<td><button onclick="addToCardNum('3'); return false;">3</button></td>
												<td width="25">&nbsp;</td>
												<td><button onclick="addToExpDate('1'); return false;">1</button></td>
												<td><button onclick="addToExpDate('2'); return false;">2</button></td>
												<td><button onclick="addToExpDate('3'); return false;">3</button></td>
											</tr>
											<tr>
												<td><button onclick="addToCardNum('4'); return false;">4</button></td>
												<td><button onclick="addToCardNum('5'); return false;">5</button></td>
												<td><button onclick="addToCardNum('6'); return false;">6</button></td>
												<td width="25">&nbsp;</td>
												<td><button onclick="addToExpDate('4'); return false;">4</button></td>
												<td><button onclick="addToExpDate('5'); return false;">5</button></td>
												<td><button onclick="addToExpDate('6'); return false;">6</button></td>
											</tr>
											<tr>
												<td><button onclick="addToCardNum('7'); return false;">7</button></td>
												<td><button onclick="addToCardNum('8'); return false;">8</button></td>
												<td><button onclick="addToCardNum('9'); return false;">9</button></td>
												<td width="25">&nbsp;</td>
												<td><button onclick="addToExpDate('7'); return false;">7</button></td>
												<td><button onclick="addToExpDate('8'); return false;">8</button></td>
												<td><button onclick="addToExpDate('9'); return false;">9</button></td>
											</tr>
											<tr>
												<td>&nbsp;</td>
												<td><button onclick="addToCardNum('0'); return false;">0</button></td>
												<td><button onclick="backspaceCardNum(); return false;">Bksp</button></td>
												<td width="25">&nbsp;</td>
												<td>&nbsp;</td>
												<td><button onclick="addToExpDate('0'); return false;">0</button></td>
												<td><button onclick="backspaceExpDate(); return false;">Bksp</button></td>
											</tr>
											<tr>
												<td colspan="7">&nbsp;</td>
											</tr>
											<tr>
												<td colspan="7" align="center"><button onclick="cancelCreditCard(); return false;">Cancel</button>&nbsp;<button onclick="setCredit(); return false;">OK</button></td>
											</tr>
											<tr>
												<td colspan="7">&nbsp;</td>
											</tr>
											<tr>
												<td colspan="7" algin="center"><center><span name="Status" id="Status">&nbsp;</span></center></td>
											</tr>
										</table>
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
												<td><button onclick="addToTip('1'); return false;">1</button></td>
												<td><button onclick="addToTip('2'); return false;">2</button></td>
												<td><button onclick="addToTip('3'); return false;">3</button></td>
											</tr>
											<tr>
												<td><button onclick="addToTip('4'); return false;">4</button></td>
												<td><button onclick="addToTip('5'); return false;">5</button></td>
												<td><button onclick="addToTip('6'); return false;">6</button></td>
											</tr>
											<tr>
												<td><button onclick="addToTip('7'); return false;">7</button></td>
												<td><button onclick="addToTip('8'); return false;">8</button></td>
												<td><button onclick="addToTip('9'); return false;">9</button></td>
											</tr>
											<tr>
												<td>&nbsp;</td>
												<td><button onclick="addToTip('0'); return false;">0</button></td>
												<td><button onclick="backspaceTip(); return false;">Bksp</button></td>
											</tr>
											<tr>
												<td colspan="3">&nbsp;</td>
											</tr>
											<tr>
												<td colspan="3"><button onclick="cancelCreditCard(); return false;">Cancel</button>&nbsp;<button onclick="setTip(); return false;">OK</button></td>
											</tr>
										</table>
									</div>
									<div id="confirmreprint" style="position: absolute; top: 0px; left: 0px; width: 680px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
										<p align="center"><strong>Do you want to reprint this order?</strong><br/><br/>
										<button onclick="goChange(); return false;">Reprint</button>&nbsp;&nbsp;&nbsp;&nbsp;<button onclick="gbDoPrint = false; goChange(); return false;">No Reprint</button>&nbsp;&nbsp;&nbsp;&nbsp;<button onclick="cancelConfirmReprint(); return false;">Cancel</button>
										</p>
									</div>
									<div id="printnow" style="position: absolute; top: 0px; left: 0px; width: 680px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
										<p align="center"><strong>Do you want to print signature receipts for this order now?</strong><br/><br/>
										<button onclick="gbPromptPrint = false; goChange(); return false;">Print Now</button>&nbsp;&nbsp;&nbsp;&nbsp;<button onclick="gbPromptPrint = false; gbDoSignaturePrint = false; goChange(); return false;">Print Later</button>&nbsp;&nbsp;&nbsp;&nbsp;<button onclick="cancelPrintNow(); return false;">Cancel</button>
										</p>
									</div>
								</div>
								</form>
							</td>
							<td align="right" valign="top" width="330">
								<div style="position: relative; width: 320px; height: 539px; text-align: left; background-color: #FFFFFF;">
<%
If gnOrderID = -1 Then
%>
									<div id="payindiv" style="position: absolute; top: 0px;">
										<div style="height: 142px; padding: 5px; overflow: auto;">
<%
	If Request("Action") = "Apply" Then
%>
											<div>Pay On Account</div>
<%
	Else
%>
											<div>Pay In</div>
<%
	End If
%>
										</div>
<%
Else
	If ganOrderLineIDs(0) <> 0 Then
		For i = 0 To UBound(ganOrderLineIDs)
			If i Mod 3 = 0 Then
				If i > 0 Then
%>
										<button style="width: 320px" onclick="toggleDivs('itemdiv<%=Int(i/3)-1%>', 'itemdiv<%=Int(i/3)%>')">Page <%=Int(i/3)%> of <%=Int(UBound(ganOrderLineIDs)/3)+1%><br/>(Next)</button>
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
										<button style="width: 320px" onclick="toggleDivs('itemdiv<%=Int(UBound(ganOrderLineIDs)/3)%>', 'itemdiv0')">Page <%=Int(UBound(ganOrderLineIDs)/3)+1%> of <%=Int(UBound(ganOrderLineIDs)/3)+1%><br/>(Next)</button>
<%
		End If
	End If
End If
%>
									</div>
								</div>
								<div style="width: 320px; text-align: center; background-color: #FFFFFF;">Tax: <%=FormatCurrency(gdTax + gdTax2)%>&nbsp; Delivery: <%=FormatCurrency(gdDeliveryCharge)%>&nbsp; Total: <%=FormatCurrency(gdOrderTotal)%></div>
								<div style="width: 320px; height: 152px; text-align: center; background-color: #FFFFFF;">
									<div id="tendertip" style="padding: 5px 0px 0px 0px; visibility: hidden;">&nbsp;</div>
									<div id="tendercash" style="padding: 5px 0px 0px 0px; visibility: hidden;">&nbsp;</div>
									<div id="tendercheck" style="padding: 5px 0px 0px 0px; visibility: hidden;">&nbsp;</div>
									<div id="tendercreditcard" style="padding: 5px 0px 0px 0px; visibility: hidden;">&nbsp;</div>
									<div id="tenderonaccount" style="padding: 5px 0px 0px 0px; visibility: hidden;">&nbsp;</div>
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
