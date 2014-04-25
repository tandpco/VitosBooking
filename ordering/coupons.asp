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

If Request("l").Count > 0 Then
	If Not IsNumeric(Request("l")) Then
		Response.Redirect("neworder.asp")
	End If
	
	If Request("m").Count = 0 Then
		Response.Redirect("neworder.asp")
	Else
		If Not IsNumeric(Request("m")) Then
			Response.Redirect("neworder.asp")
		End If
	End If
End If

If Request("k").Count > 0 Then
	If Not IsNumeric(Request("k")) Then
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
Dim ganCouponIDs(), gasCouponDescriptions(), gasCouponShortDescriptions()
Dim lnTop, lnLeft
Dim gasCouponIDs, gbFound, gsCouponIDs, gsCoupons
Dim gnUnitID, gnSpecialtyID, gnSizeID, gnStyleID, gnHalf1SauceID, gnHalf2SauceID, gnHalf1SauceModifierID, gnHalf2SauceModifierID, gsOrderLineNotes, gnQuantity, gdOrderLineCost, gdOrderLineDiscount, gnCouponID
Dim ganOrderUnitIDs(), ganOrderSizeIDs(), ganOrderStyleIDs(), ganOrderSpecialtyIDs(), gnPos
Dim gbNeedPrinterAlert
Dim gasMPOReasons()

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

If Request("l").Count > 0 Then
	If SetManagerPriceOverride(Request("l"), Request("m"), Request("MPOReason")) Then
		gbFound = FALSE
		If Len(Session("CouponIDs")) > 0 Then
			gasCouponIDs = Split(Session("CouponIDs"), ",")
			For i = 0 To UBound(gasCouponIDs)
				If gasCouponIDs(i) = "1" Then
					gbFound = TRUE
				End If
			Next
		End If
		
		If Not gbFound Then
			If Len(Session("CouponIDs")) > 0 Then
				gsCouponIDs = Session("CouponIDs") & ",1"
			Else
				gsCouponIDs = "1"
			End If
		End If
		
		If RecalculateOrderTax(Session("StoreID"), gnOrderID) Then
			Session("CouponIDs") = gsCouponIDs
			Session("OrderEdited") = TRUE
		Else
			Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
		End If
	Else
		Response.Redirect("/error.asp?err=" & Server.URLEncode(gsLocalErrorMsg))
	End If
End If

If Request("k").Count > 0 Then
	If CLng(Request("k")) = 0 Then
'		If Len(Session("CouponIDs")) > 0 Then
			ClearAllCoupons gnOrderID
			If RecalculateOrderTax(Session("StoreID"), gnOrderID) Then
				Session("CouponIDs") = ""
				Session("OrderEdited") = TRUE
			Else
				Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
			End If
'		End If
	Else
		gbFound = FALSE
		If Len(Session("CouponIDs")) > 0 Then
			gasCouponIDs = Split(Session("CouponIDs"), ",")
			For i = 0 To UBound(gasCouponIDs)
				If gasCouponIDs(i) = Request("k") Then
					gbFound = TRUE
				End If
			Next
		End If
		
		If Not gbFound Then
			If Len(Session("CouponIDs")) > 0 Then
				gsCouponIDs = Session("CouponIDs") & "," & Request("k")
			Else
				gsCouponIDs = Request("k")
			End If
			
' 2013-11-14 TAM: Not sure why discount recalc would fail but it happened on a simple order and then didn't recalc the tax. For now always recalc tax.
'			If RecalculateOrderDiscounts(Session("StoreID"), gnOrderID, gsCouponIDs) Then
'				If RecalculateOrderTax(Session("StoreID"), gnOrderID) Then
'					Session("CouponIDs") = gsCouponIDs
'					Session("OrderEdited") = TRUE
'				Else
'					Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
'				End If
''Else
''Response.Write("Recalc Failed")
''Response.End
'			End If
			RecalculateOrderDiscounts Session("StoreID"), gnOrderID, gsCouponIDs
			If RecalculateOrderTax(Session("StoreID"), gnOrderID) Then
				Session("CouponIDs") = gsCouponIDs
				Session("OrderEdited") = TRUE
			Else
				Response.Redirect("/error.asp?o=" & gnOrderID & "&err=" & Server.URLEncode(gsDBErrorMessage))
			End If
		End If
	End If
End If

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

If Not GetOrderCoupons(gnOrderID, gsCoupons) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If Not GetAllStoreManagers(gnStoreID, ganEmpIDs, ganEmployeeIDs, gasCardIDs) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

'If Not GetActiveCoupons(gnStoreID, gnOrderTypeID, Session("TransactionDate"), FALSE, ganCouponIDs, gasCouponDescriptions, gasCouponShortDescriptions) Then
'	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
'End If
ReDim ganOrderUnitIDs(0), ganOrderSizeIDs(0), ganOrderStyleIDs(0), ganOrderSpecialtyIDs(0)

ganOrderUnitIDs(0) = 0
ganOrderSizeIDs(0) = 0
ganOrderStyleIDs(0) = 0
ganOrderSpecialtyIDs(0) = 0
gnPos = 0

If ganOrderLineIDs(0) <> 0 Then
	For i = 0 To UBound(ganOrderLineIDs)
		If GetOrderLineDetails(ganOrderLineIDs(i), gnUnitID, gnSpecialtyID, gnSizeID, gnStyleID, gnHalf1SauceID, gnHalf2SauceID, gnHalf1SauceModifierID, gnHalf2SauceModifierID, gsOrderLineNotes, gnQuantity, gdOrderLineCost, gdOrderLineDiscount, gnCouponID) Then
			If ganOrderUnitIDs(0) <> 0 Then
				gnPos = gnPos + 1
				
				ReDim Preserve ganOrderUnitIDs(gnPos), ganOrderSizeIDs(gnPos), ganOrderStyleIDs(gnPos), ganOrderSpecialtyIDs(gnPos)
			End If
			
			ganOrderUnitIDs(gnPos) = gnUnitID
			ganOrderSizeIDs(gnPos) = gnSizeID
			ganOrderStyleIDs(gnPos) = gnStyleID
			ganOrderSpecialtyIDs(gnPos) = gnSpecialtyID
		End If
	Next
	
	If Not GetActiveCouponsByUnitSizeSpecialtyStyle(gnStoreID, gnOrderTypeID, Session("TransactionDate"), FALSE, ganOrderUnitIDs, ganOrderSizeIDs, ganOrderStyleIDs, ganOrderSpecialtyIDs, ganCouponIDs, gasCouponDescriptions, gasCouponShortDescriptions) Then
		Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
	End If
End If

If Not GetMPOReasons(gasMPOReasons) Then
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
var gnOrderLines = <%=UBound(ganOrderLineIDs) + 1%>;
var gnCurrentLine = -1;
var gnCurrentOrderLineID = 0;
var gnCurrentOrderLineCost = 0.00;
var gbZeroOrder = false;
var gdPrice = 0.00;
var gbMPOReasonInLowerCase = false;

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

function highlightLine(pnLine, pnOrderLineID, pnCurrentOrderLineCost)
{
	var loDiv, i;
	
	gnCurrentLine = pnLine;
	gnCurrentOrderLineID = pnOrderLineID;
	gnCurrentOrderLineCost = pnCurrentOrderLineCost;
	
	if (pnLine != -1) {
		loDiv = ie4? eval("document.all.linediv" + pnLine.toString()) : document.getElementById("linediv" + pnLine.toString());
		loDiv.style.backgroundColor = "#000000";
		loDiv.style.color = "#FFFFFF";
	}
	
	for (i = 0; i < gnOrderLines; i++) {
		if (i != pnLine) {
			loDiv = ie4? eval("document.all.linediv" + i.toString()) : document.getElementById("linediv" + i.toString());
			loDiv.style.backgroundColor = "#FFFFFF";
			loDiv.style.color = "#000000";
		}
	}
	
	resetRedirect();
}

function gotoMgrOverride(pbZeroOrder)
{
	var loManager, loDiv;
	
	if (pbZeroOrder || (gnCurrentLine != -1)) {
		gbZeroOrder = pbZeroOrder;
		
		loManager = ie4? eval("document.all.manager") : document.getElementById('manager');
		loManager.value = "";
		
		loDiv = ie4? eval("document.all.coupondiv") : document.getElementById('coupondiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.pricediv") : document.getElementById('pricediv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.mporeasondiv") : document.getElementById('mporeasondiv');
		loDiv.style.visibility = "hidden";
		loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
		loDiv.style.visibility = "visible";
		
		loManager.focus();
	}
	
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
	
	gnCurrentLine = -1;
	highlightLine(-1, 0, 0.00);
	gbZeroOrder = false;
	
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.pricediv") : document.getElementById('pricediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.mporeasondiv") : document.getElementById('mporeasondiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.coupondiv") : document.getElementById('coupondiv');
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
			if (gbZeroOrder) {
				gotoMPOReason();
			}
			else {
				gotoPrice();
			}
		}
		else {
			loManager.value = "";
			loManager.focus();
		}
	}
	
	resetRedirect();
}

function gotoPrice() {
	var loText, loDiv;

	loText = ie4? eval("document.all.price") : document.getElementById('price');
	loText.value = "";
	
	loDiv = ie4? eval("document.all.coupondiv") : document.getElementById('coupondiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.mporeasondiv") : document.getElementById('mporeasondiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.pricediv") : document.getElementById('pricediv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function addToPrice(psDigit) {
	var loText, lsText, loRE;
	
	loText = ie4? eval("document.all.price") : document.getElementById('price');
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

function backspacePrice() {
	var loText, lsText, loRE;
	
	loText = ie4? eval("document.all.price") : document.getElementById('price');
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

function cancelPrice() {
	var loText, loDiv;
	
	loText = ie4? eval("document.all.price") : document.getElementById('price');
	loText.value = "";
	
	gnCurrentLine = -1;
	highlightLine(-1, 0, 0.00);
	
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.pricediv") : document.getElementById('pricediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.mporeasondiv") : document.getElementById('mporeasondiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.coupondiv") : document.getElementById('coupondiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function setPrice() {
	var loText, ldPrice, lsLocation;
	
	loText = ie4? eval("document.all.price") : document.getElementById('price');
	if (loText.value == "") {
		ldPrice = 0.00;
	}
	else {
		ldPrice = new Number(loText.value);
		if (ldPrice >= gnCurrentOrderLineCost) {
			loText.value = "";
			loText.focus();
			return false;
		}
	}
	
	gdPrice = ldPrice;
	
	gotoMPOReason();
}

function gotoMPOReason() {
	var loNotes, loDiv;
	
	loNotes = ie4? eval("document.all.mporeason") : document.getElementById('mporeason');
	loNotes.value = "";
	
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.pricediv") : document.getElementById('pricediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.coupondiv") : document.getElementById('coupondiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.mporeasondiv") : document.getElementById('mporeasondiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function cancelMPOReason() {
	var loText, loDiv;
	
	loText = ie4? eval("document.all.mporeason") : document.getElementById('mporeason');
	loText.value = "";
	
	gnCurrentLine = -1;
	highlightLine(-1, 0, 0.00);
	
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.pricediv") : document.getElementById('pricediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.mporeasondiv") : document.getElementById('mporeasondiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.coupondiv") : document.getElementById('coupondiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function addToMPOReason(psDigit) {
	var loNotes, lsNotes;
	
	loNotes = ie4? eval("document.all.mporeason") : document.getElementById('mporeason');
	lsNotes = loNotes.value;
	
	if (psDigit.length > 1) {
		if (lsNotes.length > 0) {
			lsNotes = lsNotes + " ";
		}
	}
	
	if (gbMPOReasonInLowerCase) {
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

function backspaceMPOReason() {
	var loNotes, lsNotes;
	
	loNotes = ie4? eval("document.all.mporeason") : document.getElementById('mporeason');
	lsNotes = loNotes.value;
	if (lsNotes.length > 0) {
		lsNotes = lsNotes.substr(0, (lsNotes.length - 1));
		loNotes.value = lsNotes;
	}
	
	resetRedirect();
}

function clearMPOReason() {
	var loNotes;
	
	loNotes = ie4? eval("document.all.mporeason") : document.getElementById('mporeason');
	loNotes.value = "";
	
	resetRedirect();
}

function properMPOReason() {
	var loNotes, i, lsText, lbDoUpper;
	
	loNotes = ie4? eval("document.all.mporeason") : document.getElementById('mporeason');
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

function shiftMPOReason() {
	var loObj;
	
	if (gbMPOReasonInLowerCase) {
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
	
	gbMPOReasonInLowerCase = !gbMPOReasonInLowerCase;
	
	resetRedirect();
}

function saveMPOReason() {
	var loNotes, lsLocation;
	
	loNotes = ie4? eval("document.all.mporeason") : document.getElementById('mporeason');
	gsOrderNotes = loNotes.value;
	if (gsOrderNotes.length > 0) {
		if (gbZeroOrder) {
			lsLocation = "unitselect.asp?o=" + gnOrderID.toString() + "&zero=yes&MPOReason=" + encodeURIComponent(gsOrderNotes);
		}
		else {
			lsLocation = "coupons.asp?o=" + gnOrderID.toString() + "&l=" + gnCurrentOrderLineID.toString() + "&m=" + gdPrice.toString() + "&MPOReason=" + encodeURIComponent(gsOrderNotes);
		}
		window.location = lsLocation;
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
						<div id="coupondiv" style="position: absolute; left: 0px; top: 0px; width: 1010px; height: 723px; background-color: #fbf3c5;">
							<table cellpadding="0" cellspacing="0" width="1010" height="723">
								<tr>
									<td valign="top" width="680">
										<p align="center"><strong>Select Coupons</strong></p>
										<p align="center"><strong>Order #&nbsp;<%=gnOrderID%></strong></p>
										<div style="position: relative; width: 670px; height: 558px;">
<%
If ganCouponIDs(0) <> 0 Then
	lnTop = 0
	lnLeft = 0
	For i = 0 To UBound(ganCouponIDs)
		If i Mod 11 = 0 Then
			If i > 0 And UBound(ganCouponIDs) > 11 Then
%>
												<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 337px; height: 75px;" onclick="toggleDivs('coupondiv<%=Int(i/11)-1%>', 'coupondiv<%=Int(i/11)%>')">(Next)</button></div>
											</div>
											<div id="coupondiv<%=Int(i/11)%>" style="position: absolute; top: 0px; left: 0px; width: 670px; visibility: hidden;">
<%
				If Session("SecurityID") > 1 And ((Left(Request.ServerVariables("remote_addr"), 9) = "10.0.254." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.1." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.2.") OR Session("Swipe")) Then
%>
												<div style="position: absolute; top: 0px; left: 0px;"><button style="width: 337px; height: 75px;" onclick="if (gnCurrentLine != -1) {gotoPrice();};">Manager Price Override</button></div>
												<div style="position: absolute; top: 0px; left: 338px;"><button style="width: 337px; height: 75px;" onclick="window.location = 'unitselect.asp?o=<%=gnOrderID%>&zero=yes';">Zero Out Order</button></div>
<%
				Else
%>
												<div style="position: absolute; top: 0px; left: 0px;"><button style="width: 337px; height: 75px;" onclick="gotoMgrOverride(false);">Manager Price Override</button></div>
												<div style="position: absolute; top: 0px; left: 338px;"><button style="width: 337px; height: 75px;" onclick="gotoMgrOverride(true);">Zero Out Order</button></div>
<%
				End If
				
				lnTop = 76
				lnLeft = 0
			Else
				If i = 0 Then
%>
											<div id="coupondiv<%=Int(i/11)%>" style="position: absolute; top: 0px; left: 0px; width: 670px;">
<%
					If Session("SecurityID") > 1 And ((Left(Request.ServerVariables("remote_addr"), 9) = "10.0.254." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.1." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.2.") OR Session("Swipe")) Then
%>
												<div style="position: absolute; top: 0px; left: 0px;"><button style="width: 337px; height: 75px;" onclick="if (gnCurrentLine != -1) {gotoPrice();};">Manager Price Override</button></div>
												<div style="position: absolute; top: 0px; left: 338px;"><button style="width: 337px; height: 75px;" onclick="window.location = 'unitselect.asp?o=<%=gnOrderID%>&zero=yes';">Zero Out Order</button></div>
<%
					Else
%>
												<div style="position: absolute; top: 0px; left: 0px;"><button style="width: 337px; height: 75px;" onclick="gotoMgrOverride(false);">Manager Price Override</button></div>
												<div style="position: absolute; top: 0px; left: 338px;"><button style="width: 337px; height: 75px;" onclick="gotoMgrOverride(true);">Zero Out Order</button></div>
<%
					End If
					
					lnTop = 76
					lnLeft = 0
				End If
			End If
		End If
%>
												<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 337px; height: 75px;" onclick="window.location = 'coupons.asp?o=<%=gnOrderID%>&k=<%=ganCouponIDs(i)%>';"><%=gasCouponShortDescriptions(i)%></button></div>
<%
		If lnLeft = 338 Then
			lnTop = lnTop + 76
			lnLeft = 0
		Else
			lnLeft = lnLeft + 338
		End If
	Next
	
	' Add hidden buttons here
	If ((UBound(ganCouponIDs) + 1) Mod 11) > 0 And UBound(ganCouponIDs) <> 11 Then
		For i = ((UBound(ganCouponIDs) + 1) Mod 11) To 10
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
	
	If UBound(ganCouponIDs) > 11 Then
%>
												<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 337px; height: 75px;" onclick="toggleDivs('coupondiv<%=Int(UBound(ganCouponIDs)/11)%>', 'coupondiv0')">(Next)</button></div>
<%
	Else
		If UBound(ganCouponIDs) <> 11 Then
%>
												<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 337px; height: 75px; background-color: #C0C0C0;">&nbsp;</button></div>
<%
		End If
	End If
%>
											</div>
<%
Else
%>
											<div id="coupondiv" style="position: absolute; top: 0px; left: 0px; width: 670px;">
<%
	If Session("SecurityID") > 1 And ((Left(Request.ServerVariables("remote_addr"), 9) = "10.0.254." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.1." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.2.") OR Session("Swipe")) Then
%>
												<div style="position: absolute; top: 0px; left: 0px;"><button style="width: 337px; height: 75px;" onclick="if (gnCurrentLine != -1) {gotoPrice();};">Manager Price Override</button></div>
												<div style="position: absolute; top: 0px; left: 338px;"><button style="width: 337px; height: 75px;" onclick="window.location = 'unitselect.asp?o=<%=gnOrderID%>&zero=yes';">Zero Out Order</button></div>
<%
	Else
%>
												<div style="position: absolute; top: 0px; left: 0px;"><button style="width: 337px; height: 75px;" onclick="gotoMgrOverride(false);">Manager Price Override</button></div>
												<div style="position: absolute; top: 0px; left: 338px;"><button style="width: 337px; height: 75px;" onclick="gotoMgrOverride(true);">Zero Out Order</button></div>
<%
	End If
%>
											</div>
<%
End If
%>
										</div>
										<div style="position: relative;">
										<div style="position: absolute; top: 0px; left: 0px;"><button style="width: 337px;" onclick="window.location = 'coupons.asp?o=<%=gnOrderID%>&k=0';">Clear All Discounts</button></div>
										<div style="position: absolute; top: 0px; left: 338px;"><button style="width: 337px;" onclick="window.location = 'unitselect.asp?o=<%=gnOrderID%>';">Done</button></div>
										</div>
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
													<div id="linediv<%=i%>" onclick="highlightLine(<%=i%>, <%=ganOrderLineIDs(i)%>, <%=gadCost(i)%>);"><%=gasOrderLineDescriptions(i)%></div>
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
						<div id="managerdiv" style="position: absolute; left: 0px; top: 0px; width: 1010px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
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
						<div id="pricediv" style="position: absolute; left: 0px; top: 0px; width: 1010px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
							<table align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="3"><div align="center">
										<strong>ENTER NEW PRICE</strong></div></td>
								</tr>
								<tr>
									<td colspan="3"><div align="center">
										<input type="text" id="price" style="width: 200px" autocomplete="off" onkeydown="disableEnterKey();" /></div></td>
								</tr>
								<tr>
									<td><button onclick="addToPrice('1')">1</button></td>
									<td><button onclick="addToPrice('2')">2</button></td>
									<td><button onclick="addToPrice('3')">3</button></td>
								</tr>
								<tr>
									<td><button onclick="addToPrice('4')">4</button></td>
									<td><button onclick="addToPrice('5')">5</button></td>
									<td><button onclick="addToPrice('6')">6</button></td>
								</tr>
								<tr>
									<td><button onclick="addToPrice('7')">7</button></td>
									<td><button onclick="addToPrice('8')">8</button></td>
									<td><button onclick="addToPrice('9')">9</button></td>
								</tr>
								<tr>
									<td><button onclick="cancelPrice()">Cancel</button></td>
									<td><button onclick="addToPrice('0')">0</button></td>
									<td><button onclick="backspacePrice()">Bksp</button></td>
								</tr>
								<tr>
									<td colspan="3"><button style="width: 235px;" onclick="setPrice()">OK</button></td>
								</tr>
							</table>
						</div>
						<div id="mporeasondiv" style="position: absolute; left: 0px; top: 0px; width: 1010px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
							<table align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="9"><div align="center">
										<strong>WHY IS THE PRICE BEING OVERRIDDEN</strong></div></td>
								</tr>
								<tr>
									<td colspan="9"><div align="center">
										<textarea id="mporeason" style="width: 930px; height: 60px;"></textarea></div></td>
								</tr>
								<tr>
<%
If Len(gasMPOReasons(0)) > 0 Then
	For i = 0 To UBound(gasMPOReasons)
%>
									<td width="11%"><button style="width: 100px;" onclick="addToMPOReason('<%=gasMPOReasons(i)%>')"><%=gasMPOReasons(i)%></button></td>
<%
		If i = 5 Then
			Exit For
		End If
	Next
	
	If UBound(gasMPOReasons) < 5 Then
		For i = 4 To UBound(gasMPOReasons) Step -1
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
									<td width="11%"><button style="width: 100px;" onclick="cancelMPOReason();">Cancel</button></td>
									<td width="11%"><button style="width: 100px;" onclick="clearMPOReason();">Clear</button></td>
									<td width="11%"><button style="width: 100px;" onclick="saveMPOReason();">Done</button></td>
								</tr>
								<tr>
<%
If UBound(gasMPOReasons) > 5 Then
	For i = 6 To UBound(gasMPOReasons)
%>
									<td width="11%"><button style="width: 100px;" onclick="addToMPOReason('<%=gasMPOReasons(i)%>')"><%=gasMPOReasons(i)%></button></td>
<%
		If i = 14 Then
			Exit For
		End If
	Next
	
	If UBound(gasMPOReasons) < 14 Then
		For i = 13 To UBound(gasMPOReasons) Step -1
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
										<button onclick="addToMPOReason('+')">+</button><button onclick="addToMPOReason('!')">!</button><button onclick="addToMPOReason('@')">@</button><button onclick="addToMPOReason('#')">#</button><button onclick="addToMPOReason('$')">$</button><button onclick="addToMPOReason('%')">%</button><button onclick="addToMPOReason('^')">^</button><button onclick="addToMPOReason('&')">&amp;</button><button onclick="addToMPOReason('*')">*</button><button onclick="addToMPOReason('(')">(</button><button onclick="addToMPOReason(')')">)</button><button onclick="addToMPOReason(':')">:</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="addToMPOReason('=')">=</button><button onclick="addToMPOReason('1')">1</button><button onclick="addToMPOReason('2')">2</button><button onclick="addToMPOReason('3')">3</button><button onclick="addToMPOReason('4')">4</button><button onclick="addToMPOReason('5')">5</button><button onclick="addToMPOReason('6')">6</button><button onclick="addToMPOReason('7')">7</button><button onclick="addToMPOReason('8')">8</button><button onclick="addToMPOReason('9')">9</button><button onclick="addToMPOReason('0')">0</button><button onclick="addToMPOReason('?')">?</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="addToMPOReason('\'')">'</button><button name="key-q" id="key-q" onclick="addToMPOReason('Q')">Q</button><button name="key-w" id="key-w" onclick="addToMPOReason('W')">W</button><button name="key-e" id="key-e" onclick="addToMPOReason('E')">E</button><button name="key-r" id="key-r" onclick="addToMPOReason('R')">R</button><button name="key-t" id="key-t" onclick="addToMPOReason('T')">T</button><button name="key-y" id="key-y" onclick="addToMPOReason('Y')">Y</button><button name="key-u" id="key-u" onclick="addToMPOReason('U')">U</button><button name="key-i" id="key-i" onclick="addToMPOReason('I')">I</button><button name="key-o" id="key-o" onclick="addToMPOReason('O')">O</button><button name="key-p" id="key-p" onclick="addToMPOReason('P')">P</button><button onclick="addToMPOReason('/')">/</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="addToMPOReason('.')">.</button><button name="key-a" id="key-a" onclick="addToMPOReason('A')">A</button><button name="key-s" id="key-s" onclick="addToMPOReason('S')">S</button><button name="key-d" id="key-d" onclick="addToMPOReason('D')">D</button><button name="key-f" id="key-f" onclick="addToMPOReason('F')">F</button><button name="key-g" id="key-g" onclick="addToMPOReason('G')">G</button><button name="key-h" id="key-h" onclick="addToMPOReason('H')">H</button><button name="key-j" id="key-j" onclick="addToMPOReason('J')">J</button><button name="key-k" id="key-k" onclick="addToMPOReason('K')">K</button><button name="key-l" id="key-l" onclick="addToMPOReason('L')">L</button><button onclick="addToMPOReason(',')">,</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="shiftMPOReason()">Shift</button><button onclick="addToMPOReason('<')">&lt;</button><button name="key-z" id="key-z" onclick="addToMPOReason('Z')">Z</button><button name="key-x" id="key-x" onclick="addToMPOReason('X')">X</button><button name="key-c" id="key-c" onclick="addToMPOReason('C')">C</button><button name="key-v" id="key-v" onclick="addToMPOReason('V')">V</button><button name="key-b" id="key-b" onclick="addToMPOReason('B')">B</button><button name="key-n" id="key-n" onclick="addToMPOReason('N')">N</button><button name="key-m" id="key-m" onclick="addToMPOReason('M')">M</button><button onclick="addToMPOReason('>')">&gt;</button><button onclick="properMPOReason()">Proper</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button style="width: 150px;" onclick="addToMPOReason('1/2')">1/2</button>&nbsp;<button style="width: 600px;" onclick="addToMPOReason(' ')">Space</button>&nbsp;<button style="width: 150px;" onclick="backspaceMPOReason()">Bksp</button></div></td>
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
