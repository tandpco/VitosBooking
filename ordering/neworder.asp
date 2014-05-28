<%
Option Explicit
Response.buffer = TRUE

If Session("SecurityID") = "" Then
	Response.Redirect("/default.asp")
End If
%>
<!-- #Include Virtual="include2/utility.asp" -->
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
<!-- #Include Virtual="include2/heartland.asp" -->
<!-- #Include Virtual="include2/mail.asp" -->
<%
OpenSQLConn 'Open Database

Dim gsExpectedDateTime
Dim gbShowMenuButtons
Dim gsReference
Dim gsLocalErrorMsg, gbNeedPrinterAlert
Dim gnCustomerID, gnAddressID, gnOrderPick
Dim gsAccountName, gsPrimaryContactName, gsPrimaryContactEmail, gsSMSEmail, gsMailBody
Dim gsStoreName, gsStoreAddress1, gsStoreAddress2, gsStoreCity, gsStoreState, gsStorePostalCode, gsStorePhone, gsStoreFAX, gsStoreHours
Dim gsVoidReason

gbShowMenuButtons 	= TRUE
gbNeedPrinterAlert 	= FALSE

If Request("q").Count <> 0 Then
	If Request("q") = "yes" Then
		Session("QuickMode") = TRUE
	End If
End If

If Request("o").Count = 0 Then
	If Session("OrderID") <> 0 Then
		If Session("OrderLineCount") = 0 Or Request("cancel") = "yes" Then
			
			If Request("VoidReason").Count > 0 Then
				gsVoidReason = Trim(Request("VoidReason"))
			Else
				gsVoidReason = ""
			End If
			
			If Not CancelOrder(Session("OrderID"), Session("EmpID"), gsVoidReason) Then
				Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
			End If
			
%>
<!-- #Include Virtual="include2/clearorder.asp" -->
<%
			Response.Redirect("neworder.asp")
		Else
			If Request("OrderNotes").Count > 0 Then
				gsOrderNotes = Request("OrderNotes")
				If Not SetOrderNotes(CLng(Session("OrderID")), Request("OrderNotes")) Then
					Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
				End If
			End If
			
			If Request("d").Count > 0 And Request("e").Count > 0 And Request("g").Count > 0 Then
				If Not IsNumeric(Request("d")) Or Not IsNumeric(Request("e")) Or Not IsNumeric(Request("g")) Then
					Response.Redirect("/default.asp")
				End If
				
				gsExpectedDateTime = Left(Request("e"), 2) & "/" & Mid(Request("e"), 3, 2) & "/" & Right(Request("e"), 2) & " " & Left(Request("g"), 2) & ":" & Mid(Request("g"), 3, 2)
				If CLng(Right(Request("g"), 1)) = 1 Then
					gsExpectedDateTime = gsExpectedDateTime & " AM"
				Else
					gsExpectedDateTime = gsExpectedDateTime & " PM"
				End If
				
				If Not SubmitHoldOrder(Session("OrderID"), gsExpectedDateTime, CLng(Request("d")), Session("CustomerID"), Session("AddressID")) Then
					Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
				End If
				
%>
<!-- #Include Virtual="include2/clearorder.asp" -->
<%
				Response.Redirect("neworder.asp")
			Else
				If Session("OrderEdited") Then
					If Not Session("QuickMode") Then
						If Not PrintOrder(Session("StoreID"), Session("OrderID"), Session("NewOrder")) Then
							gbNeedPrinterAlert = TRUE
							If Session("NewOrder") Then
								gsLocalErrorMsg = "PRINT FAILURE, SAVING AS A HOLD ORDER!"
							Else
								gsLocalErrorMsg = "PRINT FAILURE, CANNOT PRINT THIS ORDER!"
							End If
						End If
					End If
				End If
				
				If Session("NewOrder") Then
					If gbNeedPrinterAlert Then
						If Not SubmitHoldOrder(Session("OrderID"), DateAdd("n", 15, Now), 15, Session("CustomerID"), Session("AddressID")) Then
							Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
						End If
					Else
						If Not SubmitOrder(Session("OrderID"), Session("CustomerID"), Session("AddressID")) Then
							Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
						End If
						
%>
<!-- #Include Virtual="include2/clearorder.asp" -->
<%
						Response.Redirect("neworder.asp")
					End If
				End If
			End If
		End If
	End If
%>
<!-- #Include Virtual="include2/clearorder.asp" -->
<%
Else
	gbShowMenuButtons = FALSE
	
	If Not IsNumeric(Request("o")) Then
		Response.Redirect("/default.asp")
	End If
	
	If Request("EditReason").Count > 0 Then
		If Len(Trim(Request("EditReason"))) > 0 Then
			Session("EditReason") = Trim(Request("EditReason"))
		End If
	End If
	
	If Len(Session("EditReason")) > 0 Then
		If Not SetOrderEdited(CLng(Request("o")), Session("EmpID"), Session("EditReason")) Then
			Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
		End If
	End If
	
	If Request("v").Count > 0 Then
		gbShowMenuButtons = TRUE
		
		If Not IsNumeric(Request("v")) Then
			Response.Redirect("neworder.asp")
		End If
		
		If Request("r").Count > 0 Then
			If CLng(Request("v")) = 4 Then
				If Not SetOrderPaymentType(CLng(Request("o")), CLng(Request("v")), "", CLng(Request("r"))) Then
					Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
				End If
				
				If Not IsAccountCollegeDebit(CLng(Request("r"))) Then
					If GetAccountContact(CLng(Request("r")), gsAccountName, gsPrimaryContactName, gsPrimaryContactEmail, gsSMSEmail) Then
						gsMailBody = "On " & Session("TransactionDate") & " order # " & CLng(Request("o")) & " for " & FormatCurrency(Session("OrderTotal")) & " was placed on the " & gsAccountName & " account "
						gsMailBody = gsMailBody & "at Store # " & Session("StoreID") & ", " & gsStoreAddress1 & ". If you have questions about your order call " & Left(gsStorePhone, 3) & "-" & Mid(gsStorePhone, 4, 3) & "-" & Mid(gsStorePhone, 7) & ". "
						gsMailBody = gsMailBody & "If you have billing questions call 866-720-8486."
						
						If Len(gsSMSEmail) > 0 Then
							SendMail gsSMTPFrom, gsSMSEmail, "", "", "", gsMailBody, "", NULL
						End If
						
						If Len(gsPrimaryContactEmail) > 0 Then
							gsMailBody = gsMailBody & CHR(13) & CHR(10) & CHR(13) & CHR(10)
							gsMailBody = gsMailBody & "This is an unattended e-mail address, please do not reply. If you have any questions please contact Vito's Franchising at 866-720-8486." & CHR(13) & CHR(10) & CHR(13) & CHR(10)
							SendMail gsSMTPFrom, gsPrimaryContactEmail, "", "", "Vito's Pizza Order On Account", gsMailBody, "", NULL
						End If
					Else
						Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
					End If
				End If
			Else
				If Not SetOrderPaymentType(CLng(Request("o")), CLng(Request("v")), Trim(Request("r")), 0) Then
					Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
				End If
			End If
		Else
			If Request("s").Count <> 0 Then
				If IsNumeric(Request("s")) Then
					If CLng(Request("s")) = 0 Then
						If Session("OrderTypeID") <> 1 Then
							If Not SetOrderPayment(CLng(Request("o")), CLng(Request("v")), "", 0, 0, Session("EmpID")) Then
								Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
							End If
						End If
					Else
						If Not SetOrderPaymentType(CLng(Request("o")), CLng(Request("v")), "", 0) Then
							Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
						End If
					End If
				Else
					If Not SetOrderPaymentType(CLng(Request("o")), CLng(Request("v")), "", 0) Then
						Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
					End If
				End If
			Else
				If Not SetOrderPaymentType(CLng(Request("o")), CLng(Request("v")), "", 0) Then
					Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
				End If
			End If
		End If
		
		If Session("OrderEdited") Then
			If Not Session("QuickMode") Then
				If Not PrintOrder(Session("StoreID"), CLng(Request("o")), Session("NewOrder")) Then
					gbNeedPrinterAlert = TRUE
					If Session("NewOrder") Then
						gsLocalErrorMsg = "PRINT FAILURE, SAVING AS A HOLD ORDER!"
					Else
						gsLocalErrorMsg = "PRINT FAILURE, CANNOT PRINT THIS ORDER!"
					End If
				End If
			End If
		End If
		
		If Session("NewOrder") Then
			If Request("a").Count <> 0 And IsNumeric(Request("a")) And Request("c").Count <> 0 And IsNumeric(Request("c")) Then
				gnCustomerID = CLng(Request("c"))
				gnAddressID = CLng(Request("a"))
			Else
				gnCustomerID = 1
				gnAddressID = 1
			End If
			
			If gbNeedPrinterAlert Then
				If Not SubmitHoldOrder(CLng(Request("o")), DateAdd("n", 15, Now), 15, gnCustomerID, gnAddressID) Then
					Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
				End If
			Else
				If Not SubmitOrder(CLng(Request("o")), gnCustomerID, gnAddressID) Then
					Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
				End If
			End If
		End If
		
'		If Request("s").Count <> 0 Then
'			If IsNumeric(Request("s")) Then
'				If CLng(Request("s")) = 0 Then
'					If Session("OrderTypeID") <> 1 Then
'						If Not SetOrderCompleted(CLng(Request("o"))) Then
'							Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
'						End If
'					End If
'				End If
'			End If
'		End If
		
%>
<!-- #Include Virtual="include2/clearorder.asp" -->
<%
		Response.Redirect("neworder.asp")
	End If
End If

Dim ganOrderIDs(), ganOrderTypeIDs(), gasOrderTypeDescriptions(), i, j, gnLeft
Dim ganLineIDs(), gasPhoneNumbers(), gasNames(), ganAreaCodes()

If Not GetStoreOrderTypes(Session("StoreID"), ganOrderTypeIDs, gasOrderTypeDescriptions) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If Not GetCallerID(Session("StoreID"), ganLineIDs, gasPhoneNumbers, gasNames) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If Not GetStoreAreaCodes(Session("StoreID"), ganAreaCodes) Then
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

var gnOrderType = 0;

var gbNameInLowerCase = false;
var gbFocusAreaCode = false;

function resetRedirect() {
	var loRedirectDiv;
	
	loRedirectDiv = ie4? eval("document.all.redirect") : document.getElementById("redirect");
    //	loRedirectDiv.innerHTML = <%=gnRedirectTime%>;
//	alert("gnOrderType = " + gnOrderType);
}

function disableEnterKey() {
	var loText, loDiv;
	
	if (event.keyCode == 13) {
		event.cancelBubble = true;
		event.returnValue = false;
		return false;
	}
}

function getDelivery() {
    var loDelivery,loPhone;
//    alert("Starting getDelivery()");

/*
    loBtnsDelivery = ie4? eval("document.all.btnDelivery") : document.getElementById('btnDelivery');
    loBtnsPhone = ie4? eval("document.all.btnPhone") : document.getElementById('btnPhone');

    loBtnsDelivery.class = "active";
    loBtnsPhone.class = "";
*/
    var loAreaCode, loPhone, loOrderTypeDiv, loPhoneDiv, loNameDiv;

    loAreaCode = ie4? eval("document.all.areacode") : document.getElementById('areacode');
    loPhone = ie4? eval("document.all.phone") : document.getElementById('phone');
    <%
    If Len(Session("CustomerPhone")) > 0 Then
    %>
        loAreaCode.value = "<%=Left(Session("CustomerPhone"), 3)%>";
    loPhone.value = "<%=Mid(Session("CustomerPhone"), 4, 3) & "-" & Right(Session("CustomerPhone"), 4)%>";
    <%
    Else
    If ganAreaCodes(0) = 0 Then
    %>
	loAreaCode.value = "419";
    <%
        Else
    %>
        loAreaCode.value = "<%=ganAreaCodes(0)%>";
    <%
        End If
    End If
    %>
	
    loMainDiv = ie4? eval("document.all.ordertypediv") : document.getElementById('ordertypediv');
    loPhoneDiv = ie4? eval("document.all.phonediv") : document.getElementById('phonediv');
    loNameDiv = ie4? eval("document.all.namediv") : document.getElementById('namediv');
	
    loMainDiv.style.visibility = "hidden";
    loNameDiv.style.visibility = "hidden";
    loPhoneDiv.style.visibility = "visible";
	
    resetRedirect();
}

function getPhone() {
    var loDelivery,loPhone;
//    alert("Starting getPhone() gnOrderType = " + gnOrderType);
    loDelivery = ie4? eval("document.all.btnDelivery") : document.getElementById('btnDelivery');
    loPhone = ie4? eval("document.all.btnPhone") : document.getElementById('btnPhone');
/*
    loDelivery.class = "";
    loPhone.class = "active";
*/
    var loAreaCode, loPhone, loOrderTypeDiv, loPhoneDiv, loNameDiv;

    loAreaCode = ie4? eval("document.all.areacode") : document.getElementById('areacode');
    loPhone = ie4? eval("document.all.phone") : document.getElementById('phone');
    <%
    If Len(Session("CustomerPhone")) > 0 Then
    %>
        loAreaCode.value = "<%=Left(Session("CustomerPhone"), 3)%>";
    loPhone.value = "<%=Mid(Session("CustomerPhone"), 4, 3) & "-" & Right(Session("CustomerPhone"), 4)%>";
    <%
    Else
    If ganAreaCodes(0) = 0 Then
%>
	    loAreaCode.value = "419";
    <%
        Else
    %>
        loAreaCode.value = "<%=ganAreaCodes(0)%>";
    <%
        End If
    End If
    %>
	
    loMainDiv = ie4? eval("document.all.ordertypediv") : document.getElementById('ordertypediv');
    loPhoneDiv = ie4? eval("document.all.phonediv") : document.getElementById('phonediv');
    loNameDiv = ie4? eval("document.all.namediv") : document.getElementById('namediv');
	
    loMainDiv.style.visibility = "hidden";
    loNameDiv.style.visibility = "hidden";
    loPhoneDiv.style.visibility = "visible";
	
    resetRedirect();
}

function setFocusAreaCode(pbAreaCode) {
	var loAreaCode, loPhone;
	
	loAreaCode = ie4? eval("document.all.areacode") : document.getElementById('areacode');
	loPhone = ie4? eval("document.all.phone") : document.getElementById('phone');
	
	if (pbAreaCode) {
		loAreaCode.style.backgroundColor = "#FFFFFF";
		loPhone.style.backgroundColor = "#CCCCCC";
	}
	else {
		loAreaCode.style.backgroundColor = "#CCCCCC";
		loPhone.style.backgroundColor = "#FFFFFF";
	}
	
	gbFocusAreaCode = pbAreaCode;
	
	resetRedirect();
}

function addToPhone(psDigit) {
	var loPhone, lsPhone;
	
	if (gbFocusAreaCode) {
		loPhone = ie4? eval("document.all.areacode") : document.getElementById('areacode');
	}
	else {
		loPhone = ie4? eval("document.all.phone") : document.getElementById('phone');
	}
	
	lsPhone = loPhone.value;
	if (gbFocusAreaCode) {
		if (lsPhone.length < 3) {
			lsPhone += psDigit;
			loPhone.value = lsPhone;
		}
		if (lsPhone.length == 3) {
			setFocusAreaCode(false);
		}
	}
	else {
		if (lsPhone.length < 8) {
			if (lsPhone.length == 3) {
				lsPhone = lsPhone + "-";
			}
			lsPhone += psDigit;
			loPhone.value = lsPhone;
		}
	}
	
	resetRedirect();
}

function setAreaCode(psDigit) {
	var loPhone, lsPhone;
	
	loPhone = ie4? eval("document.all.areacode") : document.getElementById('areacode');
	lsPhone = psDigit;
	loPhone.value = lsPhone;
	
	resetRedirect();
}

function clearAreaCode(psDigit) {
	var loPhone, lsPhone;
	
	loPhone = ie4? eval("document.all.areacode") : document.getElementById('areacode');
	loPhone.value = "";
	
	setFocusAreaCode(true);
}

function backspacePhone() {
	var loText, lsText;
	
	if (gbFocusAreaCode) {
		loText = ie4? eval("document.all.areacode") : document.getElementById('areacode');
	}
	else {
		loText = ie4? eval("document.all.phone") : document.getElementById('phone');
	}
	
	lsText = loText.value;
	if (lsText.length > 0) {
		lsText = lsText.substr(0, (lsText.length - 1));
		if ((!gbFocusAreaCode) && (lsText.length == 4)) {
			lsText = lsText.substr(0, (lsText.length - 1));
		}
		loText.value = lsText;
	}
	
	resetRedirect();
}

function cancelPhone() {
	var loPhone, loOrderTypeDiv, loPhoneDiv, loNameDiv;
	
	loPhone = ie4? eval("document.all.phone") : document.getElementById('phone');
	loPhone.value = "";
	
	loMainDiv = ie4? eval("document.all.ordertypediv") : document.getElementById('ordertypediv');
	loPhoneDiv = ie4? eval("document.all.phonediv") : document.getElementById('phonediv');
	loNameDiv = ie4? eval("document.all.namediv") : document.getElementById('namediv');
	
	loPhoneDiv.style.visibility = "hidden";
	loNameDiv.style.visibility = "hidden";
	loMainDiv.style.visibility = "visible";
	
	resetRedirect();
}

function getName() {
	var loName, loOrderTypeDiv, loPhoneDiv, loNameDiv;

	loName = ie4? eval("document.all.name") : document.getElementById('name');
	loName.value = "";
	
	loMainDiv = ie4? eval("document.all.ordertypediv") : document.getElementById('ordertypediv');
	loPhoneDiv = ie4? eval("document.all.phonediv") : document.getElementById('phonediv');
	loNameDiv = ie4? eval("document.all.namediv") : document.getElementById('namediv');
	
	loMainDiv.style.visibility = "hidden";
	loPhoneDiv.style.visibility = "hidden";
	loNameDiv.style.visibility = "visible";
	
	resetRedirect();
}

function addToName(psDigit) {
	var loName, lsName;
	
	loName = ie4? eval("document.all.name") : document.getElementById('name');
	lsName = loName.value;
	if (gbNameInLowerCase) {
		lsName += psDigit.toLowerCase();
	}
	else {
		lsName += psDigit;
	}
	loName.value = lsName;
	
	resetRedirect();
}

function backspaceName() {
	var loName, lsName;
	
	loName = ie4? eval("document.all.name") : document.getElementById('name');
	lsName = loName.value;
	if (lsName.length > 0) {
		lsName = lsName.substr(0, (lsName.length - 1));
		loName.value = lsName;
	}
	
	resetRedirect();
}

function cancelName() {
	var loName, loOrderTypeDiv, loPhoneDiv, loNameDiv;
	
	loName = ie4? eval("document.all.name") : document.getElementById('name');
	loName.value = "";
	
	loMainDiv = ie4? eval("document.all.ordertypediv") : document.getElementById('ordertypediv');
	loPhoneDiv = ie4? eval("document.all.phonediv") : document.getElementById('phonediv');
	loNameDiv = ie4? eval("document.all.namediv") : document.getElementById('namediv');
	
	loPhoneDiv.style.visibility = "hidden";
	loNameDiv.style.visibility = "hidden";
	loMainDiv.style.visibility = "visible";
	
	resetRedirect();
}

function clearName() {
	var loName;
	
	loName = ie4? eval("document.all.name") : document.getElementById('name');
	loName.value = "";
	
	resetRedirect();
}

function properName() {
	var loName, i, lsText, lbDoUpper;
	
	loName = ie4? eval("document.all.name") : document.getElementById('name');
	if (loName.value.length > 0) {
		lsText = loName.value.substr(0, 1).toUpperCase();
		lbDoUpper = false;
		for (i = 1; i < loName.value.length; i++) {
			if (loName.value.substr(i, 1) == " ") {
				lsText += " ";
				lbDoUpper = true;
			}
			else {
				if (lbDoUpper) {
					lsText += loName.value.substr(i, 1).toUpperCase();
					lbDoUpper = false;
				}
				else {
					lsText += loName.value.substr(i, 1).toLowerCase();
				}
			}
		}
		loName.value = lsText;
	}
	
	resetRedirect();
}

function shiftName() {
	var loObj;
	
	if (gbNameInLowerCase) {
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
	
	gbNameInLowerCase = !gbNameInLowerCase;
	
	resetRedirect();
}

function clrPhone() {
    var loPhone;
	
    loPhone = ie4? eval("document.all.phone") : document.getElementById('phone');
    loPhone.value = "";
}

function goNext() {
	var loName, lsName, loAreaCode, loPhone, lsValue, lsLocation;
	
	if (gnOrderType == 1 || gnOrderType == 2) {
		loAreaCode = ie4? eval("document.all.areacode") : document.getElementById('areacode');
		loPhone = ie4? eval("document.all.phone") : document.getElementById('phone');
		if (loAreaCode.value.length != 3)
			return false;
		if (loPhone.value.length != 8)
			return false;
		lsValue = loAreaCode.value + loPhone.value.substr(0, 3) + loPhone.value.substr(4);
		
		lsLocation = "customerfind.asp?t=" + gnOrderType.toString() + "&p=" + lsValue + "&r=3";
	}
	else {
		loName = ie4? eval("document.all.name") : document.getElementById('name');
		
		lsLocation = "unitselect.asp?t=" + gnOrderType.toString();
		lsName = loName.value;
		if (lsName.length > 0) {
			lsLocation = lsLocation + "&n=" + encodeURIComponent(lsName);
		}
	}
	
	window.location = lsLocation;
}

function goQuick() {
	window.location = "unitselect.asp?t=" + gnOrderType.toString() + "&q=yes"
}

function goCallerID(psPhone) {
//    alert("Starting goCallerID. gnOrderType = " + gnOrderType);
//    loDelivery = ie4? eval("document.all.btnDelivery") : document.getElementById('btnDelivery');
//    loPhone = ie4? eval("document.all.btnPhone") : document.getElementById('btnPhone');
	document.getElementById('tabs').style.display = 'block'
	document.getElementById('statusBlock').style.right = '0'
	document.getElementById('statusBlock').style.borderBottomLeftRadius = '6px'
	if(gnOrderType === 2) {
		document.getElementById('bookingType').innerHTML = 'Pickup'

	}

  var loName, lsName, loPhone, lsValue, lsLocation;
	
	if (isNaN(Number(psPhone))) {
		getPhone();
	} else {
		if (psPhone.length != 10) {
			loPhone = ie4? eval("document.all.phone") : document.getElementById('phone');
			loPhone.value = psPhone;
			getPhone();
		} else {
			lsLocation = "customerfind.asp?t=" + gnOrderType.toString() + "&p=" + psPhone + "&r=3";
			
			window.location = lsLocation;
		}
	}
}

function goBack() {
    window.location = "neworder.asp";
}

function iframeclick() {
document.getElementById("litmosiframe").contentWindow.document.body.onclick = function() {
		window.location = "/Litmos/Litmos.aspx?EmpID=<%=Session("EmpID")%>&StoreID=<%=Session("StoreID")%>";
    }
}
//-->
</script>
<script src="/include2/redirect2.js" type="text/javascript"></script>
</head>

<body onload="clockInit(clockLocalStartTime, clockServerStartTime); clockOnLoad(); redirect('/default.asp')" onunload="clockOnUnload()">

<div id="mainwindow" style="position: absolute; top: 0px; left: 0px; width=1010px; height: 768px; overflow: hidden;">
<table cellspacing="0" cellpadding="0" width="1010" height="764" border="1">

  <tr>
    <td valign="top" width="1010" height="764">
    <table cellspacing="0" cellpadding="0" width="1010">
      <tr height="72">
        <td valign="top" width="1010" height="72">
          <div id="statusBlock" style="right:314px;border-radius:0">
            <strong><%=IIf(gbTestMode,IIf(gbDevMode,"[DEV]","[TEST]"),"")%> Store <%=Session("StoreID")%></strong> |
            <b><%=Session("name")%></b>
            <div>
            <span id="ClockDate"><%=clockDateString(gDate)%></span> |
            <span id="ClockTime" onclick="clockToggleSeconds()"><%=clockTimeString(Hour(gDate), Minute(gDate), Second(gDate))%></span> |
						<span class="counter" id="redirect"><%=gnRedirectTime%></span></div> 

          </div>
          <ol id='tabs' style="display:none">
                <%if gnOrderPick = 1 Then %>
                    <li><a onclick='goBack();' title='Delivery' id="bookingType">Delivery</a></li>
                <% elseif gnOrderPick = 2 Then%>
                    <li><a onclick='goBack();' title='Delivery' id="bookingType">Pickup</a></li>
                <% else %>
                    <li><a onclick='goBack();' title='Delivery' id="bookingType">Delivery</a></li>
                <%end if%>
                <li class="active"><a onclick='getPhone();' title='Phone'>Phone</a></li>
                <li class="disabled">Address</li>
                <li class="disabled">Customer Name</li>
                <li class="disabled">Order</li>
                <li class="disabled">Notes</li>
            </ol>
        </td>
      </tr>
			<tr height="500">
				<td valign="top" width="1010">
					<div id="content-wrapper">
					<div id="content" style="position: relative; width: 1010px; height: 615px; overflow: auto;">
						<div id="ordertypediv" style="position: absolute; top: 0px; left: 0px; width: 1010px;">
							<table cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td valign="top" colspan="3">
<%
For i = 0 To UBound(ganOrderTypeIDs)
	Select Case ganOrderTypeIDs(i)
		Case 1, 2
%>
										<div style="position: relative; height: 100px;left:184px">
<!--										<div style="position: absolute; top: 0px; left: 0px;"><button style="width:125px; height: 100px;" onclick="gnOrderType = <%=ganOrderTypeIDs(i)%>; getPhone();"><%=gasOrderTypeDescriptions(i)%></button></div> -->
                							<div style="position: absolute; top: 0px; left: 0px;"><button style="width:125px; height: 100px;" ><%=gasOrderTypeDescriptions(i)%></button></div>

<%
			If ganLineIDs(0) <> 0 Then
				gnLeft = 125
				For j = 0 To UBound(ganLineIDs)
%>
					<div style="position: absolute; top: 0px; left: <%=gnLeft%>px;"><button style="width:125px; height: 100px;" onclick="gnOrderType = <%=ganOrderTypeIDs(i)%>; goCallerID('<%=gasPhoneNumbers(j)%>');">Line <%=ganLineIDs(j)%><br/><%=gasPhoneNumbers(j)%><br/><%=gasNames(j)%></button></div>
<%
					gnLeft = gnLeft + 125
				Next
			End If
%>
										</div>
										&nbsp;<br/>
									</td>
								</tr>
								<tr>
									<td valign="top">
<%
	End Select
Next

If Not gbShowMenuButtons Then
%>
										<button style="width:125px;" onclick="window.location = 'unitselect.asp';">Return To Order</button><br/>
<%
End If
%>
									</td>
									<td valign="top" align="right" width="310">
<%
If gbShowMenuButtons Then
%>
<%If Session("SecurityID") > 1 Then%>

<%
Dim sqlHoldOrders, RSHoldOrders

sqlHoldOrders = "SELECT OrderID, SessionID, IPAddress, EmpID, RefID, TransactionDate, SubmitDate, ReleaseDate, StoreID, CustomerID, CustomerName, CustomerPhone, AddressID, OrderTypeID, IsPaid, PaymentTypeID, PaymentReference, PaidDate, DeliveryCharge, DriverMoney, (Tax + Tax2) as Tax, Tip, OrderStatusID, OrderNotes, RADRAT FROM tblOrders WHERE (StoreID = "&Session("StoreID")&") AND OrderStatusID = 2 AND ReleaseDate <= dateadd(day, 7, getdate()) ORDER BY ReleaseDate"

Set RSHoldOrders = Conn.Execute(sqlHoldOrders)

If RSHoldOrders.BOF And RSHoldOrders.EOF Then 'None found
Else
%>
<%'------------------------------------ ON HOLD ORDERS BOX ---------------------------------------------------%>
					<table cellspacing="0" cellpadding="1" width="300" bgcolor="#c30a0d" border="0">
                      <tbody>
                        <tr>
                          <td><table width="100%" border="0" cellpadding="1" cellspacing="3" bgcolor="fbf3c5">
                              <tr>
                                <td align="left" bgcolor="#c30a0d"><b><font color="#ffffff">Hold Orders</font></b></td>
                              </tr>
<%Do While Not RSHoldOrders.EOF%>
                              <tr>
                                <td><input type="button" value="<%=RSHoldOrders("OrderID")%>" class="redbuttonthinwide"  onclick="document.location = '/ticket.asp?OrderID=<%=RSHoldOrders("OrderID")%>';" /> <%=RSHoldOrders("ReleaseDate")%>
								</td>
                              </tr>
<%RSHoldOrders.MoveNext
Loop
%>
                          </table>
						  </td>
                        </tr>
                      </tbody>
                    </table>
<%'------------------------------------ END ON HOLD ORDERS BOX ---------------------------------------------------%>
<% End If 'If Hold Orders Exist%>
<% End If 'Hold Orders Secuity%>

<%'------------------------------------ INFORMATION BOX ---------------------------------------------------%>
<%
Dim  SQL, RSEmp

SQL="Select * from tblEmployee where EmployeeID = " & Session("EmployeeID") &" and StoreID = "&Session("StoreID")&" and IsActive='True'"

Set RSEmp=Conn.Execute(SQL)

If RSEmp.EOF And RSEmp.BOF Then  'Not Found REDIRECT back to sign in page
	RSEmp.Close
	
	SQL="Select * from tblEmployee where EmployeeID = " & Session("EmployeeID") &" and SystemRoleID >= 4 and IsActive='True'"
	
	Set RSEmp=Conn.Execute(SQL)
End If

SQL="SELECT  ShiftID, EmpID, StoreID, StoreRoleID, Rate, PunchInType, PunchInTime, PunchOutType, PunchOutTime, ReportDate, Modified, RADRAT FROM tblShifts WHERE (EmpID = "&Session("EmpID")&") AND (PunchOutType IS NULL)"

Dim RSPunch

Set RSPunch=Conn.Execute(SQL)

Dim sqlActiveDriver, RSActiveDriver, OkToSignOut

sqlActiveDriver="SELECT DeliveryID FROM tblDelivery WHERE (EmpID = "&Session("EmpID")&") AND (CashedOutTime IS NULL)"

Set RSActiveDriver=Conn.Execute(sqlActiveDriver)

If RSActiveDriver.BOF And RSActiveDriver.EOF Then 'cashed out
	OkToSignOut = "True"
Else
	OkToSignOut = "False"
End If
%>

                      <table cellspacing="0" cellpadding="1" width="300" bgcolor="#006d31" border="0">
                        <tbody>
                          <tr>
                            <td>
							
							<table width="100%" border="0" cellpadding="1" cellspacing="3" bgcolor="fbf3c5">
                                <tr>
                                  <td align="left" bgcolor="#006D31"><b><font color="#ffffff">Information</font></b></td>
                                </tr>
                                <tr>
                                  <td align="left"><b>Driver Status - <%=RSEmp("DriverStatus")%></b></td>
                                </tr>
<%If Not RSPunch.EOF Then
If session("intAge") < 18 And Session("CurrentShiftID") <>"" Then 'reminder for break time if under the age of 18%>
                                <tr>
                                  <td align="left" style="height: 21px"><b><font color=red>You need to take a break by: <%=FormatDateTime(DateAdd("h",5,rspunch("RADRAT")),3)%></font></b></td>
                                </tr>
<%End If
End If
%>

<% If Session("SecurityID") > 1 Then 'need to loop through those that need break reminders

Dim sqlBreakReminder, RSBreakReminder, intAge
sqlBreakReminder="SELECT dbo.tblShifts.ShiftID, dbo.tblShifts.EmpID, dbo.tblShifts.StoreID, dbo.tblShifts.StoreRoleID, dbo.tblShifts.Rate, dbo.tblShifts.PunchInType, dbo.tblShifts.PunchInTime, dbo.tblShifts.PunchOutType, dbo.tblShifts.PunchOutTime, dbo.tblShifts.ReportDate, dbo.tblShifts.Modified, dbo.tblShifts.RADRAT, dbo.tblEmployee.Firstname, dbo.tblEmployee.LastName, dbo.tblEmployee.birthdate FROM dbo.tblShifts INNER JOIN dbo.tblEmployee ON dbo.tblShifts.EmpID = dbo.tblEmployee.EmpID WHERE (dbo.tblShifts.PunchOutType IS NULL) and (dbo.tblShifts.StoreID = "&Session("StoreID")&");"

Set RSBreakReminder=Conn.Execute(sqlBreakReminder)

Do While Not RSBreakReminder.EOF

   intAge = DateDiff("yyyy", RSBreakReminder("Birthdate"), Session("TransactionDate"))
'   If Now() < DateSerial(Year(now()), Month(RSEmp("Birthdate")), Day(RSEmp("Birthdate"))) Then
'     intAge = intAge - 1
'	End If
If intAge < 18 Then
%>
                                <tr>
                                  <td align="left"><b><font color=red><%=RSBreakReminder("FirstName")%>&nbsp;<%=RSBreakReminder("LastName")%>-break by: <%=FormatDateTime(DateAdd("h",5,RSBreakReminder("RADRAT")),3)%></font></b></td>
                                </tr>
<%
End If

RSBreakReminder.MoveNext
Loop

End If%>
                            </table>
							
							</td>
                          </tr>
                        </tbody>
                      </table>
<%'------------------------------------ END INFORMATION BOX ---------------------------------------------------%>

<% '----------------------------------------Message Processing --------------------------------

Dim SQLMessages, RSMessages
SQLMessages = "Select * from tblMessages where RecipientID = "&Session("EmpID")&" and Status <> 'Deleted'"
Set RSMessages=Conn.Execute(SQLMessages)
%>
<%If RSMessages.BOF And RSMessages.EOF Then

Else%>

                        <table cellspacing="0" cellpadding="1" width="300" bgcolor="#006d31" border="0">
                          <tbody>
                            <tr>
                              <td>
                                <table width="100%" border="0" cellpadding="1" cellspacing="3" bgcolor="fbf3c5">
                                  <tr>
                                    <td align="left" bgcolor="#006D31" style="height: 23px"><font color="#ffffff"><b>Messages</b></font></td>
                                  </tr>

<%
Do While Not RSMessages.EOF%>
<%
If RSMessages("Status") = "Unread" Or RSMessages("Status") = "Replied" Then 
	Response.Redirect("/readmessage.asp?MessageID="&RSMessages("MessageID"))
End If
%>

                                  <tr>
                                    <td align="left"><span class="smallblack10">

									<img src="/images/icon_<%=RSMessages("Status")%>.gif" alt="Icon" width="16" height="15" hspace="3" border="0" align="absmiddle" />

									<a href="/readmessage.asp?MessageID=<%=RSMessages("MessageID")%>" class="smallblack11">Message from <%=RSMessages("SenderName")%></a>, <%=FormatDateTime(RSMessages("RADRAT"),2)%></span></td>
                                  </tr>
<%
RSMessages.MoveNext
Loop
%>
                                  <tr>
                                    <td><span class="style3"><br />
                                        </span>
                                      <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                          <tr>
                                            <td height="1" colspan="3" bgcolor="#006D31" class="style3"></td>
                                          </tr>
                                          <tr>
                                            <td width="33%" class="style3"><div align="center"><img src="/images/icon_unread.gif" alt="Unread Icon" width="16" height="15" hspace="3" vspace="3" align="absmiddle" />Unread</div></td>
                                            <td width="33%" class="style3"><div align="center"><img src="/images/icon_read.gif" alt="Read Icon" width="16" height="15" hspace="3" vspace="3" align="absmiddle" />Read</div></td>
                                            <td width="33%" class="style3"><div align="center"><img src="/images/icon_replied.gif" alt="Replied Icon" width="16" height="15" hspace="3" vspace="3" align="absmiddle" />Replied</div></td>
                                          </tr>
                                      </table></td>
                                  </tr>
                              </table>
                                </td>
                            </tr>
                          </tbody>
                      </table>
					  
<%

End If
'----------------------------------------End Message Processing --------------------------------%>
<%
End If
%>
									</td>
									<td valign="top" align="right" width="310">
<%
If gbShowMenuButtons Then
%>
										<iframe src="/Litmos/LitmosIFrame.aspx?EmpID=<%=Session("EmpID")%>&StoreID=<%=Session("StoreID")%>" id="litmosiframe" name="litmosiframe" width="300" height="350" scrolling="no" onload="iframeclick()"></iframe>
<%
End If
%>
									</td>
								</tr>
							</table>
						</div>
						<div id="phonediv" align="center" style="position: absolute; top: 0px; left: 0px; width: 1010px; visibility: hidden;">

                            
							<table align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td valign="top" width="115">&nbsp;</td>
									<td valign="top">
										<table align="center" cellpadding="0" cellspacing="0">
                                            <tr></tr>
											<tr>
												<td colspan="3"><div align="center">
													<strong>AREA CODE</strong></div></td>
											</tr>
											<tr>
												<td colspan="3"><div align="center">
													<input type="text" id="areacode" autocomplete="off" onkeydown="disableEnterKey();" onfocus="setFocusAreaCode(true);" style="width: 100px; text-align: center; background-color: #cccccc;" /></div></td>
											</tr>
<%
If ganAreaCodes(0) = 0 Then
%>
											<tr>
												<td><button onclick="setAreaCode('419')">419</button></td>
												<td><button onclick="setAreaCode('567')">567</button></td>
												<td><button onclick="setAreaCode('734')">734</button></td>
											</tr>
<%
Else
%>
											<tr>
<%
	For i = 0 To UBound(ganAreaCodes)
		If i > 0 And i Mod 3 = 0 Then
%>
											</tr>
											<tr>
<%
		End If
%>
												<td><button onclick="setAreaCode('<%=ganAreaCodes(i)%>')"><%=ganAreaCodes(i)%></button></td>
<%
'		If i = 11 Then
'			Exit For
'		End If
	Next
	
'	If UBound(ganAreaCodes) Mod 3 <> 0 Then
		For i = (UBound(ganAreaCodes) Mod 3) To 1
%>
												<td>&nbsp;</td>
<%
		Next
'	End If
%>
											</tr>
<%
End If
%>
											<tr>
												<td colspan="3"><button style="width: 235px;" onclick="clearAreaCode()">Clear Area Code</button></td>
											</tr>
										</table>
									</td>
                                    <td valign="top" width="75"></td>
									<td valign="top">
										<table align="center" cellpadding="0" cellspacing="0">
											<tr>
												<td colspan="3"><div align="center">
													<strong>ENTER PHONE NUMBER</strong></div></td>
											</tr>
											<tr>
												<td colspan="3"><div align="center">
													<input type="text" id="phone" autocomplete="off" onkeydown="disableEnterKey();" onfocus="setFocusAreaCode(false);" style="width: 200px; text-align: center;" /></div></td>
											</tr>
											<tr>
												<td><button onclick="addToPhone('1')">1</button></td>
												<td><button onclick="addToPhone('2')">2</button></td>
												<td><button onclick="addToPhone('3')">3</button></td>
											</tr>
											<tr>
												<td><button onclick="addToPhone('4')">4</button></td>
												<td><button onclick="addToPhone('5')">5</button></td>
												<td><button onclick="addToPhone('6')">6</button></td>
											</tr>
											<tr>
												<td><button onclick="addToPhone('7')">7</button></td>
												<td><button onclick="addToPhone('8')">8</button></td>
												<td><button onclick="addToPhone('9')">9</button></td>
											</tr>
											<tr>
												<td><button onclick="cancelPhone()">Cancel</button></td>
												<td><button onclick="addToPhone('0')">0</button></td>
												<td><button onclick="backspacePhone()">Bksp</button></td>
											</tr>
											<tr>
												<td colspan="3"><button style="width: 235px;" onclick="clrPhone()">Clear Phone Number</button></td>
											</tr>
										</table>
									</td>
                                    <td valign="top" width="75"></td>
                                    <td valign="top">
                                        <br /><br /><br /><br />
                                        <table align="center" cellpadding="0" cellspacing="0">
                                            <tr>
                                                <td colspan="3"><button style="width: 235px;" onclick="">Add Extension</button></td>
                                            </tr>
                                            <tr>
                                                <td colspan="3"><button style="width: 235px;" onclick="">Add Department</button></td>
                                            </tr>
                                        </table>
                                    </td>
									<td valign="top" width="75">&nbsp;</td>
									<td valign="top" width="115">&nbsp;</td>
								</tr>
							</table>
                            <br /><br />
                            <table align="center" cellpadding="0" cellspacing="0">
                                <tr>
                                    <td colspan="3"><button style="width: 235px;" onclick="goBack()">Back</button></td><td colspan="3"><button style="width: 235px;" onclick="goNext()">Next</button></td>
                                </tr>
                            </table>
						</div>
						<div id="namediv" style="position: absolute; top: 0px; left: 0px; width: 1010px; visibility: hidden;">
							<table align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td><div align="center">
										<strong>ENTER NAME</strong></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<input type="text" id="name" style="width: 930px" autocomplete="off" onkeydown="disableEnterKey();" /></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button style="width: 233px;" onclick="addToName('TABLE #')">TABLE #</button>&nbsp;<button style="width: 233px;" onclick="cancelName()">Cancel</button>&nbsp;<button style="width: 233px;" onclick="clearName()">Clear</button>&nbsp;<button style="width: 233px;" onclick="goNext()">Done</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="addToName('+')">+</button><button onclick="addToName('!')">!</button><button onclick="addToName('@')">@</button><button onclick="addToName('#')">#</button><button onclick="addToName('$')">$</button><button onclick="addToName('%')">%</button><button onclick="addToName('^')">^</button><button onclick="addToName('&')">&amp;</button><button onclick="addToName('*')">*</button><button onclick="addToName('(')">(</button><button onclick="addToName(')')">)</button><button onclick="addToName(':')">:</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="addToName('=')">=</button><button onclick="addToName('1')">1</button><button onclick="addToName('2')">2</button><button onclick="addToName('3')">3</button><button onclick="addToName('4')">4</button><button onclick="addToName('5')">5</button><button onclick="addToName('6')">6</button><button onclick="addToName('7')">7</button><button onclick="addToName('8')">8</button><button onclick="addToName('9')">9</button><button onclick="addToName('0')">0</button><button onclick="addToName('?')">?</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="addToName('\'')">'</button><button name="key-q" id="key-q" onclick="addToName('Q')">Q</button><button name="key-w" id="key-w" onclick="addToName('W')">W</button><button name="key-e" id="key-e" onclick="addToName('E')">E</button><button name="key-r" id="key-r" onclick="addToName('R')">R</button><button name="key-t" id="key-t" onclick="addToName('T')">T</button><button name="key-y" id="key-y" onclick="addToName('Y')">Y</button><button name="key-u" id="key-u" onclick="addToName('U')">U</button><button name="key-i" id="key-i" onclick="addToName('I')">I</button><button name="key-o" id="key-o" onclick="addToName('O')">O</button><button name="key-p" id="key-p" onclick="addToName('P')">P</button><button onclick="addToName('/')">/</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="addToName('.')">.</button><button name="key-a" id="key-a" onclick="addToName('A')">A</button><button name="key-s" id="key-s" onclick="addToName('S')">S</button><button name="key-d" id="key-d" onclick="addToName('D')">D</button><button name="key-f" id="key-f" onclick="addToName('F')">F</button><button name="key-g" id="key-g" onclick="addToName('G')">G</button><button name="key-h" id="key-h" onclick="addToName('H')">H</button><button name="key-j" id="key-j" onclick="addToName('J')">J</button><button name="key-k" id="key-k" onclick="addToName('K')">K</button><button name="key-l" id="key-l" onclick="addToName('L')">L</button><button onclick="addToName(',')">,</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="shiftName()">Shift</button><button onclick="addToName('<')">&lt;</button><button name="key-z" id="key-z" onclick="addToName('Z')">Z</button><button name="key-x" id="key-x" onclick="addToName('X')">X</button><button name="key-c" id="key-c" onclick="addToName('C')">C</button><button name="key-v" id="key-v" onclick="addToName('V')">V</button><button name="key-b" id="key-b" onclick="addToName('B')">B</button><button name="key-n" id="key-n" onclick="addToName('N')">N</button><button name="key-m" id="key-m" onclick="addToName('M')">M</button><button onclick="addToName('>')">&gt;</button><button onclick="properName()">Proper</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button style="width: 150px;" onclick="addToName('1/2')">1/2</button>&nbsp;<button style="width: 600px;" onclick="addToName(' ')">Space</button>&nbsp;<button style="width: 150px;" onclick="backspaceName()">Bksp</button></div></td>
								</tr>
							</table>
						</div>
					</div>
				</td>
			</tr>
			<tr height="105">
				<td valign="top" colspan="2" width="1010" style="position: relative; top: -15px;">
					<div align="center">
<%
If gbShowMenuButtons Then
%>
<% 
If RSEmp("StoreID") = Session("StoreID") Then
If Session("CurrentShiftID") = "" Then 'Not Punched In

	'Need to determine if punched out or just on break
	Dim sqlBreakCheck, RSBreakCheck
	sqlBreakCheck = "Select TOP 1 dbo.tblShifts.* from tblShifts where EmpID = '"&Session("EmpID")&"' and ReportDate = '"&Session("TransactionDate")&"' Order By RADRAT Desc;"
	Set RSBreakCheck = Conn.Execute(sqlBreakCheck)
	If RSBreakCheck.BOF And RSBreakCheck.EOF Then
%>
		<a href="/punch.asp?PunchType=In" class="navs"><img src="/images/btn_punchIN.jpg" alt="Punch In" width="75" height="75" border="0" /></a>
<%
	Else
		If  RSBreakCheck("PunchOutType") = "Break" Then
%>
		<a href="/break.asp?Break=OFF&EmpID=<%=Session("EmpID")%>&OdometerIn=<%=RSBreakCheck("OdometerIn")%>"><img src="/images/btn_ReturnFrombreak.jpg" alt="Return From Break" width="75" height="75" border="0" /></a>
<%
		Else
%>
				<a href="/punch.asp?PunchType=In" class="navs"><img src="/images/btn_punchIN.jpg" alt="Punch In" width="75" height="75" border="0" /></a>
<%
		End If
	End If
Else

		Dim sqlDeliveries, RSDeliveries
		'Check to see if they are on a delivery
		SQLDeliveries="SELECT * FROM dbo.tblDelivery WHERE (EmpID = "&Session("EmpID")&") AND (DriverReturnTime IS NULL) and OrderID > 0"
		Set RSDeliveries=Conn.Execute(SQLDeliveries)

		If RSDeliveries.EOF And RSDeliveries.BOF Then 'Not out on delivery
			If OkToSignOut = "True" then%>
				<a href="/punch.asp?PunchType=Out&StoreRoleID=<%= RSPunch("StoreRoleID")%>"><img src="/images/btn_punchOUT.jpg" alt="Punch Out" width="75" height="75" border="0" /></a>
<%
			End If
%>
				<a href="/break.asp?Break=ON&EmpID=<%=Session("EmpID")%>"><img src="/images/btn_goOnbreak.jpg" alt="Go On Break" width="75" height="75" border="0" /></a>
<%
		Else 'Out on Delivery
%>
			<img src="/images/btn_outondelivery.jpg" alt="Out On Delivery" width="75" height="75" border="0" />
<%
		End If 
End If
End If
%>
<!--
   						<a href="/opentickets.asp"><img src="/images/btn_opentickets.jpg" alt="Open Tickets" width="75" height="75"  border="0"/></a>
						<a href="/orderlookup.asp"><img src="/images/btn_orderlookup.jpg" alt="Order Lookup" width="75" height="75" border="0" /></a>
						<a href="/ordersearch.asp"><img src="/images/btn_ordersearch.jpg" alt="Order Search" width="75" height="75" border="0" /></a>
						<a href="/driverdispatch.asp"><img src="/images/btn_driverdispatch.jpg" alt="Driver Dispatch" width="75" height="75" border="0" /></a>
-->
                        

<%If RSEmp("DriverStatus") = "Active" Or Session("SecurityID") > 1 Then%>
<!--						  <a href="/viewdrives.asp"><img src="/images/btn_drives.jpg" alt="Drives" width="75" height="75" border="0" /></a>-->
<%End If%>
<!--						<a href="/main.asp"><img src="/images/btn_mainmenu.jpg" alt="Main Menu" border="0" /></a><a href="/default.asp"><img src="/images/btn_signoff.jpg" alt="Sign Off" border="0" /></a><br />-->
<%
End If
%>
						<span class="orangetext">For technical assistance, please call 419.720.5050</span>
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
