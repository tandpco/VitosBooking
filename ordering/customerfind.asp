<%
Option Explicit
Response.buffer = TRUE

If Session("SecurityID") = "" Then
	Response.Redirect("/default.asp")
End If

If Request("p").Count = 0 Then
	Response.Redirect("neworder.asp")
End If

If Not IsNumeric(Request("p")) Then
	Response.Redirect("neworder.asp")
End If

If Request("t").Count = 0 Then
	Response.Redirect("neworder.asp")
End If

If Not IsNumeric(Request("t")) Then
	Response.Redirect("neworder.asp")
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
Dim gbShowMenuButtons
Dim gnOrderTypeID, gsPhone, ganCustomerIDs(), ganPrimaryAddressIDs(), ganAddressIDs(), ganStoreIDs(), gasAddresses(), gasNames(), i
Dim ganOrderIDs(), gsLocalErrorMsg
Dim gsAssignVisible, gsPostalVisible, gsPhoneVisible
Dim gasPostalCodes(), ganAreaCodes()

If Session("OrderID") = 0 Then
	gbShowMenuButtons = TRUE
Else
	gbShowMenuButtons = FALSE
End If

gsAssignVisible = "visible"
gsPostalVisible = "hidden"
gsPhoneVisible = "hidden"
Dim gbNeedPrinterAlert

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

gnOrderTypeID = CLng(Request("t"))

gsPhone = Request("p")
Session("CustomerPhone") = gsPhone

Session("ReturnURL") = "/ordering/customerfind.asp?t=" & gnOrderTypeID & "&p=" & gsPhone
Session("SaveURL") = "/ordering/addressfind.asp?t=" & gnOrderTypeID & "&p=" & gsPhone

If GetCustomersByPhone(gsPhone, ganCustomerIDs, ganPrimaryAddressIDs, ganAddressIDs, ganStoreIDs, gasAddresses, gasNames) Then
	If ganCustomerIDs(0) = 0 Then
		If Request("c").Count = 0 And Request("p2").Count > 0 Then
			Response.Redirect("/custmaint/newaddress.asp?t=" & gnOrderTypeID)
		Else
			gsAssignVisible = "hidden"
			gsPhoneVisible = "visible"
		End If
	End If
	If gnOrderTypeID = 1 Then
		If UBound(ganCustomerIDs) = 0 And ganCustomerIDs(0) <> 0 And ganAddressIDs(0) = 0 Then
			Response.Redirect("/custmaint/newaddress.asp?t=" & gnOrderTypeID & "&c=" & ganCustomerIDs(0))
		End If
	End If
Else
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If Not GetStorePostalCodes(Session("StoreID"), gasPostalCodes) Then
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

<%
If Request("c").Count <> 0 Then
	If IsNumeric(Request("c")) Then
		gsAssignVisible = "hidden"
		gsPostalVisible = "visible"
%>
var gnCustomerID = <%=Request("c")%>;
<%
	Else
%>
var gnCustomerID = 0;
<%
	End If
Else
%>
var gnCustomerID = 0;
<%
End If
%>
var gbHalfAddress = false;
var gbFocusAreaCode = false;

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

function getAddress() {
	var loText, loDiv;
	
	loText = ie4? eval("document.all.postalcode") : document.getElementById('postalcode');
	loText.value = "";
	loText = ie4? eval("document.all.streetnumber") : document.getElementById('streetnumber');
	loText.value = "";
	
	loDiv = ie4? eval("document.all.assigndiv") : document.getElementById('assigndiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.addressdiv") : document.getElementById('addressdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.namediv") : document.getElementById('namediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.phonediv") : document.getElementById('phonediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.phoneconfirmdiv") : document.getElementById('phoneconfirmdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.postalcodediv") : document.getElementById('postalcodediv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function addToPostalCode(psDigit) {
	var loText, lsText;
	
	loText = ie4? eval("document.all.postalcode") : document.getElementById('postalcode');
	lsText = loText.value;
	lsText += psDigit;
	loText.value = lsText;
	
	resetRedirect();
}

function backspacePostalCode() {
	var loText, lsText;
	
	loText = ie4? eval("document.all.postalcode") : document.getElementById('postalcode');
	lsText = loText.value;
	if (lsText.length > 0) {
		lsText = lsText.substr(0, (lsText.length - 1));
		loText.value = lsText;
	}
	
	resetRedirect();
}

function setPostalCode(psDigit) {
	var loPhone, lsPhone;
	
	loPhone = ie4? eval("document.all.postalcode") : document.getElementById('postalcode');
	loPhone.value = psDigit;
	
	resetRedirect();
}

function addToStreetNumber(psDigit) {
	var loText, lsText;
	
	loText = ie4? eval("document.all.streetnumber") : document.getElementById('streetnumber');
	lsText = loText.value;
	lsText += psDigit;
	loText.value = lsText;
	
	resetRedirect();
}

function backspaceStreetNumber() {
	var loText, lsText;
	
	loText = ie4? eval("document.all.streetnumber") : document.getElementById('streetnumber');
	lsText = loText.value;
	if (lsText.length > 0) {
		lsText = lsText.substr(0, (lsText.length - 1));
		loText.value = lsText;
	}
	
	resetRedirect();
}

function cancelPostalCode() {
	var loText, loDiv;
	
	loText = ie4? eval("document.all.postalcode") : document.getElementById('postalcode');
	loText.value = "";
	loText = ie4? eval("document.all.streetnumber") : document.getElementById('streetnumber');
	loText.value = "";
	
	gbHalfAddress = false;
	
	loDiv = ie4? eval("document.all.postalcodediv") : document.getElementById('postalcodediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.addressdiv") : document.getElementById('addressdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.namediv") : document.getElementById('namediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.phonediv") : document.getElementById('phonediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.phoneconfirmdiv") : document.getElementById('phoneconfirmdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.assigndiv") : document.getElementById('assigndiv');
	loDiv.style.visibility = "visible";
}

function getStreetLetter(pbHalfAddress) {
	var loText, lsPostalCode, lsStreetNumber, loDiv;
	
	loText = ie4? eval("document.all.postalcode") : document.getElementById('postalcode');
	lsPostalCode = loText.value;
	if (lsPostalCode.length != 5)
		return false;
	loText = ie4? eval("document.all.streetnumber") : document.getElementById('streetnumber');
	lsStreetNumber = loText.value;
	if (lsStreetNumber.length == 0)
		return false;
	
	gbHalfAddress = pbHalfAddress;
	
	loDiv = ie4? eval("document.all.postalstreetspan") : document.getElementById('postalstreetspan');
	loDiv.innerHTML = "<strong>Zip Code: " + lsPostalCode + " &nbsp; Street Number: " + lsStreetNumber
	if (gbHalfAddress) {
		loDiv.innerHTML = loDiv.innerHTML + " 1/2"
	}
	loDiv.innerHTML = loDiv.innerHTML + "</strong><br/>"
	
	loDiv = ie4? eval("document.all.assigndiv") : document.getElementById('assigndiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.postalcodediv") : document.getElementById('postalcodediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.namediv") : document.getElementById('namediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.phonediv") : document.getElementById('phonediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.phoneconfirmdiv") : document.getElementById('phoneconfirmdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.addressdiv") : document.getElementById('addressdiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function cancelAddress() {
	var loText, loDiv;
	
	loDiv = ie4? eval("document.all.assigndiv") : document.getElementById('assigndiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.addressdiv") : document.getElementById('addressdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.namediv") : document.getElementById('namediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.phonediv") : document.getElementById('phonediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.phoneconfirmdiv") : document.getElementById('phoneconfirmdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.postalcodediv") : document.getElementById('postalcodediv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function getName() {
	var loText, loDiv;
	
	loText = ie4? eval("document.all.name") : document.getElementById('name');
	loText.value = "";
	
	loDiv = ie4? eval("document.all.assigndiv") : document.getElementById('assigndiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.addressdiv") : document.getElementById('addressdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.postalcodediv") : document.getElementById('postalcodediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.phonediv") : document.getElementById('phonediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.phoneconfirmdiv") : document.getElementById('phoneconfirmdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.namediv") : document.getElementById('namediv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function addToName(psDigit) {
	var loName, lsName;
	
	loName = ie4? eval("document.all.name") : document.getElementById('name');
	lsName = loName.value;
	lsName += psDigit;
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
	var loDiv;
	
	loDiv = ie4? eval("document.all.addressdiv") : document.getElementById('addressdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.postalcodediv") : document.getElementById('postalcodediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.namediv") : document.getElementById('namediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.phonediv") : document.getElementById('phonediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.phoneconfirmdiv") : document.getElementById('phoneconfirmdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.assigndiv") : document.getElementById('assigndiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function getNewPhone() {
	var loDiv;
	
	loDiv = ie4? eval("document.all.assigndiv") : document.getElementById('assigndiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.addressdiv") : document.getElementById('addressdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.postalcodediv") : document.getElementById('postalcodediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.namediv") : document.getElementById('namediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.phoneconfirmdiv") : document.getElementById('phoneconfirmdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.phonediv") : document.getElementById('phonediv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function cancelNewPhone() {
<%
If gnOrderTypeID = 1 And ganCustomerIDs(0) = 0 Then
%>
	var loAreaCode, loPhone, lsValue;
	
	loAreaCode = ie4? eval("document.all.areacode") : document.getElementById('areacode');
	loPhone = ie4? eval("document.all.phone") : document.getElementById('phone');
	if (loPhone.value.length != 8)
		return false;
	lsValue = loAreaCode.value + loPhone.value.substr(0, 3) + loPhone.value.substr(4);
	
	window.location = "/custmaint/newaddress.asp?t=<%=gnOrderTypeID%>&p=" + lsValue;
<%
Else
%>
	var loDiv;
	
	loDiv = ie4? eval("document.all.addressdiv") : document.getElementById('addressdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.postalcodediv") : document.getElementById('postalcodediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.namediv") : document.getElementById('namediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.phonediv") : document.getElementById('phonediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.phoneconfirmdiv") : document.getElementById('phoneconfirmdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.assigndiv") : document.getElementById('assigndiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
<%
End If
%>
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
	
	window.location = "neworder.asp";
}

function goNewPhone() {
	var loName, lsName, loAreaCode, loPhone, lsValue, lsLocation;
	
	loAreaCode = ie4? eval("document.all.areacode") : document.getElementById('areacode');
	loPhone = ie4? eval("document.all.phone") : document.getElementById('phone');
	if (loAreaCode.value.length != 3)
		return false;
	if (loPhone.value.length != 8)
		return false;
	lsValue = loAreaCode.value + loPhone.value.substr(0, 3) + loPhone.value.substr(4);
	
	lsLocation = "customerfind.asp?t=<%=gnOrderTypeID%>&p=" + lsValue + "&p2=yes";
	
	window.location = lsLocation;
}

function goNext(psDigit) {
	var loText, lsLocation;
	
	lsLocation = "streetfind.asp?t=<%=Request("t")%>&z=";
	loText = ie4? eval("document.all.postalcode") : document.getElementById('postalcode');
	lsLocation = lsLocation + encodeURIComponent(loText.value) + "&y=";
	
	loText = ie4? eval("document.all.streetnumber") : document.getElementById('streetnumber');
	lsLocation = lsLocation + encodeURIComponent(loText.value) + "&x=" + psDigit;
	
	if (gbHalfAddress) {
		lsLocation = lsLocation + "&w=true";
	}
	else {
		lsLocation = lsLocation + "&w=false";
	}
	lsLocation = lsLocation + "&c=" + gnCustomerID.toString();
	
	window.location = lsLocation;
}

function goPickupNoCustomer() {
	var loText, lsLocation;
	
	loText = ie4? eval("document.all.name") : document.getElementById('name');
	lsText = loText.value;
	if (lsText.length == 0) {
		return false;
	}
	
	lsLocation = "unitselect.asp?t=<%=gnOrderTypeID%>&n=" + encodeURIComponent(lsText);
	
	window.location = lsLocation;
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
			<tr height="628">
				<td valign="top" width="1010">
					<div id="content" style="position: relative; width: 1010px; height: 618px; overflow: auto;">
						<div id="assigndiv" style="position: absolute; top: 0px; left: 0px; width: 1010px; visibility: <%=gsAssignVisible%>;">
<%
If gnOrderTypeID = 1 Then
%>
							<div align="center"><strong>SELECT DELIVERY LOCATION</strong></div><br/>
<%
Else
%>
							<div align="center"><strong>SELECT CUSTOMER FOR PICKUP</strong></div><br/>
<%
End If
%>
							<button style="width: 1010px;" onclick="window.location = '/custmaint/newaddress.asp?t=<%=gnOrderTypeID%>'">Add A New Customer</button>
<%
If gnOrderTypeID <> 1 Then
%>
							<button style="width: 1010px;" onclick="getName();">Enter Name and Begin Order</button>
<%
End If

If ganCustomerIDs(0) <> 0 Then
	Dim lnLastCustomer
	lnLastCustomer = ganCustomerIDs(0)
	For i = 0 To UBound(ganCustomerIDs)
		If 	lnLastCustomer <> ganCustomerIDs(i) Then
%>
							<button style="width: 1010px;" onclick="window.location = '/custmaint/newaddress.asp?c=<%=lnLastCustomer%>'"><%=gasNames(i - 1)%><br/>Add A New Address for this Customer</button>
<%
			lnLastCustomer = ganCustomerIDs(i)
		End If
		
		If ganAddressIDs(i) <> 0 Then
			If ganPrimaryAddressIDs(i) = ganAddressIDs(i) Then
				If gnOrderTypeID = 1 And ganStoreIDs(i) <> Session("StoreID") Then
%>
							<button style="width: 1010px;" onclick="window.location='otherstore.asp?t=<%=gnOrderTypeID%>&s=<%=ganStoreIDs(i)%>&c=<%=ganCustomerIDs(i)%>&a=<%=ganAddressIDs(i)%>'">PRIMARY ADDRESS<br/><%=gasNames(i)%><br/><%=gasAddresses(i)%></button>
<%
				Else
%>
							<button style="width: 1010px;" onclick="window.location='unitselect.asp?t=<%=gnOrderTypeID%>&c=<%=ganCustomerIDs(i)%>&a=<%=ganAddressIDs(i)%>'">PRIMARY ADDRESS<br/><%=gasNames(i)%><br/><%=gasAddresses(i)%></button>
<%
				End If
			Else
				If gnOrderTypeID = 1 And ganStoreIDs(i) <> Session("StoreID") Then
%>
							<button style="width: 1010px;" onclick="window.location='otherstore.asp?t=<%=gnOrderTypeID%>&s=<%=ganStoreIDs(i)%>&c=<%=ganCustomerIDs(i)%>&a=<%=ganAddressIDs(i)%>'"><%=gasNames(i)%><br/><%=gasAddresses(i)%></button>
<%
				Else
%>
							<button style="width: 1010px;" onclick="window.location='unitselect.asp?t=<%=gnOrderTypeID%>&c=<%=ganCustomerIDs(i)%>&a=<%=ganAddressIDs(i)%>'"><%=gasNames(i)%><br/><%=gasAddresses(i)%></button>
<%
				End If
			End If
		End If
	Next
%>
							<button style="width: 1010px;" onclick="window.location = '/custmaint/newaddress.asp?c=<%=ganCustomerIDs(UBound(ganCustomerIDs))%>'"><%=gasNames(UBound(ganCustomerIDs))%><br/>Add A New Address for this Customer</button>
<%
End If
%>
						</div>
						<div id="postalcodediv" style="position: absolute; top: 0px; left: 0px; width: 1010px; visibility: <%=gsPostalVisible%>;">
							<table align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="3"><div align="center">
										<strong>ENTER POSTAL CODE</strong></div></td>
									<td width="25">&nbsp;</td>
									<td colspan="3"><div align="center">
										<strong>ENTER STREET #</strong></div></td>
								</tr>
								<tr>
									<td colspan="3"><div align="center">
										<input type="text" id="postalcode" style="width: 200px" autocomplete="off" onkeydown="disableEnterKey();" value="<%=Request("z")%>" /></div></td>
									<td width="25">&nbsp;</td>
									<td colspan="3"><div align="center">
										<input type="text" id="streetnumber" style="width: 200px" autocomplete="off" onkeydown="disableEnterKey();" value="<%=Request("y")%>" /></div></td>
								</tr>
								<tr>
									<td><button onclick="addToPostalCode('1')">1</button></td>
									<td><button onclick="addToPostalCode('2')">2</button></td>
									<td><button onclick="addToPostalCode('3')">3</button></td>
									<td width="25">&nbsp;</td>
									<td><button onclick="addToStreetNumber('1')">1</button></td>
									<td><button onclick="addToStreetNumber('2')">2</button></td>
									<td><button onclick="addToStreetNumber('3')">3</button></td>
								</tr>
								<tr>
									<td><button onclick="addToPostalCode('4')">4</button></td>
									<td><button onclick="addToPostalCode('5')">5</button></td>
									<td><button onclick="addToPostalCode('6')">6</button></td>
									<td width="25">&nbsp;</td>
									<td><button onclick="addToStreetNumber('4')">4</button></td>
									<td><button onclick="addToStreetNumber('5')">5</button></td>
									<td><button onclick="addToStreetNumber('6')">6</button></td>
								</tr>
								<tr>
									<td><button onclick="addToPostalCode('7')">7</button></td>
									<td><button onclick="addToPostalCode('8')">8</button></td>
									<td><button onclick="addToPostalCode('9')">9</button></td>
									<td width="25">&nbsp;</td>
									<td><button onclick="addToStreetNumber('7')">7</button></td>
									<td><button onclick="addToStreetNumber('8')">8</button></td>
									<td><button onclick="addToStreetNumber('9')">9</button></td>
								</tr>
								<tr>
									<td>&nbsp;</td>
									<td><button onclick="addToPostalCode('0')">0</button></td>
									<td><button onclick="backspacePostalCode()">Bksp</button></td>
									<td width="25">&nbsp;</td>
									<td>&nbsp;</td>
									<td><button onclick="addToStreetNumber('0')">0</button></td>
									<td><button onclick="backspaceStreetNumber()">Bksp</button></td>
								</tr>
								<tr>
<%
Dim gnMaxI

If Len(gasPostalCodes(0)) > 0 Then
	If UBound(gasPostalCodes) > 2 Then
		gnMaxI = 2
	Else
		gnMaxI = UBound(gasPostalCodes)
	End If
	
	For i = 0 To gnMaxI
%>
									<td><button onclick="setPostalCode('<%=gasPostalCodes(i)%>')"><%=gasPostalCodes(i)%></button></td>
<%
	Next
	
	If gnMaxI < UBound(gasPostalCodes) Then
		gnMaxI = gnMaxI + 1
		
		Response.Write("<td width=""25"">&nbsp;</td>")
		
		For i = gnMaxI To UBound(gasPostalCodes)
			If i > gnMaxI And i Mod 3 = 0 Then
				Response.Write("<td width=""25"">&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr><tr>")
			End If
%>
									<td><button onclick="setPostalCode('<%=gasPostalCodes(i)%>')"><%=gasPostalCodes(i)%></button></td>
<%
		Next
	End If
	
	If (UBound(gasPostalCodes) + 1) Mod 3 > 0 Then
		For i = ((UBound(gasPostalCodes) + 1) Mod 3) To 2
%>
									<td>&nbsp;</td>
<%
		Next
	End If
Else
%>
									<td>&nbsp;</td>
									<td>&nbsp;</td>
									<td>&nbsp;</td>
<%
End If

If UBound(gasPostalCodes) < 3 Then
%>
									<td width="25">&nbsp;</td>
									<td>&nbsp;</td>
									<td>&nbsp;</td>
									<td>&nbsp;</td>
<%
End If
%>
								</tr>
								<tr>
									<td colspan="7">&nbsp;</td>
								</tr>
								<tr>
									<td colspan="7" align="center"><button onclick="cancelPostalCode();">Cancel</button>&nbsp;<button onclick="getStreetLetter(true);">1/2 Address</button>&nbsp;<button onclick="getStreetLetter(false);">Next</button></td>
								</tr>
							</table>
						</div>
						<div id="addressdiv" style="position: absolute; top: 0px; left: 0px; width: 1010px; visibility: hidden;">
							<table align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="11"><div align="center">
										<span id="postalstreetspan"></span><strong>ENTER FIRST LETTER OF STREET NAME</strong></div></td>
								</tr>
								<tr>
									<td><button onclick="goNext('1')">1</button></td>
									<td><button onclick="goNext('2')">2</button></td>
									<td><button onclick="goNext('3')">3</button></td>
									<td><button onclick="goNext('4')">4</button></td>
									<td><button onclick="goNext('5')">5</button></td>
									<td><button onclick="goNext('6')">6</button></td>
									<td><button onclick="goNext('7')">7</button></td>
									<td><button onclick="goNext('8')">8</button></td>
									<td><button onclick="goNext('9')">9</button></td>
									<td>&nbsp;</td>
									<td>&nbsp;</td>
								</tr>
								<tr>
									<td><button onclick="goNext('Q')">Q</button></td>
									<td><button onclick="goNext('W')">W</button></td>
									<td><button onclick="goNext('E')">E</button></td>
									<td><button onclick="goNext('R')">R</button></td>
									<td><button onclick="goNext('T')">T</button></td>
									<td><button onclick="goNext('Y')">Y</button></td>
									<td><button onclick="goNext('U')">U</button></td>
									<td><button onclick="goNext('I')">I</button></td>
									<td><button onclick="goNext('O')">O</button></td>
									<td><button onclick="goNext('P')">P</button></td>
									<td>&nbsp;</td>
								</tr>
								<tr>
									<td><button onclick="goNext('A')">A</button></td>
									<td><button onclick="goNext('S')">S</button></td>
									<td><button onclick="goNext('D')">D</button></td>
									<td><button onclick="goNext('F')">F</button></td>
									<td><button onclick="goNext('G')">G</button></td>
									<td><button onclick="goNext('H')">H</button></td>
									<td><button onclick="goNext('J')">J</button></td>
									<td><button onclick="goNext('K')">K</button></td>
									<td><button onclick="goNext('L')">L</button></td>
									<td>&nbsp;</td>
									<td>&nbsp;</td>
								</tr>
								<tr>
									<td><button onclick="goNext('Z')">Z</button></td>
									<td><button onclick="goNext('X')">X</button></td>
									<td><button onclick="goNext('C')">C</button></td>
									<td><button onclick="goNext('V')">V</button></td>
									<td><button onclick="goNext('B')">B</button></td>
									<td><button onclick="goNext('N')">N</button></td>
									<td><button onclick="goNext('M')">M</button></td>
									<td>&nbsp;</td>
									<td>&nbsp;</td>
									<td>&nbsp;</td>
									<td><button onclick="cancelAddress()">Back</button></td>
								</tr>
							</table>
						</div>
						<div id="namediv" style="position: absolute; top: 0px; left: 0px; width: 1010px; visibility: hidden;">
							<table align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="11"><div align="center">
										<strong>ENTER CUSTOMER NAME</strong></div></td>
								</tr>
								<tr>
									<td colspan="11"><div align="center">
										<input type="text" id="name" style="width: 800px" autocomplete="off" onkeydown="disableEnterKey();" /></div></td>
								</tr>
								<tr>
									<td><button onclick="addToName('1')">1</button></td>
									<td><button onclick="addToName('2')">2</button></td>
									<td><button onclick="addToName('3')">3</button></td>
									<td><button onclick="addToName('4')">4</button></td>
									<td><button onclick="addToName('5')">5</button></td>
									<td><button onclick="addToName('6')">6</button></td>
									<td><button onclick="addToName('7')">7</button></td>
									<td><button onclick="addToName('8')">8</button></td>
									<td><button onclick="addToName('9')">9</button></td>
									<td><button onclick="addToName('0')">0</button></td>
									<td><button onclick="backspaceName()">Bksp</button></td>
								</tr>
								<tr>
									<td><button onclick="addToName('Q')">Q</button></td>
									<td><button onclick="addToName('W')">W</button></td>
									<td><button onclick="addToName('E')">E</button></td>
									<td><button onclick="addToName('R')">R</button></td>
									<td><button onclick="addToName('T')">T</button></td>
									<td><button onclick="addToName('Y')">Y</button></td>
									<td><button onclick="addToName('U')">U</button></td>
									<td><button onclick="addToName('I')">I</button></td>
									<td><button onclick="addToName('O')">O</button></td>
									<td><button onclick="addToName('P')">P</button></td>
									<td><button onclick="addToName('#')">#</button></td>
								</tr>
								<tr>
									<td><button onclick="addToName('A')">A</button></td>
									<td><button onclick="addToName('S')">S</button></td>
									<td><button onclick="addToName('D')">D</button></td>
									<td><button onclick="addToName('F')">F</button></td>
									<td><button onclick="addToName('G')">G</button></td>
									<td><button onclick="addToName('H')">H</button></td>
									<td><button onclick="addToName('J')">J</button></td>
									<td><button onclick="addToName('K')">K</button></td>
									<td><button onclick="addToName('L')">L</button></td>
									<td><button onclick="addToName(''')">'</button></td>
									<td><button onclick="cancelName()">Back</button></td>
								</tr>
								<tr>
									<td><button onclick="addToName('Z')">Z</button></td>
									<td><button onclick="addToName('X')">X</button></td>
									<td><button onclick="addToName('C')">C</button></td>
									<td><button onclick="addToName('V')">V</button></td>
									<td><button onclick="addToName('B')">B</button></td>
									<td><button onclick="addToName('N')">N</button></td>
									<td><button onclick="addToName('M')">M</button></td>
									<td><button onclick="addToName('.')">.</button></td>
									<td><button onclick="addToName(',')">,</button></td>
									<td><button onclick="addToName(' ')">Space</button></td>
									<td><button onclick="goPickupNoCustomer()">OK</button></td>
								</tr>
							</table>
						</div>
						<div id="phonediv" style="position: absolute; top: 0px; left: 0px; width: 1010px; visibility: <%=gsPhoneVisible%>;">
							<table align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td valign="top">
										<table align="center" cellpadding="0" cellspacing="0">
											<tr>
												<td colspan="3"><div align="center">
													<strong>AREA CODE</strong></div></td>
											</tr>
											<tr>
												<td colspan="3"><div align="center">
													<input type="text" id="areacode" autocomplete="off" onkeydown="disableEnterKey();" onfocus="setFocusAreaCode(true);" style="width: 100px; text-align: center; background-color: #cccccc;" value="<%=Left(gsPhone, 3)%>" /></div></td>
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
													<input type="text" id="phone" autocomplete="off" onkeydown="disableEnterKey();" onfocus="setFocusAreaCode(false);" style="width: 200px; text-align: center;" value="<%=Mid(gsPhone, 4, 3) & "-" & Mid(gsPhone, 7)%>"/></div></td>
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
												<td colspan="3"><button style="width: 235px;" onclick="goNewPhone()">OK</button></td>
											</tr>
										</table>
									</td>
									<td valign="top" width="75">&nbsp;</td>
									<td valign="top" width="235">&nbsp;</td>
								</tr>
							</table>
						</div>
						<div id="phoneconfirmdiv" style="position: absolute; top: 50px; left: 75px; width: 860px; height: 400px; z-index: 20; background: #fbf3c5; border: medium #000000 solid; visibility: <%=gsPhoneVisible%>;">
							<table align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td valign="top" width="25">&nbsp;</td>
									<td valign="top">&nbsp;</td>
									<td valign="top" width="25">&nbsp;</td>
								</tr>
								<tr>
									<td valign="top" width="25">&nbsp;</td>
									<td valign="top"><strong>This phone number does not match anything in the database. Ask the customer if they are sure that this is the correct number. Is this the correct number?</strong></td>
									<td valign="top" width="25">&nbsp;</td>
								</tr>
								<tr>
									<td valign="top" width="25">&nbsp;</td>
									<td valign="top">&nbsp;</td>
									<td valign="top" width="25">&nbsp;</td>
								</tr>
								<tr>
									<td valign="top" width="25">&nbsp;</td>
									<td valign="top">&nbsp;</td>
									<td valign="top" width="25">&nbsp;</td>
								</tr>
								<tr>
									<td valign="top" width="25">&nbsp;</td>
									<td valign="top" align="center"><button style="width: 235px;" onclick="cancelNewPhone()">Yes</button>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;<button style="width: 235px;" onclick="getNewPhone()">No</button></td>
									<td valign="top" width="25">&nbsp;</td>
								</tr>
							</table>
						</div>
					</div>
				</td>
			</tr>
			<tr height="105">
				<td valign="top" colspan="2" width="1010">
					<div align="center">
<%
If gbShowMenuButtons Then
%>
						<a href="/main.asp"><img src="/images/btn_mainmenu.jpg" alt="Main Menu" border="0" /></a><a href="/default.asp"><img src="/images/btn_signoff.jpg" alt="Sign Off" border="0" /></a><br />
<%
End If
%>
						<span class="orangetext">For technical assistance, please call 419.720.5050</span>
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
