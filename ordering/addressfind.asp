<%
Option Explicit
Response.buffer = TRUE

If Session("SecurityID") = "" Then
	Response.Redirect("/default.asp")
End If

If Request("b").Count = 0 Then
	Response.Redirect("neworder.asp")
End If

If Request("z").Count = 0 Then
	Response.Redirect("neworder.asp")
Else
	If Not IsNumeric(Request("z")) Then
		Response.Redirect("neworder.asp")
	End If
End If

If Request("c").Count <> 0 Then
	If Not IsNumeric(Request("c")) Then
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
Dim gnOrderTypeID, gsPhone, gsAddress1, gsPostalCode, gnStoreID1, gnStoreID2, gsAddress, gnCustomerID, gsName
Dim gsAddress2, gsCity, gsState, gdDeliveryCharge, gdDriverMoney
Dim gnAddressID, gsAddressNotes
Dim ganCustomerIDs(), gasNames()
Dim gsEMail, gsFirstName, gsLastName, gdtBirthdate, gnPrimaryAddressID, gsHomePhone, gsCellPhone, gsWorkPhone, gsFAXPhone, gbIsEMailList, gbIsTextList
Dim i, gbAssignDiv, gbNameDiv
Dim ganOrderIDs(), gsLocalErrorMsg
Dim gbNeedPrinterAlert
Dim gbIsManual

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
gsPhone = Session("CustomerPhone")
gsAddress1 = Request("b")
gsPostalCode = Request("z")

If Request("c").Count <> 0 Then
	gnCustomerID = CLng(Request("c"))
Else
	gnCustomerID = 0
End If

If Request("Manual").Count > 0 Then
	If Request("Manual") = "Yes" Then
		gbIsManual = TRUE
	Else
		gbIsManual = FALSE
	End If
Else
	gbIsManual = FALSE
End If

gnStoreID1 = GetStoreByAddress(gsPostalCode, gsAddress1, gsAddress2, gsCity, gsState, gdDeliveryCharge, gdDriverMoney)

'If gnOrderTypeID = 1 And gnStoreID1 = -1 Or (gsCity = "UNKNOWN CITY" And gsState = "US") Then
'	Response.Redirect("/error.asp?err=" & Server.URLEncode("There are no stores that deliver to that address."))
'Else
'	If gnStoreID1 = -1 Then
'		gnStoreID1 = 0
'	End If
'	
'	If Not LookupAddress(gsAddress1, gsAddress2, gsCity, gsState, gsPostalCode, gnAddressID, gnStoreID2, gsAddressNotes) Then
'		Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
'	End If
'	
'	If gnAddressID = 0 Then
'		gnStoreID2 = gnStoreID1
'		gnAddressID = AddAddress(gnStoreID1, gsAddress1, gsAddress2, gsCity, gsState, gsPostalCode, "", FALSE)
'		If gnAddressID = 0 Then
'			Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
'		End If
'	End If
'End If
If gnStoreID1 = -1 Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
Else
	If gsCity = "UNKNOWN CITY" And gsState = "US" Then
		Response.Redirect("otherstore.asp?t=" & gnOrderTypeID & "&s=0&a=0&c=" & gnCustomerID & "&b=" & Server.URLEncode(gsAddress1) & "&z=" & gsPostalCode)
	Else
		If gnStoreID1 = -1 Then
			gnStoreID1 = 0
		End If
		
		If Not LookupAddress(gsAddress1, gsAddress2, gsCity, gsState, gsPostalCode, gnAddressID, gnStoreID2, gsAddressNotes) Then
			Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
		End If
		
		If gnAddressID = 0 Then
			gnStoreID2 = gnStoreID1
			gnAddressID = AddAddress(gnStoreID1, gsAddress1, gsAddress2, gsCity, gsState, gsPostalCode, "", gbIsManual)
			If gnAddressID = 0 Then
				Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
			End If
		End If
	End If
End If

If Len(gsAddress2) = 0 Then
	gsAddress = gsAddress1
Else
	gsAddress = gsAddress1 & " #" & gsAddress2
End If

If gnCustomerID = 0 Then
	If Not GetCustomersByAddress(gnAddressID, ganCustomerIDs, gasNames) Then
		Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
	End If
	
	If ganCustomerIDs(0) = 0 Then
		gbAssignDiv = "hidden"
		gbNameDiv = "visible"
	Else
		gbAssignDiv = "visible"
		gbNameDiv = "hidden"
	End If
Else
	ReDim ganCustomerIDs(0)
	ganCustomerIDs(0) = 0
	
	If Not GetCustomerDetails(gnCustomerID, gsEMail, gsFirstName, gsLastName, gdtBirthdate, gnPrimaryAddressID, gsHomePhone, gsCellPhone, gsWorkPhone, gsFAXPhone, gbIsEMailList, gbIsTextList) Then
		Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
	End If
	
	If Len(gsLastName) = 0 Then
		gsName = gsFirstName
	Else
		gsName = gsFirstName & " " & gsLastName
	End If
	
	gbAssignDiv = "visible"
	gbNameDiv = "hidden"
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

var gnCustomerID = 0;
var gnPhoneType = 0;

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

function getPhoneType() {
	var loText, loDiv;
	
	if (gnCustomerID == 0) {
		loText = ie4? eval("document.all.email") : document.getElementById('email');
		if (!echeck(loText.value)) {
			return false;
		}
	}
	
	loDiv = ie4? eval("document.all.assigndiv") : document.getElementById('assigndiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.namediv") : document.getElementById('namediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.emaildiv") : document.getElementById('emaildiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.phonetypediv") : document.getElementById('phonetypediv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function getName() {
	var loText, loDiv;
	
	loText = ie4? eval("document.all.name") : document.getElementById('name');
	loText.value = "";
	
	loDiv = ie4? eval("document.all.assigndiv") : document.getElementById('assigndiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.emaildiv") : document.getElementById('emaildiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.phonetypediv") : document.getElementById('phonetypediv');
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
//	window.location = "neworder.asp";
	var loText, loDiv;
	
<%
If gnCustomerID = 0 And ganCustomerIDs(0) = 0 Then
%>
//	window.location = "newcustomer.asp?t=<%=Request("t")%>&p=<%=Session("CustomerPhone")%>&z=<%=Request("z")%>&y=<%=Left(Request("b"), (InStr(Request("b"), " ") - 1))%>";
	window.location = "customerfind.asp?t=<%=Request("t")%>&p=<%=Session("CustomerPhone")%>";
<%
Else
%>
	loText = ie4? eval("document.all.name") : document.getElementById('name');
	loText.value = "";
	
	loDiv = ie4? eval("document.all.namediv") : document.getElementById('namediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.emaildiv") : document.getElementById('emaildiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.phonetypediv") : document.getElementById('phonetypediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.assigndiv") : document.getElementById('assigndiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
<%
End If
%>
}

function getEmail() {
	var loText, loDiv;
	
	loText = ie4? eval("document.all.name") : document.getElementById('name');
	if (loText.value.length == 0) {
		return false;
	}
	
	loText = ie4? eval("document.all.email") : document.getElementById('email');
	loText.value = "";
	
	loDiv = ie4? eval("document.all.assigndiv") : document.getElementById('assigndiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.namediv") : document.getElementById('namediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.phonetypediv") : document.getElementById('phonetypediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.emaildiv") : document.getElementById('emaildiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function addToEmail(psDigit) {
	var loName, lsName;
	
	loName = ie4? eval("document.all.email") : document.getElementById('email');
	lsName = loName.value;
	lsName += psDigit;
	loName.value = lsName;
	
	resetRedirect();
}

function backspaceEmail() {
	var loName, lsName;
	
	loName = ie4? eval("document.all.email") : document.getElementById('email');
	lsName = loName.value;
	if (lsName.length > 0) {
		lsName = lsName.substr(0, (lsName.length - 1));
		loName.value = lsName;
	}
	
	resetRedirect();
}

function cancelEmail() {
//	window.location = "neworder.asp";
	var loText, loDiv;
	
	loText = ie4? eval("document.all.email") : document.getElementById('email');
	loText.value = "";
	
	loDiv = ie4? eval("document.all.assigndiv") : document.getElementById('assigndiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.emaildiv") : document.getElementById('emaildiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.phonetypediv") : document.getElementById('phonetypediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.namediv") : document.getElementById('namediv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function echeck(str) {
	var at="@"
	var dot="."
	var lat=str.indexOf(at)
	var lstr=str.length
	var ldot=str.indexOf(dot)
	
	if (lstr > 0) {
		// Ensure @ is present and not at beginning or end
		if (str.indexOf(at)==-1 || str.indexOf(at)==0 || str.indexOf(at)==lstr){
		   alert("Invalid E-mail ID")
		   return false
		}
		
		// Ensure . is present and not a beginning or at end
		if (str.indexOf(dot)==-1 || str.indexOf(dot)==0 || str.indexOf(dot)==lstr){
		    alert("Invalid E-mail ID")
		    return false
		}
		
		// Ensure . is present after @
		 if (str.indexOf(at,(lat+1))!=-1){
		    alert("Invalid E-mail ID")
		    return false
		 }
		
		// Ensure . is not immeditately before or after the @
		 if (str.substring(lat-1,lat)==dot || str.substring(lat+1,lat+2)==dot){
		    alert("Invalid E-mail ID")
		    return false
		 }
		
		// Ensure . is present after @
		 if (str.indexOf(dot,(lat+2))==-1){
		    alert("Invalid E-mail ID")
		    return false
		 }
		
		// Ensure no spaces
		 if (str.indexOf(" ")!=-1){
		    alert("Invalid E-mail ID")
		    return false
		 }
	}
	
	return true					
}

function goNext() {
	var loText, loText2, lsLocation;

<%
If gnCustomerID = 0 Then
%>	
	if (gnCustomerID == 0) {
		loText = ie4? eval("document.all.name") : document.getElementById('name');
		loText2 = ie4? eval("document.all.email") : document.getElementById('email');
		if ((<%=gnOrderTypeID%> == 1) && (<%=gnStoreID1%> != <%=Session("StoreID")%>)) {
			lsLocation = "otherstore.asp?t=<%=gnOrderTypeID%>&s=<%=gnStoreID1%>&a=<%=gnAddressID%>&h=" + gnPhoneType.toString() + "&n=" + encodeURIComponent(loText.value) + "&em=" + encodeURIComponent(loText2.value) + "&newcustomer=yes";
		}
		else {
			lsLocation = "unitselect.asp?t=<%=gnOrderTypeID%>&a=<%=gnAddressID%>&h=" + gnPhoneType.toString() + "&n=" + encodeURIComponent(loText.value) + "&em=" + encodeURIComponent(loText2.value) + "&newcustomer=yes";
		}
	}
	else {
		if ((<%=gnOrderTypeID%> == 1) && (<%=gnStoreID1%> != <%=Session("StoreID")%>)) {
			lsLocation = "otherstore.asp?t=<%=gnOrderTypeID%>&s=<%=gnStoreID1%>&c=" + gnCustomerID.toString() + "&a=<%=gnAddressID%>&h=" + gnPhoneType.toString() + "&assigncustomerphone=yes";
		}
		else {
			lsLocation = "unitselect.asp?t=<%=gnOrderTypeID%>&c=" + gnCustomerID.toString() + "&a=<%=gnAddressID%>&h=" + gnPhoneType.toString() + "&assigncustomerphone=yes";
		}
	}
<%
Else
%>
	if ((<%=gnOrderTypeID%> == 1) && (<%=gnStoreID1%> != <%=Session("StoreID")%>)) {
		lsLocation = "otherstore.asp?t=<%=gnOrderTypeID%>&s=<%=gnStoreID1%>&c=<%=gnCustomerID%>&a=<%=gnAddressID%>&assigncustomeraddress=yes";
	}
	else {
		lsLocation = "unitselect.asp?t=<%=gnOrderTypeID%>&c=<%=gnCustomerID%>&a=<%=gnAddressID%>&assigncustomeraddress=yes";
	}
<%
End If
%>
	
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
						<div id="assigndiv" style="position: absolute; top: 0px; left: 0px; width: 1010px; visibility: <%=gbAssignDiv%>;">
<%
If gnOrderTypeID = 1 Then
%>
							<div align="center"><strong>SELECT DELIVERY CUSTOMER</strong></div><br/>
<%
Else
%>
							<div align="center"><strong>SELECT CUSTOMER FOR PICKUP</strong></div><br/>
<%
End If

If gnCustomerID = 0 Then
	If ganCustomerIDs(0) <> 0 Then
		For i = 0 To UBound(ganCustomerIDs)
%>
							<button style="width: 1010px;" onclick="gnCustomerID = <%=ganCustomerIDs(i)%>; getPhoneType();">Assign New Phone Number To: <%=gasNames(i)%><br/><%=gsAddress%></button>
<%
		Next
	End If
%>
							<button style="width: 1010px;" onclick="getName();">Enter New Customer Name<br/><%=gsAddress%></button>
<%
Else
%>
							<button style="width: 1010px;" onclick="goNext();">Assign Address To: <%=gsName%><br/><%=gsAddress%></button>
<%
End If
%>
						</div>
						<div id="phonetypediv" style="position: absolute; top: 0px; left: 0px; width: 1010px; visibility: hidden;">
							<div align="center"><strong>SELECT PHONE NUMBER TYPE</strong></div><br/>
							<button style="width: 1010px;" onclick="gnPhoneType = 0; goNext();">Home</button>
							<button style="width: 1010px;" onclick="gnPhoneType = 1; goNext();">Cell</button>
							<button style="width: 1010px;" onclick="gnPhoneType = 2; goNext();">Work</button>
							<button style="width: 1010px;" onclick="gnPhoneType = 3; goNext();">FAX</button>
						</div>
						<div id="namediv" style="position: absolute; top: 0px; left: 0px; width: 1010px; visibility: <%=gbNameDiv%>;">
							<table align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="11"><div align="center">
										<strong>ENTER NEW CUSTOMER NAME - REQUIRED</strong></div></td>
								</tr>
								<tr>
									<td colspan="11"><div align="center">
										<input type="text" id="name" style="width: 800px" autocomplete="off" onkeydown="disableEnterKey();" value="<%=Session("CustomerName")%>"/></div></td>
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
									<td><button onclick="getEmail()">OK</button></td>
								</tr>
							</table>
						</div>
						<div id="emaildiv" style="position: absolute; top: 0px; left: 0px; width: 1010px; visibility: hidden;">
							<table align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="11"><div align="center">
										<strong>ENTER NEW CUSTOMER E-MAIL ADDRESS</strong></div></td>
								</tr>
								<tr>
									<td colspan="11"><div align="center">
										<input type="text" id="email" style="width: 800px" autocomplete="off" onkeydown="disableEnterKey();" /></div></td>
								</tr>
								<tr>
									<td><button onclick="addToEmail('1')">1</button></td>
									<td><button onclick="addToEmail('2')">2</button></td>
									<td><button onclick="addToEmail('3')">3</button></td>
									<td><button onclick="addToEmail('4')">4</button></td>
									<td><button onclick="addToEmail('5')">5</button></td>
									<td><button onclick="addToEmail('6')">6</button></td>
									<td><button onclick="addToEmail('7')">7</button></td>
									<td><button onclick="addToEmail('8')">8</button></td>
									<td><button onclick="addToEmail('9')">9</button></td>
									<td><button onclick="addToEmail('0')">0</button></td>
									<td><button onclick="backspaceEmail()">Bksp</button></td>
								</tr>
								<tr>
									<td><button onclick="addToEmail('Q')">Q</button></td>
									<td><button onclick="addToEmail('W')">W</button></td>
									<td><button onclick="addToEmail('E')">E</button></td>
									<td><button onclick="addToEmail('R')">R</button></td>
									<td><button onclick="addToEmail('T')">T</button></td>
									<td><button onclick="addToEmail('Y')">Y</button></td>
									<td><button onclick="addToEmail('U')">U</button></td>
									<td><button onclick="addToEmail('I')">I</button></td>
									<td><button onclick="addToEmail('O')">O</button></td>
									<td><button onclick="addToEmail('P')">P</button></td>
									<td><button onclick="addToEmail('#')">#</button></td>
								</tr>
								<tr>
									<td><button onclick="addToEmail('A')">A</button></td>
									<td><button onclick="addToEmail('S')">S</button></td>
									<td><button onclick="addToEmail('D')">D</button></td>
									<td><button onclick="addToEmail('F')">F</button></td>
									<td><button onclick="addToEmail('G')">G</button></td>
									<td><button onclick="addToEmail('H')">H</button></td>
									<td><button onclick="addToEmail('J')">J</button></td>
									<td><button onclick="addToEmail('K')">K</button></td>
									<td><button onclick="addToEmail('L')">L</button></td>
									<td><button onclick="addToEmail(''')">'</button></td>
									<td><button onclick="addToEmail('^')">^</button></td>
								</tr>
								<tr>
									<td><button onclick="addToEmail('Z')">Z</button></td>
									<td><button onclick="addToEmail('X')">X</button></td>
									<td><button onclick="addToEmail('C')">C</button></td>
									<td><button onclick="addToEmail('V')">V</button></td>
									<td><button onclick="addToEmail('B')">B</button></td>
									<td><button onclick="addToEmail('N')">N</button></td>
									<td><button onclick="addToEmail('M')">M</button></td>
									<td><button onclick="addToEmail('.')">.</button></td>
									<td><button onclick="addToEmail('_')">_</button></td>
									<td><button onclick="addToEmail('!')">!</button></td>
									<td><button onclick="addToEmail('`')">`</button></td>
								</tr>
								<tr>
									<td><button onclick="addToEmail('@')">@</button></td>
									<td><button onclick="addToEmail('$')">$</button></td>
									<td><button onclick="addToEmail('%')">%</button></td>
									<td><button onclick="addToEmail('&')">&amp;</button></td>
									<td><button onclick="addToEmail('*')">*</button></td>
									<td><button onclick="addToEmail('+')">+</button></td>
									<td><button onclick="addToEmail('-')">-</button></td>
									<td><button onclick="addToEmail('/')">/</button></td>
									<td><button onclick="addToEmail('=')">=</button></td>
									<td><button onclick="addToEmail('?')">?</button></td>
									<td><button onclick="cancelEmail()">Back</button></td>
								</tr>
								<tr>
									<td><button onclick="addToEmail('{')">{</button></td>
									<td><button onclick="addToEmail('}')">}</button></td>
									<td><button onclick="addToEmail('|')">|</button></td>
									<td><button onclick="addToEmail('~')">~</button></td>
									<td><button onclick="addToEmail('.COM')">.COM</button></td>
									<td><button onclick="addToEmail('.NET')">.NET</button></td>
									<td><button onclick="addToEmail('.ORG')">.ORG</button></td>
									<td><button onclick="addToEmail('.EDU')">.EDU</button></td>
									<td><button onclick="addToEmail('.GOV')">.GOV</button></td>
									<td><button onclick="addToEmail('.INFO')">.INFO</button></td>
									<td><button onclick="getPhoneType()">OK</button></td>
								</tr>
							</table>
						</div>
					</div>
				</td>
			</tr>
			<tr height="105">
				<td valign="top" colspan="2" width="1010">
					<div align="center">
						<a href="/main.asp"><img src="/images/btn_mainmenu.jpg" alt="Main Menu" border="0" /></a><a href="/default.asp"><img src="/images/btn_signoff.jpg" alt="Sign Off" border="0" /></a><br />
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
