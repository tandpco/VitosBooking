<%
Option Explicit
Response.buffer = TRUE

If Session("SecurityID") = "" Then
	Response.Redirect("/default.asp")
End If

If Request("o").Count > 0 Then
	If Not IsNumeric(Request("o")) Then
		Response.Redirect("../main.asp")
	End If
'Else
'	Response.Redirect("../main.asp")
End If

If Request("c").Count > 0 Then
	If Not IsNumeric(Request("c")) Then
		Response.Redirect("../ticket.asp?OrderID=" & Request("o"))
	End If
'Else
'	Response.Redirect("../ticket.asp?OrderID=" & Request("o"))
End If

If Request("a").Count > 0 Then
	If Not IsNumeric(Request("a")) Then
		Response.Redirect("../ticket.asp?OrderID=" & Request("o"))
	End If
'Else
'	Response.Redirect("../ticket.asp?OrderID=" & Request("o"))
End If

If Request("y").Count > 0 Then
	If Not IsNumeric(Request("y")) Then
		Response.Redirect("../ticket.asp?OrderID=" & Request("o"))
	End If
Else
	Response.Redirect("../ticket.asp?OrderID=" & Request("o"))
End If

If Request("x").Count = 0 Then
	Response.Redirect("../ticket.asp?OrderID=" & Request("o"))
End If

If Request("z").Count > 0 Then
	If Not IsNumeric(Request("y")) Then
		Response.Redirect("../ticket.asp?OrderID=" & Request("o"))
	End If
End If
%>
<!-- #Include Virtual="include2/globals.asp" -->
<!-- #Include Virtual="include2/math.asp" -->
<!-- #Include Virtual="include2/db-connect.asp" -->
<!-- #Include Virtual="include2/customer.asp" -->
<!-- #Include Virtual="include2/address.asp" -->
<!-- #Include Virtual="include2/store.asp" -->
<%
Dim gnCustomerID, gnOrderAddressID, gnOrderID, gnStreetNumber, gsStreetLetter, gbHalfAddress
Dim gsEMail, gsFirstName, gsLastName, gdtBirthdate, gnPrimaryAddressID, gsHomePhone, gsCellPhone, gsWorkPhone, gsFAXPhone, gbIsEMailList, gbIsTextList
Dim gsPostalCode, gasSearchPostalCodes(), gasStreets(), gasPostalCodes(), i

If Request("c").Count > 0 Then
	gnCustomerID = CLng(Request("c"))
Else
	gnCustomerID = 0
End If
If Request("a").Count > 0 Then
	gnOrderAddressID = CLng(Request("a"))
Else
	gnOrderAddressID = 0
End If
If Request("o").Count > 0 Then
	gnOrderID = CLng(Request("o"))
Else
	gnOrderID = 0
End If
gnStreetNumber = CLng(Request("y"))
gsStreetLetter = Request("x")
If Request("w") = "yes" Then
	gbHalfAddress = TRUE
Else
	gbHalfAddress = FALSE
End If

If gnCustomerID = 0 Then
	gsFirstName = "NEW"
	gsLastName = "CUSTOMER"
Else
	If Not GetCustomerDetails(gnCustomerID, gsEMail, gsFirstName, gsLastName, gdtBirthdate, gnPrimaryAddressID, gsHomePhone, gsCellPhone, gsWorkPhone, gsFAXPhone, gbIsEMailList, gbIsTextList) Then
		Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
	End If
End If

If Request("z").Count > 0 Then
	gsPostalCode = Request("z")
	
	If Not GetStreetList(gsPostalCode, gnStreetNumber, gsStreetLetter, gasStreets) Then
		Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
	End If
	
	ReDim gasPostalCodes(UBound(gasStreets))
	For i = 0 To UBound(gasStreets)
		gasPostalCodes(i) = gsPostalCode
	Next
Else
	gsPostalCode = ""
	
	If Not GetStorePostalCodes(Session("StoreID"), gasSearchPostalCodes) Then
		Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
	End If
	
	If Not GetStreetListByZipCodes(gasSearchPostalCodes, gnStreetNumber, gsStreetLetter, gasStreets, gasPostalCodes) Then
		Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
	End If
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
var gsStreet = "";
var gsPostalCode = "";
var gbManual = false;

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

function cancelStreet() {
	window.location = "<%=Session("ReturnURL")%>";
}

function selectStreet(psStreet, psPostalCode) {
	var loText, loDiv;
	
	gsStreet = psStreet;
	gsPostalCode = psPostalCode;
	
	loText = ie4? eval("document.all.apt") : document.getElementById('apt');
	loText.value = "";
	
	loDiv = ie4? eval("document.all.streetdiv") : document.getElementById('streetdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.postalcodediv") : document.getElementById('postalcodediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.manualdiv") : document.getElementById('manualdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.aptdiv") : document.getElementById('aptdiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function gotoZipCode() {
	var loDiv;
	
	loDiv = ie4? eval("document.all.streetdiv") : document.getElementById('streetdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.aptdiv") : document.getElementById('aptdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.manualdiv") : document.getElementById('manualdiv');
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

function selectPostalCode() {
	var loText, lsLocation;
	
	lsLocation = "newaddress2.asp?<%=Request.ServerVariables("QUERY_STRING")%>&z=";
	loText = ie4? eval("document.all.postalcode") : document.getElementById('postalcode');
	if (loText.value.length > 0) {
		lsLocation = lsLocation + encodeURIComponent(loText.value)
		
		window.location = lsLocation;
	}
}

function addToApt(psDigit) {
	var loName, lsName;
	
	loName = ie4? eval("document.all.apt") : document.getElementById('apt');
	lsName = loName.value;
	lsName += psDigit;
	loName.value = lsName;
	
	resetRedirect();
}

function backspaceApt() {
	var loName, lsName;
	
	loName = ie4? eval("document.all.apt") : document.getElementById('apt');
	lsName = loName.value;
	if (lsName.length > 0) {
		lsName = lsName.substr(0, (lsName.length - 1));
		loName.value = lsName;
	}
	
	resetRedirect();
}

function cancelApt() {
	var loText, loDiv;
	
	gbManual = false;
	
	loDiv = ie4? eval("document.all.aptdiv") : document.getElementById('aptdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.postalcodediv") : document.getElementById('postalcodediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.manualdiv") : document.getElementById('manualdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.streetdiv") : document.getElementById('streetdiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function saveAddress() {
	var lsLocation, loText;
	
	lsLocation = "<%=Session("SaveURL")%>&Action=SaveAddress&z=" + gsPostalCode + "&b=";
<%
If gbHalfAddress Then
%>
	lsLocation = lsLocation + encodeURIComponent("<%=gnStreetNumber%> 1/2 " + gsStreet);
<%
Else
%>
	lsLocation = lsLocation + encodeURIComponent("<%=gnStreetNumber%> " + gsStreet);
<%
End If
%>
	loText = ie4? eval("document.all.apt") : document.getElementById('apt');
	if (loText.value.length > 0) {
		lsNumbers = "";
		lsLetters = "";
		
		for (i = 0; i < loText.value.length; i++) {
			if (loText.value.substr(i, 1) != " ") {
				if (isNaN(Number(loText.value.substr(i, 1)))) {
					lsLetters = lsLetters + loText.value.substr(i, 1);
				}
				else {
					lsNumbers = lsNumbers + loText.value.substr(i, 1);
				}
			}
		}
		
		lsLocation = lsLocation + encodeURIComponent(" #" + lsNumbers + lsLetters);
	}
	if (gbManual) {
		lsLocation = lsLocation + "&Manual=Yes";
	}
<%
If Request("c").Count > 0 And InStr(Session("SaveURL"), "&c=") = 0 Then
%>
	lsLocation = lsLocation + "&c=<%=Request("c")%>";
<%
End If
%>
	
	window.location = lsLocation;
}

function gotoManual(psPostalCode) {
	var loText, loDiv;
	
	gbManual = true;
	gsPostalCode = psPostalCode;
	
	loText = ie4? eval("document.all.manual") : document.getElementById('manual');
	loText.value = "";
	
	loDiv = ie4? eval("document.all.streetdiv") : document.getElementById('streetdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.aptdiv") : document.getElementById('aptdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.postalcodediv") : document.getElementById('postalcodediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.manualdiv") : document.getElementById('manualdiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function addToManual(psDigit) {
	var loName, lsName;
	
	loName = ie4? eval("document.all.manual") : document.getElementById('manual');
	lsName = loName.value;
	lsName += psDigit;
	loName.value = lsName;
	
	resetRedirect();
}

function backspaceManual() {
	var loName, lsName;
	
	loName = ie4? eval("document.all.manual") : document.getElementById('manual');
	lsName = loName.value;
	if (lsName.length > 0) {
		lsName = lsName.substr(0, (lsName.length - 1));
		loName.value = lsName;
	}
	
	resetRedirect();
}

function cancelManual() {
	var loText, loDiv;
	
	gbManual = false;
	
	loDiv = ie4? eval("document.all.aptdiv") : document.getElementById('aptdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.postalcodediv") : document.getElementById('postalcodediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.manualdiv") : document.getElementById('manualdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.streetdiv") : document.getElementById('streetdiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function saveManual() {
	var loText, loDiv;
	
	loText = ie4? eval("document.all.manual") : document.getElementById('manual');
	gsStreet = loText.value;
	
	loText = ie4? eval("document.all.apt") : document.getElementById('apt');
	loText.value = "";
	
	loDiv = ie4? eval("document.all.streetdiv") : document.getElementById('streetdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.postalcodediv") : document.getElementById('postalcodediv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.manualdiv") : document.getElementById('manualdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.aptdiv") : document.getElementById('aptdiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
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
						<span id="ClockDate"><%=clockDateString(gDate)%></span> 
						|
						<span id="ClockTime" onclick="clockToggleSeconds()"><%=clockTimeString(Hour(gDate), Minute(gDate), Second(gDate))%></span>
					</div>
				</td>
			</tr>
			<tr height="733">
				<td valign="top" width="1010">
					<div id="content" style="position: relative; width: 1010px; height: 723px; overflow: auto;">
						<div id="streetdiv" style="position: absolute; top: 0px; left: 0px; width: 1010px;">
							<div align="center"><strong>SELECT STREET TO ADD 
								A NEW ADDRESS TO <%=gsFirstName & " " & gsLastName%></strong></div>
							<div style="position: relative; width: 1000px;">
<%
Dim lnTop, lnLeft

lnTop = 0
lnLeft = 0
For i = 0 To UBound(gasStreets)
	If i Mod 42 = 0 Then
		If i > 0 And UBound(gasStreets) > 42 Then
			If Len(gsPostalCode) > 0 Then
%>
									<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 195px; height: 75px;" onclick="gotoManual()">Still Not Found?</button></div>
<%
			Else
%>
									<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 195px; height: 75px;" onclick="gotoZipCode()">Can't Find Street?</button></div>
<%
			End If
%>
									<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft+196%>px;"><button style="width: 195px; height: 75px;" onclick="cancelStreet()">Back</button></div>
									<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft+392%>px;"><button style="width: 195px; height: 75px;" onclick="toggleDivs('streetdiv<%=Int(i/33)-1%>', 'streetdiv<%=Int(i/33)%>')">(Next)</button></div>
								</div>
								<div id="streetdiv<%=Int(i/33)%>" style="position: absolute; top: 0px; left: 0px; width: 1000px; visibility: hidden;">
<%
		Else
			If i = 0 Then
%>
								<div id="streetdiv<%=Int(i/33)%>" style="position: absolute; top: 0px; left: 0px; width: 1000px;">
<%
			End If
		End If
		
		lnTop = 0
		lnLeft = 0
	End If
	
	If Len(gasStreets(i)) = 0 Then
%>
									<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 195px; height: 75px; background-color: #C0C0C0;">&nbsp;</button></div>
<%
	Else
%>
									<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 195px; height: 75px;" onclick="selectStreet('<%=gasStreets(i)%>', '<%=gasPostalCodes(i)%>')"><%=gasStreets(i)%></button></div>
<%
	End If
	
	If lnLeft = 784 Then
		lnTop = lnTop + 76
		lnLeft = 0
	Else
		lnLeft = lnLeft + 196
	End If
Next

' Add hidden buttons here
If ((UBound(gasStreets) + 1) Mod 42) > 0 And UBound(gasStreets) <> 42 Then
	For i = ((UBound(gasStreets) + 1) Mod 42) To 41
%>
									<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 195px; height: 75px; background-color: #C0C0C0;">&nbsp;</button></div>
<%
		If lnLeft = 784 Then
			lnTop = lnTop + 76
			lnLeft = 0
		Else
			lnLeft = lnLeft + 196
		End If
	Next
End If

If UBound(gasStreets) > 42 Then
			If Len(gsPostalCode) > 0 Then
%>
									<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 195px; height: 75px;" onclick="gotoManual('<%=gsPostalCode%>')">Still Not Found?</button></div>
<%
			Else
%>
									<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 195px; height: 75px;" onclick="gotoZipCode()">Can't Find Street?</button></div>
<%
			End If
%>
									<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft+196%>px;"><button style="width: 195px; height: 75px;" onclick="cancelStreet()">Back</button></div>
									<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft+392%>px;"><button style="width: 195px; height: 75px;" onclick="toggleDivs('streetdiv<%=Int(UBound(gasStreets)/33)%>', 'streetdiv0')">(Next)</button></div>
<%
Else
%>
									<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 195px; height: 75px; background-color: #C0C0C0;">&nbsp;</button></div>
<%
			If Len(gsPostalCode) > 0 Then
%>
									<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft+196%>px;"><button style="width: 195px; height: 75px;" onclick="gotoManual('<%=gsPostalCode%>')">Still Not Found?</button></div>
<%
			Else
%>
									<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft+196%>px;"><button style="width: 195px; height: 75px;" onclick="gotoZipCode()">Can't Find Street?</button></div>
<%
			End If
%>
									<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft+392%>px;"><button style="width: 195px; height: 75px;" onclick="cancelStreet()">Back</button></div>
<%
End If
%>
								</div>
							</div>
						</div>
						<div id="postalcodediv" style="position: absolute; top: 0px; left: 0px; width: 1010px; visibility: hidden;">
							<table align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="3"><div align="center">
										<strong>ENTER POSTAL CODE</strong></div></td>
								</tr>
								<tr>
									<td colspan="3"><div align="center">
										<input type="text" id="postalcode" style="width: 200px" autocomplete="off" onkeydown="disableEnterKey();" value="<%=Request("z")%>" /></div></td>
								</tr>
								<tr>
									<td><button onclick="addToPostalCode('1')">1</button></td>
									<td><button onclick="addToPostalCode('2')">2</button></td>
									<td><button onclick="addToPostalCode('3')">3</button></td>
								</tr>
								<tr>
									<td><button onclick="addToPostalCode('4')">4</button></td>
									<td><button onclick="addToPostalCode('5')">5</button></td>
									<td><button onclick="addToPostalCode('6')">6</button></td>
								</tr>
								<tr>
									<td><button onclick="addToPostalCode('7')">7</button></td>
									<td><button onclick="addToPostalCode('8')">8</button></td>
									<td><button onclick="addToPostalCode('9')">9</button></td>
								</tr>
								<tr>
									<td>&nbsp;</td>
									<td><button onclick="addToPostalCode('0')">0</button></td>
									<td><button onclick="backspacePostalCode()">Bksp</button></td>
								</tr>
								<tr>
									<td colspan="3">&nbsp;</td>
								</tr>
								<tr>
									<td colspan="3" align="center"><button onclick="cancelStreet();">Cancel</button>&nbsp;<button onclick="selectPostalCode();">Search</button></td>
								</tr>
							</table>
						</div>
						<div id="aptdiv" style="position: absolute; top: 0px; left: 0px; width: 1010px; visibility: hidden;">
							<table align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="11"><div align="center">
										<strong>ENTER APT/SUITE/UNIT</strong></div></td>
								</tr>
								<tr>
									<td colspan="11"><div align="center">
										<input type="text" id="apt" style="width: 800px" autocomplete="off" onkeydown="disableEnterKey();" /></div></td>
								</tr>
								<tr>
									<td><button onclick="addToApt('1')">1</button></td>
									<td><button onclick="addToApt('2')">2</button></td>
									<td><button onclick="addToApt('3')">3</button></td>
									<td><button onclick="addToApt('4')">4</button></td>
									<td><button onclick="addToApt('5')">5</button></td>
									<td><button onclick="addToApt('6')">6</button></td>
									<td><button onclick="addToApt('7')">7</button></td>
									<td><button onclick="addToApt('8')">8</button></td>
									<td><button onclick="addToApt('9')">9</button></td>
									<td><button onclick="addToApt('0')">0</button></td>
									<td><button onclick="backspaceApt()">Bksp</button></td>
								</tr>
								<tr>
									<td><button onclick="addToApt('Q')">Q</button></td>
									<td><button onclick="addToApt('W')">W</button></td>
									<td><button onclick="addToApt('E')">E</button></td>
									<td><button onclick="addToApt('R')">R</button></td>
									<td><button onclick="addToApt('T')">T</button></td>
									<td><button onclick="addToApt('Y')">Y</button></td>
									<td><button onclick="addToApt('U')">U</button></td>
									<td><button onclick="addToApt('I')">I</button></td>
									<td><button onclick="addToApt('O')">O</button></td>
									<td><button onclick="addToApt('P')">P</button></td>
									<td><button onclick="addToApt('#')">#</button></td>
								</tr>
								<tr>
									<td><button onclick="addToApt('A')">A</button></td>
									<td><button onclick="addToApt('S')">S</button></td>
									<td><button onclick="addToApt('D')">D</button></td>
									<td><button onclick="addToApt('F')">F</button></td>
									<td><button onclick="addToApt('G')">G</button></td>
									<td><button onclick="addToApt('H')">H</button></td>
									<td><button onclick="addToApt('J')">J</button></td>
									<td><button onclick="addToApt('K')">K</button></td>
									<td><button onclick="addToApt('L')">L</button></td>
									<td><button onclick="addToApt(''')">'</button></td>
									<td><button onclick="cancelApt()">Cancel</button></td>
								</tr>
								<tr>
									<td><button onclick="addToApt('Z')">Z</button></td>
									<td><button onclick="addToApt('X')">X</button></td>
									<td><button onclick="addToApt('C')">C</button></td>
									<td><button onclick="addToApt('V')">V</button></td>
									<td><button onclick="addToApt('B')">B</button></td>
									<td><button onclick="addToApt('N')">N</button></td>
									<td><button onclick="addToApt('M')">M</button></td>
									<td><button onclick="addToApt('.')">.</button></td>
									<td><button onclick="addToApt(',')">,</button></td>
									<td><button onclick="addToApt(' ')">Space</button></td>
									<td><button onclick="saveAddress()">OK</button></td>
								</tr>
							</table>
						</div>
						<div id="manualdiv" style="position: absolute; top: 0px; left: 0px; width: 1010px; visibility: hidden;">
							<table align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="11"><div align="center">
										<strong>ENTER STREET NAME</strong></div></td>
								</tr>
								<tr>
									<td colspan="11"><div align="center">
										<input type="text" id="manual" style="width: 800px" autocomplete="off" onkeydown="disableEnterKey();" /></div></td>
								</tr>
								<tr>
									<td><button onclick="addToManual('1')">1</button></td>
									<td><button onclick="addToManual('2')">2</button></td>
									<td><button onclick="addToManual('3')">3</button></td>
									<td><button onclick="addToManual('4')">4</button></td>
									<td><button onclick="addToManual('5')">5</button></td>
									<td><button onclick="addToManual('6')">6</button></td>
									<td><button onclick="addToManual('7')">7</button></td>
									<td><button onclick="addToManual('8')">8</button></td>
									<td><button onclick="addToManual('9')">9</button></td>
									<td><button onclick="addToManual('0')">0</button></td>
									<td><button onclick="backspaceManual()">Bksp</button></td>
								</tr>
								<tr>
									<td><button onclick="addToManual('Q')">Q</button></td>
									<td><button onclick="addToManual('W')">W</button></td>
									<td><button onclick="addToManual('E')">E</button></td>
									<td><button onclick="addToManual('R')">R</button></td>
									<td><button onclick="addToManual('T')">T</button></td>
									<td><button onclick="addToManual('Y')">Y</button></td>
									<td><button onclick="addToManual('U')">U</button></td>
									<td><button onclick="addToManual('I')">I</button></td>
									<td><button onclick="addToManual('O')">O</button></td>
									<td><button onclick="addToManual('P')">P</button></td>
									<td><button onclick="addToManual('#')">#</button></td>
								</tr>
								<tr>
									<td><button onclick="addToManual('A')">A</button></td>
									<td><button onclick="addToManual('S')">S</button></td>
									<td><button onclick="addToManual('D')">D</button></td>
									<td><button onclick="addToManual('F')">F</button></td>
									<td><button onclick="addToManual('G')">G</button></td>
									<td><button onclick="addToManual('H')">H</button></td>
									<td><button onclick="addToManual('J')">J</button></td>
									<td><button onclick="addToManual('K')">K</button></td>
									<td><button onclick="addToManual('L')">L</button></td>
									<td><button onclick="addToManual(''')">'</button></td>
									<td><button onclick="cancelManual()">Cancel</button></td>
								</tr>
								<tr>
									<td><button onclick="addToManual('Z')">Z</button></td>
									<td><button onclick="addToManual('X')">X</button></td>
									<td><button onclick="addToManual('C')">C</button></td>
									<td><button onclick="addToManual('V')">V</button></td>
									<td><button onclick="addToManual('B')">B</button></td>
									<td><button onclick="addToManual('N')">N</button></td>
									<td><button onclick="addToManual('M')">M</button></td>
									<td><button onclick="addToManual('.')">.</button></td>
									<td><button onclick="addToManual(',')">,</button></td>
									<td><button onclick="addToManual(' ')">Space</button></td>
									<td><button onclick="saveManual()">OK</button></td>
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

</body>

</html>
<!-- #Include Virtual="include2/db-disconnect.asp" -->
