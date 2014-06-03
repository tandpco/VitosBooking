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



' Highlight the currently selected tab.
Dim currentTab
currentTab = "address"

%>
<!-- #Include Virtual="include2/globals.asp" -->
<!-- #Include Virtual="include2/math.asp" -->
<!-- #Include Virtual="include2/db-connect.asp" -->
<!-- #Include Virtual="include2/customer.asp" -->
<!-- #Include Virtual="include2/address.asp" -->
<%
Dim gnCustomerID, gnOrderAddressID, gnOrderID
Dim gsEMail, gsFirstName, gsLastName, gdtBirthdate, gnPrimaryAddressID, gsHomePhone, gsCellPhone, gsWorkPhone, gsFAXPhone, gbIsEMailList, gbIsTextList

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

If gnCustomerID = 0 Then
	gsFirstName = "NEW"
	gsLastName = "CUSTOMER"
Else
	If Not GetCustomerDetails(gnCustomerID, gsEMail, gsFirstName, gsLastName, gdtBirthdate, gnPrimaryAddressID, gsHomePhone, gsCellPhone, gsWorkPhone, gsFAXPhone, gbIsEMailList, gbIsTextList) Then
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
<script type="text/javascript" src="http://code.jquery.com/jquery-latest.js"></script>
<script type="text/javascript">
<!--
var ie4=document.all;
var gbHalfAddress = false;

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

function cancelAddress() {
	window.location = "<%=Session("ReturnURL")%>";
}

function toggleHalfAddress() {
	var loSpan;
	
	gbHalfAddress = !gbHalfAddress;
	loSpan = ie4? eval("document.all.HalfAddress") : document.getElementById('HalfAddress');
	if (gbHalfAddress) {
		loSpan.innerHTML = "&nbsp;1/2";
	}
	else {
		loSpan.innerHTML = "";
	}
	
	resetRedirect();
}

function goNext(psDigit) {
	var loText, lsLocation;
	
	lsLocation = "newaddress2.asp?<%=Request.ServerVariables("QUERY_STRING")%>&y=";
	loText = ie4? eval("document.all.streetnumber") : document.getElementById('streetnumber');
	if (loText.value.length > 0) {
		lsLocation = lsLocation + encodeURIComponent(loText.value) + "&x=" + psDigit;
		if (gbHalfAddress) {
			lsLocation = lsLocation + "&w=yes";
		}
		
		window.location = lsLocation;
	}
}

function back2Delivery() {
    var lsLocation = "/ordering/neworder.asp";
    //    alert("Back 2 Delivery");
    window.location = lsLocation;
}

function back2Phone() {
    var lsLocation = "/ordering/neworder.asp";
    //    alert("Back 2 Phone");
    window.location = lsLocation;
}

function back2Adx() {
    var lsLocation = "customerfind.asp";
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
			<!-- #Include Virtual="ordering/top-header.asp" -->
			<tr height="733">
				<td valign="top" width="1010">
					<div id="content-wrapper">
					<div id="content" style="position: relative; width: 1010px; height: 723px; overflow: auto;">
						<table align="center" cellpadding="0" cellspacing="0">
							<tr>
								<td valign="top" align="center" colspan="3"><strong>ADD 
								A NEW ADDRESS TO <%=gsFirstName & " " & gsLastName%></strong><br/><br/></td>
							</tr>
							<tr>
								<td valign="top">
									<table align="center" cellpadding="0" cellspacing="0">
										<tr>
											<td colspan="3"><div align="center">
												<strong>1) ENTER STREET #</strong></div></td>
										</tr>
										<tr>
											<td colspan="3" height="30"><div align="center">
												<input type="text" id="streetnumber" style="width: 150px" autocomplete="off" onkeydown="disableEnterKey();" /><span id="HalfAddress" name="HalfAddress" style="font-weight: bold;"></span></div></td>
										</tr>
										<tr>
											<td><button onclick="addToStreetNumber('1')">1</button></td>
											<td><button onclick="addToStreetNumber('2')">2</button></td>
											<td><button onclick="addToStreetNumber('3')">3</button></td>
										</tr>
										<tr>
											<td><button onclick="addToStreetNumber('4')">4</button></td>
											<td><button onclick="addToStreetNumber('5')">5</button></td>
											<td><button onclick="addToStreetNumber('6')">6</button></td>
										</tr>
										<tr>
											<td><button onclick="addToStreetNumber('7')">7</button></td>
											<td><button onclick="addToStreetNumber('8')">8</button></td>
											<td><button onclick="addToStreetNumber('9')">9</button></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td><button onclick="addToStreetNumber('0')">0</button></td>
											<td><button onclick="backspaceStreetNumber()">Bksp</button></td>
										</tr>
										<tr>
											<td colspan="3">&nbsp;</td>
										</tr>
										<tr>
											<td colspan="3" align="center"><button onclick="toggleHalfAddress()" style="width: 225px;">Toggle 1/2</button></td>
										</tr>
									</table>
								</td>
								<td valign="top" width="30">&nbsp;</td>
								<td valign="top">
									<table align="center" cellpadding="0" cellspacing="0">
										<tr>
											<td colspan="9"><div align="center">
												<strong>2) ENTER FIRST LETTER OF STREET NAME</strong></div></td>
										</tr>
										<tr>
											<td colspan="9" height="30"></td>
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
										</tr>
										<tr>
											<td><button onclick="goNext('A')">A</button></td>
											<td><button onclick="goNext('B')">B</button></td>
											<td><button onclick="goNext('C')">C</button></td>
											<td><button onclick="goNext('D')">D</button></td>
											<td><button onclick="goNext('E')">E</button></td>
											<td><button onclick="goNext('F')">F</button></td>
											<td><button onclick="goNext('G')">G</button></td>
											<td><button onclick="goNext('H')">H</button></td>
											<td><button onclick="goNext('I')">I</button></td>
										</tr>
										<tr>
											<td><button onclick="goNext('J')">J</button></td>
											<td><button onclick="goNext('K')">K</button></td>
											<td><button onclick="goNext('L')">L</button></td>
											<td><button onclick="goNext('M')">M</button></td>
											<td><button onclick="goNext('N')">N</button></td>
											<td><button onclick="goNext('O')">O</button></td>
											<td><button onclick="goNext('P')">P</button></td>
											<td><button onclick="goNext('Q')">Q</button></td>
											<td><button onclick="goNext('R')">R</button></td>
										</tr>
										<tr>
											<td><button onclick="goNext('S')">S</button></td>
											<td><button onclick="goNext('T')">T</button></td>
											<td><button onclick="goNext('U')">U</button></td>
											<td><button onclick="goNext('V')">V</button></td>
											<td><button onclick="goNext('W')">W</button></td>
											<td><button onclick="goNext('X')">X</button></td>
											<td><button onclick="goNext('Y')">Y</button></td>
											<td><button onclick="goNext('Z')">Z</button></td>
											<td><button onclick="cancelAddress()">Back</button></td>
										</tr>
									</table>
								</td>
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
