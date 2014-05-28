<%
Option Explicit
Response.buffer = TRUE

If Session("SecurityID") = "" Then
	Response.Redirect("/default.asp")
End If

Dim gsReturnURL

If Request("CustomerID").Count > 0 Then
	If Not IsNumeric(Request("CustomerID")) Then
		Response.Redirect("../main.asp")
	End If
Else
	Response.Redirect("../main.asp")
End If

If Request("AddressID").Count > 0 Then
	If Not IsNumeric(Request("AddressID")) Then
		Response.Redirect("../main.asp")
	End If
Else
	Response.Redirect("../main.asp")
End If

If Request("OrderID").Count > 0 Then
	If Not IsNumeric(Request("OrderID")) Then
		Response.Redirect("../main.asp")
	End If
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
Dim gsCustomerAddressDescription, gsCustomerAddressNotes

gnCustomerID = CLng(Request("CustomerID"))
gnAddressID = CLng(Request("AddressID"))
If Request("OrderID").Count > 0 Then
	gnOrderID = CLng(Request("OrderID"))
Else
	gnOrderID = 0
End If

If Request("ReturnURL").Count > 0 Then
	gsReturnURL = Request("ReturnURL")
Else
	If LCase(Right(Request.ServerVariables("HTTP_REFERER"), 14)) = "viewdrives.asp" Then
		gsReturnURL = "../viewdrives.asp"
	Else
		gsReturnURL = "editcustomer.asp?o=" & gnOrderID & "&c=" & gnCustomerID & "&a=" & gnAddressID
	End If
End If

If Request("action") = "savenotes" Then
	gsAddressNotes = Request("SaveAddressNotes")
	gsCustomerAddressNotes = Request("SaveCustomerNotes")
	
	If Not UpdateAddressNotes(gnAddressID, gsAddressNotes) Then
		Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
	End If
	
	If Not UpdateCustomerAddressNotes(gnCustomerID, gnAddressID, gsCustomerAddressNotes) Then
		Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
	End If
	
	Response.Redirect(gsReturnURL)
End If

If Not GetAddressDetails(gnAddressID, gnStoreID, gsAddress1, gsAddress2, gsCity, gsState, gsPostalCode, gsAddressNotes) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If Not GetCustomerAddressDetails(gnCustomerID, gnAddressID, gsCustomerAddressDescription, gsCustomerAddressNotes) Then
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
//-->
</script>
<script type="text/javascript">
<!--
var gbNotesInLowerCase = false;
var gsCurrentField = "addressnotes";

function setCurrentField(psCurrentField) {
	var loField;
	
	loField = ie4? eval("document.all." + gsCurrentField) : document.getElementById(gsCurrentField);
	loField.style.backgroundColor = "#CCCCCC";
	
	loField = ie4? eval("document.all." + psCurrentField) : document.getElementById(psCurrentField);
	loField.style.backgroundColor = "#FFFFFF";
	
	gsCurrentField = psCurrentField;
	
	resetRedirect();
}

function addToNotes(psDigit) {
	var loNotes, lsNotes;
	
	loNotes = ie4? eval("document.all." + gsCurrentField) : document.getElementById(gsCurrentField);
	lsNotes = loNotes.value;
	
	if (psDigit.length > 1) {
		if (lsNotes.length > 0) {
			lsNotes = lsNotes + " ";
		}
	}
	
	if (gbNotesInLowerCase) {
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

function backspaceNotes() {
	var loNotes, lsNotes;
	
	loNotes = ie4? eval("document.all." + gsCurrentField) : document.getElementById(gsCurrentField);
	lsNotes = loNotes.value;
	if (lsNotes.length > 0) {
		lsNotes = lsNotes.substr(0, (lsNotes.length - 1));
		loNotes.value = lsNotes;
	}
	
	resetRedirect();
}

function clearCurrentField() {
	var loNotes;
	
	loNotes = ie4? eval("document.all." + gsCurrentField) : document.getElementById(gsCurrentField);
	loNotes.value = "";
	
	resetRedirect();
}

function properNotes() {
	var loNotes, i, lsText, lbDoUpper;
	
	loNotes = ie4? eval("document.all." + gsCurrentField) : document.getElementById(gsCurrentField);
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

function shiftNotes() {
	var loObj;
	
	if (gbNotesInLowerCase) {
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
	
	gbNotesInLowerCase = !gbNotesInLowerCase;
	
	resetRedirect();
}

function saveNotes() {
	var loNotes, loFormNotes, loForm;
	
	loNotes = ie4? eval("document.all.addressnotes") : document.getElementById('addressnotes');
	loFormNotes = ie4? eval("document.all.SaveAddressNotes") : document.getElementById('SaveAddressNotes');
	loFormNotes.value = loNotes.value;
	
	loNotes = ie4? eval("document.all.customeraddressnotes") : document.getElementById('customeraddressnotes');
	loFormNotes = ie4? eval("document.all.SaveCustomerNotes") : document.getElementById('SaveCustomerNotes');
	loFormNotes.value = loNotes.value;
	
	loForm = ie4? eval("document.all.formNotes") : document.getElementById("formNotes");
	loForm.submit();
}

function back2Delivery() {
    var lsLocation = "neworder.asp";
    //    alert("Back 2 Delivery");
    window.location = lsLocation;
}

function back2Phone() {
    var lsLocation = "neworder.asp";
    //    alert("Back 2 Phone");
    window.location = lsLocation;
}

function back2Adx() {
    var lsLocation = "customerfind.asp";
    window.location = lsLocation;
}

function back2Order() {
    var lsLocation = "unitselect.asp";
    windows.location = lsLocation;
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
					<div id="content-wrapper" style="top:10px">
					<div id="content" style="position: relative; width: 1010px; height: 723px; overflow: auto;">
						<div id="orderlinenotes" style="position: absolute; left: 0px; top: 0px; width: 1010px; height: 723px; background-color: #fbf3c5;">
							<form id="formNotes" name="formNotes" method="post" action="addressnotes.asp">
								<input type="hidden" id="action" name="action" value="savenotes" />
								<input type="hidden" id="ReturnURL" name="ReturnURL" value="<%=gsReturnURL%>" />
								<input type="hidden" id="CustomerID" name="CustomerID" value="<%=gnCustomerID%>" />
								<input type="hidden" id="AddressID" name="AddressID" value="<%=gnAddressID%>" />
								<input type="hidden" id="SaveAddressNotes" name="SaveAddressNotes" value="<%=gsAddressNotes%>" />
								<input type="hidden" id="SaveCustomerNotes" name="SaveCustomerNotes" value="<%=gsCustomerAddressNotes%>" />
								<p align="center"><strong>Editing Notes For <%=gsAddress1%>&nbsp;<%=gsAddress2%>, <%=gsCity%>, <%=gsState%>&nbsp;<%=gsPostalCode%></strong></p>
							</form>
							<table width="925" align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="3">
										<table align="center" cellpadding="0" cellspacing="0">
											<tr>
												<td><strong>General Address Notes:</strong><br/>
												<textarea id="addressnotes" name="addressnotes" style="width: 450px; height: 60px;" onfocus="setCurrentField('addressnotes');"><%=Server.HTMLEncode(gsAddressNotes)%></textarea></td>
												<td>&nbsp;</td>
												<td><strong>Customer Specific Address Notes:</strong><br/>
												<textarea id="customeraddressnotes" name="customeraddressnotes" style="width: 450px; height: 60px; background-color: #cccccc;" onfocus="setCurrentField('customeraddressnotes');"><%=Server.HTMLEncode(gsCustomerAddressNotes)%></textarea></td>
											</tr>
										</table>
									</td>
								</tr>
							</table>
							<table width="925" align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td><button style="width: 100px;" onclick="window.location = '<%=gsReturnURL%>'">Cancel</button></td>
									<td align="center"><button style="width: 100px;" onclick="clearCurrentField()">Clear</button></td>
									<td align="right"><button style="width: 100px;" onclick="saveNotes()">Done</button></td>
								</tr>
							</table>
							<table align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td><div align="center">
										<button onclick="addToNotes('+')">+</button><button onclick="addToNotes('!')">!</button><button onclick="addToNotes('@')">@</button><button onclick="addToNotes('#')">#</button><button onclick="addToNotes('$')">$</button><button onclick="addToNotes('%')">%</button><button onclick="addToNotes('^')">^</button><button onclick="addToNotes('&')">&amp;</button><button onclick="addToNotes('*')">*</button><button onclick="addToNotes('(')">(</button><button onclick="addToNotes(')')">)</button><button onclick="addToNotes(':')">:</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="addToNotes('=')">=</button><button onclick="addToNotes('1')">1</button><button onclick="addToNotes('2')">2</button><button onclick="addToNotes('3')">3</button><button onclick="addToNotes('4')">4</button><button onclick="addToNotes('5')">5</button><button onclick="addToNotes('6')">6</button><button onclick="addToNotes('7')">7</button><button onclick="addToNotes('8')">8</button><button onclick="addToNotes('9')">9</button><button onclick="addToNotes('0')">0</button><button onclick="addToNotes('?')">?</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="addToNotes('\'')">'</button><button name="key-q" id="key-q" onclick="addToNotes('Q')">Q</button><button name="key-w" id="key-w" onclick="addToNotes('W')">W</button><button name="key-e" id="key-e" onclick="addToNotes('E')">E</button><button name="key-r" id="key-r" onclick="addToNotes('R')">R</button><button name="key-t" id="key-t" onclick="addToNotes('T')">T</button><button name="key-y" id="key-y" onclick="addToNotes('Y')">Y</button><button name="key-u" id="key-u" onclick="addToNotes('U')">U</button><button name="key-i" id="key-i" onclick="addToNotes('I')">I</button><button name="key-o" id="key-o" onclick="addToNotes('O')">O</button><button name="key-p" id="key-p" onclick="addToNotes('P')">P</button><button onclick="addToNotes('/')">/</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="addToNotes('.')">.</button><button name="key-a" id="key-a" onclick="addToNotes('A')">A</button><button name="key-s" id="key-s" onclick="addToNotes('S')">S</button><button name="key-d" id="key-d" onclick="addToNotes('D')">D</button><button name="key-f" id="key-f" onclick="addToNotes('F')">F</button><button name="key-g" id="key-g" onclick="addToNotes('G')">G</button><button name="key-h" id="key-h" onclick="addToNotes('H')">H</button><button name="key-j" id="key-j" onclick="addToNotes('J')">J</button><button name="key-k" id="key-k" onclick="addToNotes('K')">K</button><button name="key-l" id="key-l" onclick="addToNotes('L')">L</button><button onclick="addToNotes(',')">,</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="shiftNotes()">Shift</button><button onclick="addToNotes('<')">&lt;</button><button name="key-z" id="key-z" onclick="addToNotes('Z')">Z</button><button name="key-x" id="key-x" onclick="addToNotes('X')">X</button><button name="key-c" id="key-c" onclick="addToNotes('C')">C</button><button name="key-v" id="key-v" onclick="addToNotes('V')">V</button><button name="key-b" id="key-b" onclick="addToNotes('B')">B</button><button name="key-n" id="key-n" onclick="addToNotes('N')">N</button><button name="key-m" id="key-m" onclick="addToNotes('M')">M</button><button onclick="addToNotes('>')">&gt;</button><button onclick="properNotes()">Proper</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button style="width: 150px;" onclick="addToNotes('1/2')">1/2</button>&nbsp;<button style="width: 600px;" onclick="addToNotes(' ')">Space</button>&nbsp;<button style="width: 150px;" onclick="backspaceNotes()">Bksp</button></div></td>
								</tr>
							</table>
						</div>
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
