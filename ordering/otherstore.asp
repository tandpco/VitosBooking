<%
Option Explicit
Response.buffer = TRUE

If Session("SecurityID") = "" Then
	Response.Redirect("/default.asp")
End If

If Request("s").Count = 0 Then
	Response.Redirect("neworder.asp")
End If

If Not IsNumeric(Request("s")) Then
	Response.Redirect("neworder.asp")
End If

If Request("c").Count <> 0 Then
	If Not IsNumeric(Request("c")) Then
		Response.Redirect("neworder.asp")
	End If
End If

If Request("a").Count = 0 Then
	Response.Redirect("neworder.asp")
End If

If Not IsNumeric(Request("a")) Then
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
Dim gnStoreID, gsStoreName, gsStoreAddress1, gsStoreAddress2, gsStoreCity, gsStoreState, gsStorePostalCode, gsStorePhone, gsStoreFAX, gsStoreHours
Dim ganOrderIDs(), gsLocalErrorMsg, i, gnCustomerID, gsCustomerName, gnAddressID
Dim lsFirst, lsLast, lsHome, lsCell, lsWork, lsFAX
Dim gnStoreID1, gsAddress1, gsAddress2, gsCity, gsState, gsPostalCode, gsAddressNotes
Dim gsDisplayAddress
Dim ganEmpIDs(), ganEmployeeIDs(), gasCardIDs()
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

If Request("c").Count <> 0 Then
	gnCustomerID = CLng(Request("c"))
Else
	gnCustomerID = 0
End If

gnAddressID = CLng(Request("a"))
gnStoreID = Request("s")
If Not GetStoreDetails(gnStoreID, gsStoreName, gsStoreAddress1, gsStoreAddress2, gsStoreCity, gsStoreState, gsStorePostalCode, gsStorePhone, gsStoreFAX, gsStoreHours) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If Request("newcustomer") = "yes" Then
	If Request("n").Count = 0 Then
		Response.Redirect("neworder.asp")
	Else
		If Request("h").Count = 0 Then
			Response.Redirect("neworder.asp")
		Else
			If Not IsNumeric(Request("h")) Then
				Response.Redirect("neworder.asp")
			Else
				If Request("a").Count = 0 Then
					Response.Redirect("neworder.asp")
				Else
					If Not IsNumeric(Request("a")) Then
						Response.Redirect("neworder.asp")
					Else
						gsCustomerName = Request("n")
						
						lsFirst = ""
						lsLast = ""
						If InStr(gsCustomerName, " ") = 0 Then
							lsFirst = gsCustomerName
							lsLast = ""
						Else
							lsFirst = Left(gsCustomerName, (InStr(gsCustomerName, " ") - 1))
							lsLast = Mid(gsCustomerName, (InStr(gsCustomerName, " ") + 1))
						End If
						
						lsHome = ""
						lsCell = ""
						lsWork = ""
						lsFAX = ""
						
						Select Case CLng(Request("h"))
							Case 1
								lsCell = Session("CustomerPhone")
							Case 2
								lsWork = Session("CustomerPhone")
							Case 3
								lsFAX = Session("CustomerPhone")
							Case Else
								lsHome = Session("CustomerPhone")
						End Select
						
						gnCustomerID = AddCustomer(LCase(Request("em")), "", lsFirst, lsLast, DateValue("1/1/1900"), gnAddressID, lsHome, lsCell, lsWork, lsFAX, FALSE, FALSE)
						If gnCustomerID = 0 Then
							Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
						Else
							If Not AddCustomerAddress(gnCustomerID, gnAddressID, "Primary Address") Then
								Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
							End If
						End If
					End If
				End If
			End If
		End If
	End If
Else
	If Request("assigncustomerphone") = "yes" Then
		If Request("h").Count = 0 Then
			Response.Redirect("neworder.asp")
		Else
			If Not IsNumeric(Request("h")) Then
				Response.Redirect("neworder.asp")
			Else
				If Not AssignCustomerPhone(gnCustomerID, Session("CustomerPhone"), CLng(Request("h"))) Then
					Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
				End If
			End If
		End If
	Else
		If Request("assigncustomeraddress") = "yes" Then
			If Not AddCustomerAddress(gnCustomerID, Request("a"), "Alternate Address") Then
				Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
			End If
		End If
	End If
End If

If gnAddressID = 0 Then
	gsDisplayAddress = Request("b") & " in zip code " & Request("z")
Else
	If Not GetAddressDetails(gnAddressID, gnStoreID1, gsAddress1, gsAddress2, gsCity, gsState, gsPostalCode, gsAddressNotes) Then
		Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
	End If
	
	If Len(gsAddress2) = 0 Then
		If Len(gsAddress1) > 0 Then
			gsDisplayAddress = gsAddress1
		End If
	Else
		gsDisplayAddress = gsAddress1 & " #" & gsAddress2
	End If
	
	If Len(gsCity) > 0 Or Len(gsState) > 0 Or Len(gsPostalCode) > 0 Then
		gsDisplayAddress = gsDisplayAddress & ", " & gsCity & ", " & gsState & " " & gsPostalCode
	End If
End If

If Not GetAllStoreManagers(Session("StoreID"), ganEmpIDs, ganEmployeeIDs, gasCardIDs) Then
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

function gotoMgrOverride()
{
	var loManager, loDiv;
	
	loManager = ie4? eval("document.all.manager") : document.getElementById('manager');
	loManager.value = "";
	
	loDiv = ie4? eval("document.all.otherstorediv") : document.getElementById('otherstorediv');
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
	
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById('managerdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.otherstorediv") : document.getElementById('otherstorediv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function checkManagerEnterKey() {
	if (event.keyCode == 13) {
		setManager();
	}
}

function setManager() {
	var loManager, i, lbFound, lsLocation;
	
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
			lsLocation = "unitselect.asp?t=1&c=<%=gnCustomerID%>&a=<%=gnAddressID%>";
			window.location = lsLocation;
		}
		else {
			loManager.value = "";
			loManager.focus();
		}
	}
	
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
						<span id="ClockDate"><%=clockDateString(gDate)%></span> |
						<span id="ClockTime" onclick="clockToggleSeconds()"><%=clockTimeString(Hour(gDate), Minute(gDate), Second(gDate))%></span>
					</div>
				</td>
			</tr>
			<tr height="628">
				<td valign="top" width="1010">
					<div id="content" style="position: relative; width: 1020px; height: 615px; overflow: auto;">
						<div id="otherstorediv" style="position: absolute; top: 0px; left: 0px; width: 1010px; height: 605px;">
							<p align="center"><strong><%=gsDisplayAddress%></strong></p>
							<p align="center"><strong>Apparently the address you entered is not in our delivery area. 
							Confirm that the address above is correct.</strong></p>
<%
If gnAddressID = 0 Then
%>
							<p align="center"><strong>There are no stores that deliver to that address.</strong></p>
							<p align="center"><button onclick="window.location='customerfind.asp?t=1&p=<%=Session("CustomerPhone")%>'">Re-Enter Address</button>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								<button onclick="window.location='unitselect.asp?t=2&c=<%=gnCustomerID%>&a=<%=gnAddressID%>'">Switch To Pickup</button></p>
<%
Else
	If gnStoreID = 0 Then
%>
							<p align="center"><strong>There are no stores that deliver to that address.</strong></p>
<%
	Else
%>
							<p align="center"><strong>If the address is correct, 
							please refer the customer to the following store:<br/>&nbsp;<br/>
								<%=gsStoreAddress1%><br/>
<%
		If Len(gsStoreAddress2) > 0 Then
%>
								<%=gsStoreAddress2%><br/>
<%
		End If
%>
								<%=gsStoreCity%>, <%=gsStoreState%>&nbsp;<%=gsStorePostalCode%><br/>
								Phone: <%=gsStorePhone%><br/>
<%
		If Len(gsStoreFAX) > 0 Then
%>
								FAX: <%=gsStoreFAX%><br/>
<%
		End If
%>
								<%=gsStoreHours%></strong></p>
<%
	End If
%>
							<p align="center"><strong>If you feel this address is in your delivery area or the order is going to be for more than $30, use Manager Override.</strong></p>
							<p align="center"><button onclick="window.location='customerfind.asp?t=1&p=<%=Session("CustomerPhone")%>'">Re-Enter Address</button>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								<button onclick="window.location='unitselect.asp?t=2&c=<%=gnCustomerID%>&a=<%=gnAddressID%>'">Switch To Pickup</button>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<%
	If Session("SecurityID") > 1 And ((Left(Request.ServerVariables("remote_addr"), 9) = "10.0.254." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.1." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.2.") OR Session("Swipe")) Then
%>
								<button onclick="window.location = 'unitselect.asp?t=1&c=<%=gnCustomerID%>&a=<%=gnAddressID%>';">Manager Override</button>
<%
	Else
%>
								<button onclick="gotoMgrOverride()">Manager Override</button>
<%
	End If
%>
								</p>
<%
End If
%>
						</div>
						<div id="managerdiv" style="position: absolute; left: 0px; top: 0px; width: 1010px; height: 605px; visibility: hidden; background-color: #fbf3c5;">
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
