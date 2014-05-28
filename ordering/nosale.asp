<%
Option Explicit
Response.buffer = TRUE

If Session("SecurityID") = "" Then
	Response.Redirect("/default.asp")
End If

If Session("SecurityID") < 2 Then
	Response.Redirect("/default.asp")
End If

If Not ((Left(Request.ServerVariables("remote_addr"), 9) = "10.0.254." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.1." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.2.") OR Session("swipe")) Then
	Response.Redirect("/default.asp")
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
Dim gnStoreID, gsIPAddress, gsPrinterIPAddress, gbIsCashDrawer2, gsPrintData
Dim ganOrderIDs(), gsLocalErrorMsg, i
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

gnStoreID = Session("StoreID")
gsIPAddress = Request.ServerVariables("REMOTE_ADDR")

If Not GetStoreStationCashDrawer(gnStoreID, gsIPAddress, gsPrinterIPAddress, gbIsCashDrawer2) Then
	If Len(gsDBErrorMessage) > 0 Then
		Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
	Else
		Response.Redirect("/error.asp?err=" & Server.URLEncode("Cash Drawer Has Not Been Defined For This Station"))
	End If
End If

' Pop cash drawer
If gbIsCashDrawer2 Then
	SendToPrinter gsPrinterIPAddress, CHR(27) & CHR(112) & CHR(1) & CHR(60) & CHR(60)
Else
	SendToPrinter gsPrinterIPAddress, CHR(27) & CHR(112) & CHR(0) & CHR(60) & CHR(60)
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
<script src="/include2/redirect2.js" type="text/javascript"></script>
<script type="text/javascript">
<!--
var ie4=document.all;

function resetRedirect() {
	var loRedirectDiv;
	
	loRedirectDiv = ie4? eval("document.all.redirect") : document.getElementById("redirect");
	loRedirectDiv.innerHTML = <%=gnRedirectTime%>;
}
//-->
</script>
</head>

<body onload="clockInit(clockLocalStartTime, clockServerStartTime); clockOnLoad(); redirect('/default.asp')" onunload="clockOnUnload()">

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
						<span id="ClockTime" onclick="clockToggleSeconds()"><%=clockTimeString(Hour(gDate), Minute(gDate), Second(gDate))%></span> | 
						<span class="counter" id="redirect"><%=gnRedirectTime%></span>
					</div>
				</td>
			</tr>
			<tr height="733">
				<td valign="top" width="1010">
					<table cellpadding="0" cellspacing="0" width="1010" height="723">
						<tr>
							<td align="center" valign="top" width="1010">
								<p><font size="72"><strong>NO SALE</strong></font></p>
								<p>&nbsp;</p>
								<p>&nbsp;</p>
								<p><font size="72"><strong>CLOSE DRAWER</strong></font></p>
								<button style="width: 680px;" onclick="window.location = '/main.asp'">Done</button>
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
