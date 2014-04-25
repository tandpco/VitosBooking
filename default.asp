<%
Option Explicit
Response.buffer = TRUE

If Request("top").count = 0 Then
	Session.Contents.RemoveAll
	Session.Abandon
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
Dim gnStoreID, gsTop, gsLeft, gsMoveLeft, gsMoveUp, gdtTransactionDate

If Request("top").count = 0 Then
	gsTop = "0px"
Else
	gsTop = Request("top")
End If

If Request("left").count = 0 Then
	gsLeft = "0px"
Else
	gsLeft = Request("left")
End If

If Request("moveleft").count = 0 Then
	gsMoveLeft = "false"
Else
	If Request("moveleft") = "true" Then
		gsMoveLeft = "true"
	Else
		gsMoveLeft = "false"
	End If
End If

If Request("moveup").count = 0 Then
	gsMoveUp = "false"
Else
	If Request("moveup") = "true" Then
		gsMoveUp = "true"
	Else
		gsMoveUp = "false"
	End If
End If

gnStoreID = GetStoreByNetwork(Request.ServerVariables("REMOTE_ADDR"))
If gnStoreID = -1 Then
	Response.Redirect("Error.asp")
End If

gdtTransactionDate = GetStoreTransactionDate(gnStoreID)
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
var gnInterval = 25; // How often to move in milliseconds
var gnStepSize = 1; // How much to move

var goTimerID = null;
var goTimerID2 = null;
var gbMoveUp = <%=gsMoveUp%>;
var gbMoveLeft = <%=gsMoveLeft%>;
var ie4=document.all;
var gsSwipeCode = "";

function moveDiv() {
	var loDiv, loTable, lnTop, lnLeft;
	
	loDiv = ie4? eval("document.all.screensaver") : document.getElementById('screensaver');
	loTable = ie4? eval("document.all.frame") : document.getElementById('frame');

	if (gbMoveUp) {
		if (parseInt(loDiv.style.top) > 0) {
			loDiv.style.top = (parseInt(loDiv.style.top) - gnStepSize).toString() + "px";
			if (parseInt(loDiv.style.top) < 0) {
				loDiv.style.top = "0px";
			}
		}
		else {
			gbMoveUp = false;
			loDiv.style.top = (parseInt(loDiv.style.top) + gnStepSize).toString() + "px";
		}
	}
	else {
		if ((parseInt(loDiv.style.top) + parseInt(loDiv.style.height)) < parseInt(loTable.style.height)) {
			loDiv.style.top = (parseInt(loDiv.style.top) + gnStepSize).toString() + "px";
			if ((parseInt(loDiv.style.top) + parseInt(loDiv.style.height)) > parseInt(loTable.style.height)) {
				loDiv.style.top = (parseInt(loTable.style.height) - parseInt(loDiv.style.height)).toString() + "px";
			}
		}
		else {
			gbMoveUp = true;
			loDiv.style.top = (parseInt(loDiv.style.top) - gnStepSize).toString() + "px";
		}
	}
	
	if (gbMoveLeft) {
		if (parseInt(loDiv.style.left) > 0) {
			loDiv.style.left = (parseInt(loDiv.style.left) - gnStepSize).toString() + "px";
			if (parseInt(loDiv.style.left) < 0) {
				loDiv.style.left = "0px";
			}
		}
		else {
			gbMoveLeft = false;
			loDiv.style.left = (parseInt(loDiv.style.left) + gnStepSize).toString() + "px";
		}
	}
	else {
		if ((parseInt(loDiv.style.left) + parseInt(loDiv.style.width) + gnStepSize) < parseInt(loTable.style.width)) {
			loDiv.style.left = (parseInt(loDiv.style.left) + gnStepSize).toString() + "px";
			if ((parseInt(loDiv.style.left) + parseInt(loDiv.style.width) + gnStepSize) > parseInt(loTable.style.width)) {
				loDiv.style.left = (parseInt(loTable.style.width) - parseInt(loDiv.style.width)).toString() + "px";
			}
		}
		else {
			gbMoveLeft = true;
			loDiv.style.left = (parseInt(loDiv.style.left) - gnStepSize).toString() + "px";
		}
	}
	
	goTimerID = setTimeout("moveDiv()", gnInterval);
}

function stopMoving() {
	if (goTimerID) {
		clearTimeout(goTimerID);
		goTimerID = null;
	}
}

function reloadPage() {
	stopReloading();
	stopMoving();
 	clockOnUnload();
 	
	var loDiv = ie4? eval("document.all.screensaver") : document.getElementById('screensaver');
	if (gbMoveLeft) {
		if (gbMoveUp) {
			window.location = "default.asp?top=" + loDiv.style.top + "&left=" + loDiv.style.left + "&moveleft=true&moveup=true";
		}
		else {
			window.location = "default.asp?top=" + loDiv.style.top + "&left=" + loDiv.style.left + "&moveleft=true&moveup=false";
		}
	}
	else {
		if (gbMoveUp) {
			window.location = "default.asp?top=" + loDiv.style.top + "&left=" + loDiv.style.left + "&moveleft=false&moveup=true";
		}
		else {
			window.location = "default.asp?top=" + loDiv.style.top + "&left=" + loDiv.style.left + "&moveleft=false&moveup=false";
		}
	}
}

function stopReloading() {
	if (goTimerID2) {
		clearTimeout(goTimerID2);
		goTimerID2 = null;
	}
}

function ValidateForm(theForm) {
	if (theForm.EmployeeID.value == "")
	{
		theForm.EmployeeID.focus();
		return (false);
	}
	
	theForm.parentElement.style.visibility = "hidden";
	return (true);
}
//-->
</script>
</head>

<body onload="document.form1.EmployeeID.focus(); clockInit(clockLocalStartTime, clockServerStartTime); clockOnLoad(); moveDiv(); goTimerID2 = setTimeout('reloadPage()', 60000);" onunload="stopReloading(); stopMoving(); clockOnUnload();">

<table align="center" cellspacing="0" cellpadding="0" id="frame" style="width: 1020px; height: 765px;" onclick="window.location='/signon.asp'">
	<tr>
		<td valign="top" width="1010" height="765">
			<div id="screensaver" style="text-align: center; position: relative; top: <%=gsTop%>; left: <%=gsLeft%>; width: 400px; height: 325px; border-width: 1px; border-style: solid; border-color: #006D31;">
				<img alt="Vito's Pizza and Subs" height="191" src="images/logo.jpg" width="400" />
				<span id="ClockDate"><%=clockDateString(gDate)%></span>
				<span id="ClockTime" onclick="clockToggleSeconds()"><%=clockTimeString(Hour(gDate), Minute(gDate), Second(gDate))%></span><br />
				<br /><strong>
<%
If gbTestMode Then
	If gbDevMode Then
%>
						DEV SYSTEM - 
<%
	Else
%>
						TEST SYSTEM - 
<%
	End If
End If
%>
				Swipe Card or Touch Screen Anywhere To Begin</strong>
				<form name="form1" method="post" action="main.asp" onSubmit="return ValidateForm(this);">
					<input type="hidden" name="StoreID" value="<%=gnStoreID%>" />
					<input type="password"  name="EmployeeID" id="EmployeeID"  autofocus autocomplete="off" style="text-align: center; background: #fbf3c5; outline: none;"/>
				</form>
			</div>
		</td>
	</tr>
</table>

</body>

</html>
<!-- #Include Virtual="include2/db-disconnect.asp" -->
