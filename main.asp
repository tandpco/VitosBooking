<%
Option Explicit
Response.buffer = TRUE
%>
<!-- #INCLUDE FILE="include2/utility.asp" -->
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
<!-- #Include Virtual="include2/clearorder.asp" -->
<!-- #INCLUDE FILE="include2/litmos.asp" -->

<%
'------------------------------------------------------------Security Check-------------------------------------------
If  (Request("EmployeeID") = "" And Session("SecurityID") ="") Or (Session("StoreID")="" And Request("StoreID")="") Then Response.Redirect("default.asp")
'------------------------------------------------------------------------------------------------------------------------

If Session("StoreID") = "" Then
	Session("StoreID") = CLng(Request("StoreID"))
End If

OpenSQLConn 'Open Database

Dim  SQL, RSID, RS, NotificationMsg

If Session("EmpID") = "" Or Len(Trim(Request("EmployeeID"))) > 0 Then 
	If Len(Trim(Request("EmployeeID"))) = 0 Then
		Response.Redirect("signon.asp")
	Else
		Session("Swipe") = FALSE
		If Right(Trim(Request("EmployeeID")),1) = "?" Then

'			SQL="Select * from tblEmployee where CardID = '" & Request("EmployeeID") & "' and IsActive='True'" 'If used swipe card
			SQL="Select * from tblEmployee where CardID = '" & Request("EmployeeID") & "' and (SystemRoleID >= 4 Or StoreID = " & Session("StoreID") & ") and IsActive='True'" 'If used swipe card
			Set RSID=Conn.Execute(SQL)
			If RSID.BOF And RSID.EOF Then

NotificationMsg = "That swipe card was not found or not authorized"
Response.Redirect("notification.asp?msg=" & Server.URLEncode(NotificationMsg)&"&NextPage=default.asp")

			Else
					Session("EmpID") =RSID("EmpID")
					Session("EmployeeID") = RSID("EmployeeID")
					Session("Swipe") = TRUE
			End If

		Else
			If IsNumeric(Request("EmployeeID")) Then
				Session("EmployeeID") =Request("EmployeeID") 'If entered employeeID
			Else
NotificationMsg = "Invalid entry."
Response.Redirect("notification.asp?msg=" & Server.URLEncode(NotificationMsg)&"&NextPage=default.asp")
			End If
		End If
	End If
End If

'----------------------------------------------------------- GET EMPLOYEE INFORMATION ----------------------------------------

Dim  RSEmp

SQL="Select * from tblEmployee where EmployeeID = " & Session("EmployeeID") &" and StoreID = "&Session("StoreID")&" and IsActive='True'"

Set RSEmp=Conn.Execute(SQL)

If RSEmp.EOF And RSEmp.BOF Then  'Not Found REDIRECT back to sign in page
	RSEmp.Close
	
	SQL="Select * from tblEmployee where EmployeeID = " & Session("EmployeeID") &" and SystemRoleID >= 4 and IsActive='True'"
	
	Set RSEmp=Conn.Execute(SQL)
End If

If RSEmp.EOF And RSEmp.BOF Then  'Not Found REDIRECT back to sign in page
	NotificationMsg = "That Employee ID was not found or not authorized"
	Response.Redirect("notification.asp?msg=" & Server.URLEncode(NotificationMsg)&"&NextPage=default.asp")
End If

Session("EmpID") = RSEmp("EmpID")
'----------------------------------------------------------- DETERMINE AGE --------------------------------------------------------

If IsNull(RSEmp("Birthdate")) Then
	NotificationMsg = "You cannot sign in without a valid birthdate entered in the system, please contact your supervisor"
	Response.Redirect("notification.asp?msg=" & Server.URLEncode(NotificationMsg)&"&NextPage=default.asp")
End If

	Dim intAge
	intAge = DateDiff("yyyy", RSEmp("Birthdate"), now())
    If Now() < DateSerial(Year(now()), Month(RSEmp("Birthdate")), Day(RSEmp("Birthdate"))) Then
        intAge = intAge - 1
    End If
session("intAge") = intAge

'--------------------------------------------------------- Determine Security Level -----------------------------------------------
Dim RSSecurity
SQL="Select SystemRoleID AS SecurityID from tblEmployee where EmpID = "& Session("EmpID")
Set RSSecurity = Conn.Execute(SQL)

If IsNull(RSSecurity("SecurityID")) Then
	NotificationMsg = "That Employee has not been configured with a role, please contact your supervisor"
	Response.Redirect("notification.asp?msg=" & Server.URLEncode(NotificationMsg)&"&NextPage=default.asp")
End If

If RSSecurity("SecurityID") <> 4 And RSEmp("StoreID") <> Session("StoreID") Then
	NotificationMsg = "That Employee ID was not found or not authorized at this store."
	Response.Redirect("notification.asp?msg=" & Server.URLEncode(NotificationMsg)&"&NextPage=default.asp")
End If

Session("EmpID") =RSEmp("EmpID")
Session("EmployeeID") =RSEmp("EmployeeID")
Session("SecurityID") = RSSecurity("SecurityID")
Session("Name") = RSEmp("FirstName") & " " & RSEmp("LastName")
Session("DriverStatus") = RSEmp("DriverStatus")

'-------------------------------------------------------- CHECK TO SEE IF STORE IS OPEN -----------------------------------------------------
Dim SQLDateCheck, RSDateCheck, gnOpenTime, gnCloseTime, gsCloseTime

SQLDateCheck = "Select top 1 CurrentStatus, ReportDate from tblStoreReportDate where StoreID = "&Session("StoreID")&" Order by RADRAT DESC"
Set RSDateCheck = Conn.Execute(SQLDateCheck)
If  (RSDateCheck.BOF And RSDateCheck.EOF) Then 'No record, Need to open Store
	Response.Redirect("store.asp?Action=Open")
ElseIf Trim(RSDateCheck("CurrentStatus")) = "Closed" Then
	Response.Redirect("store.asp?Action=Open")
Else
	Session("TransactionDate") = Right("0" & CStr(Month(RSDateCheck("ReportDate"))), 2) & "/" & Right("0" & CStr(Day(RSDateCheck("ReportDate"))), 2) & "/" & CStr(Year(RSDateCheck("ReportDate")))
	
	If GetStoreOpenCloseTime(Session("StoreID"), Weekday(DateValue(Session("TransactionDate"))), gnOpenTime, gnCloseTime) Then

    'Response.Write(gnOpenTime)
    'Response.Write("---")
    'Response.Write(gnCloseTime)
    'Response.Write("---")
		If gnCloseTime = 2400 Then
			gsCloseTime = DateAdd("d", 1, DateValue(Session("TransactionDate"))) & " 12:00 AM"
		Else
			If gnCloseTime < gnOpenTime Then
				If gnCloseTime < 1200 Then
					gsCloseTime = DateAdd("d", 1, DateValue(Session("TransactionDate"))) & " " & Int(gnCloseTime / 100) & ":" & Right("0" & (gnCloseTime Mod 100), 2) & " AM"
				Else
					If gnCloseTime < 1300 Then
						gsCloseTime = DateAdd("d", 1, DateValue(Session("TransactionDate"))) & " 12:" & Right("0" & (gnCloseTime Mod 100), 2) & " PM"
					Else
						gsCloseTime = DateAdd("d", 1, DateValue(Session("TransactionDate"))) & " " & (Int(gnCloseTime / 100) - 12) & ":" & Right("0" & (gnCloseTime Mod 100), 2) & " PM"
					End If
				End If
			Else
				If gnCloseTime < 1200 Then
					gsCloseTime = Session("TransactionDate") & " " & Int(gnCloseTime / 100) & ":" & Right("0" & (gnCloseTime Mod 100), 2) & " AM"
				Else
					If gnCloseTime < 1300 Then
						gsCloseTime = Session("TransactionDate") & " 12:" & Right("0" & (gnCloseTime Mod 100), 2) & " PM"
					Else
						gsCloseTime = Session("TransactionDate") & " " & (Int(gnCloseTime / 100) - 12) & ":" & Right("0" & (gnCloseTime Mod 100), 2) & " PM"
					End If
				End If
			End If
		End If
		
		If Now >= DateAdd("h", gnCloseOutThreshold, CDate(gsCloseTime)) Then
      'Response.Write(Session("StoreID"))
      'Response.Write("---")
      'Response.Write(gsCloseTime)
      'Response.Write("---")
      'Response.Write(Weekday(DateValue(Session("TransactionDate"))))
      'Response.Write("---")
      'Response.Write(DateAdd("h", gnCloseOutThreshold, CDate(gsCloseTime)))
      'Response.End
			'Response.Redirect("closepastdue.asp")
		End If
	End If
End If


'Determine Current Employee Punch Status
SQL="SELECT  ShiftID, EmpID, StoreID, StoreRoleID, Rate, PunchInType, PunchInTime, PunchOutType, PunchOutTime, ReportDate, Modified, RADRAT FROM tblShifts WHERE (EmpID = "&Session("EmpID")&") AND (PunchOutType IS NULL)"

Dim RSPunch

Set RSPunch=Conn.Execute(SQL)

If RSPunch.BOF And RSPunch.EOF Then
Session.Contents.Remove("CurrentShiftID")
Else
Session("CurrentShiftID") = RSPunch("ShiftID")
End If

'--------------------------------------------------------- DETERMINE IF DRIVER CAN CASH OUT -----------------------------------------------
Dim sqlActiveDriver, RSActiveDriver, OkToSignOut

sqlActiveDriver="SELECT DeliveryID FROM tblDelivery WHERE (EmpID = "&Session("EmpID")&") AND (CashedOutTime IS NULL)"

Set RSActiveDriver=Conn.Execute(sqlActiveDriver)

If RSActiveDriver.BOF And RSActiveDriver.EOF Then 'cashed out
	OkToSignOut = "True"
Else
	OkToSignOut = "False"
End If


Dim ganOrderIDs(), gsLocalErrorMsg
Dim i

Dim gbCoursesDue

gbCoursesDue = FALSE
If Session("LitmosCheck") <> "CHECKED" Then
	If IsCoursesDue(Session("EmpID"), Session("StoreID"), gbCoursesDue) Then
		Session("LitmosCheck") = "CHECKED"
        'Session("LitmosCheck") = "CHECKEVERYTIME"
		If gbCoursesDue Then
			Response.Redirect("Litmos/Litmos.aspx?EmpID=" & Session("EmpID") & "&StoreID=" & Session("StoreID"))
		End If
	Else
		Response.Redirect("notification.asp?msg=" & Server.URLEncode(gsDBErrorMessage)&"&NextPage=default.asp")
	End If
End If

If Right(LCase(Request.ServerVariables("HTTP_REFERER")), 11) = "default.asp" Or Right(LCase(Request.ServerVariables("HTTP_REFERER")), 10) = "signon.asp" Then
	Response.Redirect("ordering/neworder.asp")
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Vitos - Management Console</title>
<LINK href="css/vitosmain.css" type=text/css rel=stylesheet>
<script src="include2/redirect2.js" type="text/javascript"></script>

<!-- #Include File="include2/clock-server.asp" -->

<script type="text/javascript">
<!--
var ie4=document.all;

function resetRedirect(pnHowLong) {
	var loRedirectDiv;
	
	loRedirectDiv = ie4? eval("document.all.redirect") : document.getElementById("redirect");
	loRedirectDiv.innerHTML = pnHowLong.toString();
}

function iframeclick() {
document.getElementById("litmosiframe").contentWindow.document.body.onclick = function() {
		window.location = "Litmos.aspx?EmpID=<%=Session("EmpID")%>&StoreID=<%=Session("StoreID")%>";
    }
}
//-->
</script>
    <script src="Scripts/jquery-1.8.2.min.js" type="text/javascript"></script>
    <script src="Scripts/jquery-ui-1.9.0.custom.min.js" type="text/javascript"></script>
    <script src="Scripts/OpenNewWindowWithinPopup.js" type="text/javascript"></script>
    <link href="Styles/NewWindowPopup.css" rel="stylesheet" type="text/css" />
    <link href="Styles/jquery-ui-1.9.0.custom.css" rel="stylesheet" type="text/css" />
    <script src="Scripts/JScript.js" type="text/javascript" language="javascript"></script>
    <link href="Styles/SchedulingOptions.css" rel="stylesheet" type="text/css" />
    <link href="Styles/GridView.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript">
        $(document).ready(function () {

            $('#btnRequestOff').click(function () {
            	resetRedirect(180);
            	
                $.OpenNewWindowPopUp('What day would you like to request off?', 'RequestOff.aspx?EmpID=<%=Session("EmpID")%>&StoreID=<%=Session("StoreID")%>', 500, 410);
            });




        });
        //if you want to refresh parent page on closing of a popup window then remove comment to the below function
        //and also call this function from the js file 
        //        function Refresh() {
        //            window.location.reload();
        //        }
    </script>
</head>
<body onLoad="clockInit(clockLocalStartTime, clockServerStartTime); clockOnLoad(); redirect('/default.asp')" onUnload="clockOnUnload()" oncontextmenu="return false;">
<table cellspacing="0" cellpadding="1" width="1010" bgcolor="#006d31" border="0" align="center">
  <tbody>
    <tr>
      <td><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#fbf3c5" style="WIDTH: 100%">
        <tr>
          <td valign="top"><table width="1010" border="0" align="center" cellpadding="0" cellspacing="3">
            <tr>
              <td align="left" valign="middle"><div align="left"><img src="images/logo_sm.jpg" alt="" width="200" height="96" hspace="10" vspace="10" /></div></td>
              <td align="left" valign="middle"><%If Session("SecurityID") > 1 Then%>
<%
'--------------------------------- determine current/daily sales information
Dim SQLDaily, RSDaily, sqlOnHold, RSOnHold, sqlDeliveryCharge, RSDeliveryCharge, DeliveryCharge

SQLDaily="SELECT SUM(quantity*(dbo.tblOrderLines.Cost - dbo.tblOrderLines.Discount)) AS DailyTotal, COUNT(DISTINCT dbo.tblorders.orderid) AS NumOrders FROM dbo.tblorders INNER JOIN dbo.tblOrderLines ON dbo.tblOrders.orderid = dbo.tblOrderLines.orderid WHERE (dbo.tblOrders.storeid = "&session("StoreID")&") AND dbo.tblOrders.transactiondate= '"&Session("TransactionDate")&"' AND (dbo.tblOrders.OrderStatusID >= 3 and dbo.tblOrders.OrderStatusID <= 10)"

Set RSDaily = Conn.Execute(SQLDaily)

'--------------added 10/28
sqlDeliveryCharge = "SELECT SUM(DeliveryCharge) AS TotalDeliveryCharge FROM dbo.tblOrders WHERE (StoreID = "&session("StoreID")&") AND (TransactionDate = '"&Session("TransactionDate")&"') AND (OrderStatusID >= 3) AND (OrderStatusID <= 10)" 
Set RSDeliveryCharge = Conn.Execute(sqlDeliveryCharge)

If IsNull(RSDeliveryCharge("TotalDeliveryCharge")) Then 'None found
DeliveryCharge = 0
Else
DeliveryCharge = RSDeliveryCharge("TotalDeliveryCharge")
End If
'------------------end added 10/28

'--------------------------------------- DETERMINE CURRENT HOURS AND WAGES ------------------------------------
Dim sqlTimeWorkedCurrent, rsTimeWorkedCurrent, MinutesToAdd, WagesToAdd, TotalMinutesWorked, TotalWagesEarned, TotalTimeWorked, CurrentLaborPercentage

sqlTimeWorkedCurrent = "Select Rate, PunchInTime, PunchOutTime from tblShifts where storeID = "&Session("StoreID")&" AND dbo.tblShifts.ReportDate= '"&Session("TransactionDate")&"'"

Set rsTimeWorkedCurrent = Conn.Execute(sqlTimeWorkedCurrent)
If rsTimeWorkedCurrent.BOF And rsTimeWorkedCurrent.EOF Then 'No punches for that day

Else
Do While Not rsTimeWorkedCurrent.EOF
'Let's calculate in minutes then divide it out

If rsTimeWorkedCurrent("PunchOutTime") <> "" Then 'Complete Shift, can calculate shift

MinutesToAdd = Round(DateDiff("n",rsTimeWorkedCurrent("PunchInTime"),rsTimeWorkedCurrent("PunchOutTime")),2)
WagesToAdd = Round((Round((DateDiff("n", rsTimeWorkedCurrent("PunchInTime"), rsTimeWorkedCurrent("PunchOutTime")) / 60), 2) * rsTimeWorkedCurrent("Rate")*gnLaborFactor), 2)

Else 'Still Punched In, calculate between time punched in and NOW

MinutesToAdd = Round(DateDiff("n",rsTimeWorkedCurrent("PunchInTime"),Now()),2)
WagesToAdd = Round((Round((DateDiff("n", rsTimeWorkedCurrent("PunchInTime"), Now()) / 60), 2) * rsTimeWorkedCurrent("Rate")*gnLaborFactor), 2)
End If

TotalMinutesWorked = TotalMinutesWorked + MinutesToAdd
TotalWagesEarned = TotalWagesEarned + WagesToAdd

rsTimeWorkedCurrent.MoveNext
Loop

End If
		If TotalMinutesWorked = "" Then TotalMinutesWorked = 0
		If TotalMinutesWorked < 60 Then 
			TotalTimeWorked = TotalMinutesWorked & " Minutes"
		Else
			TotalTimeWorked =  Round(TotalMinutesWorked/60,2)
		End If

If (RSDaily("DailyTotal")+DeliveryCharge) = 0 Or TotalWagesEarned = 0 Or IsNull(RSDaily("DailyTotal")) Then
	CurrentLaborPercentage = 0
Else
	CurrentLaborPercentage = formatpercent((TotalWagesEarned)/(RSDaily("DailyTotal")+DeliveryCharge),2)
End If
	If CurrentLaborPercentage = "0" Then CurrentLaborPercentage = "&#8734"


'--------------------------------------- END DETERMINE CURRENT HOURS AND WAGES ------------------------------------

'------------------------------------ DETERMINE NUMBER OF ON HOLD ORDERS -----------------------------
sqlOnHold = "Select count(OrderID) as NumOnHold from tblOrders where OrderStatusID = 2 and StoreID = " & Session("StoreID")
Set RSOnHold = Conn.Execute(sqlOnHold)


%>
                <table cellspacing="0" cellpadding="1" width="100%" bgcolor="#006d31" border="0" align="center">
                    <tr>
                      <td><table width="100%" border="0" cellpadding="0" cellspacing="3" bgcolor="fbf3c5">
					            <tr>
                                <td align="CENTER" bgcolor="#006D31" colspan="2"><b><font color="#ffffff">Transaction Date - <%=Session("TransactionDate")%></b></span></td>
                              </tr>
                          <tr>
                            <td align="LEFT" width="50%" valign="top"><span class="smallblack11">
                              <% If RSDaily("NumOrders") > 0 Then %>
                              Num. Orders: <%=RSDaily("NumOrders")%><br />
                              Sales: <%=FormatCurrency(RSDaily("DailyTotal")+DeliveryCharge)%><br />
                              Average Ticket: <%=FormatCurrency((cdbl(RSDaily("DailyTotal")+DeliveryCharge)/CInt(RSDaily("NumOrders"))))%><br />
                              <%Else%>
                             No orders for today yet.<br />
                              <%End IF%>
                              Orders On Hold: <%=RSOnHold("NumOnHold")%>
                            </span></td>
							<TD valign="top"><span class="smallblack11">Labor Hours: <%=TotalTimeWorked%> <br>Labor Wages: <%=FormatCurrency(TotalWagesEarned)%><br>Labor %: <%=CurrentLaborPercentage%></span></td>
                          </tr>
                      </table></td>
                    </tr>
                </table>
                <%End If%></td>
              <td width="33%" align="center"><div align="right"><b>
<%
If gbTestMode Then
	If gbDevMode Then
%>
						DEV SYSTEM
<%
	Else
%>
						TEST SYSTEM
<%
	End If
End If
%>
              		Store: <%=Session("StoreID")%></b><br />
						<b><%=Session("Name")%></b><br />
                      <div id="ClockDate"><%=clockDateString(gDate)%></div>
                      <div id="ClockTime" onClick="clockToggleSeconds()"><%=clockTimeString(Hour(gDate), Minute(gDate), Second(gDate))%></div>
                      <div class="counter" id="redirect"><%=gnRedirectTime%></div>
              </div></td>
            </tr>
            <tr>
              <td colspan="3" align="center"><table width="100%" border="0" cellspacing="0" cellpadding="3">
                  <tr>
                    <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                        <tr>
                          <td width="508" align="left"><a href="ordering/neworder.asp"><img src="images/btn_placeorder.jpg" alt="Place Order" width="75" height="75" border="0" /></a>
						  <a href="opentickets.asp"><img src="images/btn_opentickets.jpg" alt="Open Tickets" width="75" height="75"  border="0"/></a>
						  <a href="driverdispatch.asp"><img src="images/btn_driverdispatch.jpg" alt="Driver Dispatch" width="75" height="75" border="0" /></a>

<% 
If RSEmp("StoreID") = Session("StoreID") Then
If Session("CurrentShiftID") = "" Then 'Not Punched In

	'Need to determine if punched out or just on break
	Dim sqlBreakCheck, RSBreakCheck
	sqlBreakCheck = "Select TOP 1 dbo.tblShifts.* from tblShifts where EmpID = '"&Session("EmpID")&"' and ReportDate = '"&Session("TransactionDate")&"' Order By RADRAT Desc;"
	Set RSBreakCheck = Conn.Execute(sqlBreakCheck)
	If RSBreakCheck.BOF And RSBreakCheck.EOF Then
%>
		<a href="punch.asp?PunchType=In" class="navs"><img src="images/btn_punchIN.jpg" alt="Punch In" width="75" height="75" border="0" /></a>
<%
	Else
		If  RSBreakCheck("PunchOutType") = "Break" Then
%>
		<a href="break.asp?Break=OFF&EmpID=<%=Session("EmpID")%>&OdometerIn=<%=RSBreakCheck("OdometerIn")%>"><img src="images/btn_ReturnFrombreak.jpg" alt="Return From Break" width="75" height="75" border="0" /></a>
<%
		Else
%>
				<a href="punch.asp?PunchType=In" class="navs"><img src="images/btn_punchIN.jpg" alt="Punch In" width="75" height="75" border="0" /></a>
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
				<a href="punch.asp?PunchType=Out&StoreRoleID=<%= RSPunch("StoreRoleID")%>"><img src="images/btn_punchOUT.jpg" alt="Punch Out" width="75" height="75" border="0" /></a>
<%
			End If
%>
				<a href="break.asp?Break=ON&EmpID=<%=Session("EmpID")%>"><img src="images/btn_goOnbreak.jpg" alt="Go On Break" width="75" height="75" border="0" /></a>
<%
		Else 'Out on Delivery
%>
			<img src="images/btn_outondelivery.jpg" alt="Out On Delivery" width="75" height="75" border="0" />
<%
		End If 
End If
End If
%>
  						  <a><img src="images/btn_requesttimeoff.jpg" id="btnRequestOff" alt="Request Time off" width="75" height="75" border="0" /></a>
                          <div id="divTimeOff">
                          </div>
						  <a href="orderlookup.asp"><img src="images/btn_orderlookup.jpg" alt="Order Lookup" width="75" height="75" border="0" /></a>
						  <a href="ordersearch.asp"><img src="images/btn_ordersearch.jpg" alt="Order Search" width="75" height="75" border="0" /></a>
						  <%'<a href="requesttimeoff.asp"><img src="images/btn_requesttimeoff.jpg" alt="Request Time Off" width="75" height="75" border="0" /></a>%>
<%If RSEmp("DriverStatus") = "Active" Or Session("SecurityID") > 1 Then%>
						  <a href="viewdrives.asp"><img src="images/btn_drives.jpg" alt="Drives" width="75" height="75" border="0" /></a>
<%End If%>
<%'					<a href="training.asp"><img src="images/btn_training.jpg" alt="Training" width="75" height="75" border="0" /></a>%>
<%'					<a href="printemployeeschedule.asp"><img src="images/btn_printemployeeschedule.jpg" alt="Print Employee Schedule" width="75" height="75" border="0" /></a> %>
  						  <a href="listingofstores.asp"><img src="images/btn_listingofstores.jpg" alt="Listing of Stores" width="75" height="75" border="0"/></a>
  						  <a href="printemployeehours.asp"><img src="images/btn_printemployeehours.jpg" alt="Print Employee Hours" width="75" height="75" border="0"/></a>
  						  <a href="issues.asp"><img src="images/btn_issues.jpg" alt="Issues" width="75" height="75" border="0"/></a>
</td>
                        </tr>
                        <tr>
                          <td height="1" align="left" bgcolor="#ea8223"></td>
                        </tr>
<%If Session("SecurityID") > 1 Then%>
                        <tr>
                          <td align="left">
						  <a href="punchreport.asp"><img src="images/btn_punchreport.jpg" alt="Punch Report" width="75" height="75" border="0" /></a>
						  <a href="employeelist.asp"><img src="images/btn_showemployeelist.jpg" alt="Show Employee List" width="75" height="75" border="0" /></a>
						  <a href="sendmessage.asp"><img src="images/btn_sendmessage.jpg" alt="Send Message" width="75" height="75" border="0" /></a>
				  <a href="payinout.asp"><img src="images/btn_payinout.jpg" alt="Pay In/Out" width="75" height="75" border="0" /></a>
				  <a href="storeselectpayonaccount.asp"><img src="images/btn_payonaccount.jpg" alt="Pay On Account" width="75" height="75" border="0" /></a>
<%
	If (Left(Request.ServerVariables("remote_addr"), 9) = "10.0.254." Or Left(Request.ServerVariables("remote_addr"), 7) = "10.0.0." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.1." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.2.") OR Session("Swipe") Then
%>
				  <a href="ordering/nosale.asp"><img src="images/btn_nosale.jpg" alt="No Sale" width="75" height="75" border="0" /></a>
<%
	End If
%>
				  <a href="https://vmd.vitos.com:4159"><img src="images/btn_vmd.jpg" alt="Vito's Management Dash" width="75" height="75" border="0" /></a>
<br />
						  <a href="makedriveractive.asp"><img src="images/btn_makedriveractive.jpg" alt="Make Driver Active" width="75" height="75" border="0" /></a>
						  <a href="makedriverinactive.asp"><img src="images/btn_makedriverinactive.jpg" alt="Make Driver Inactive" width="75" height="75" border="0" /></a>
						  <a href="opendrawers.asp"><img src="images/btn_opendrawers.jpg" alt="Open Drawers" width="75" height="75" border="0" /></a>
						  <a href="closestore.asp?Step=1"><img src="images/btn_closestore.jpg" alt="Close Store" width="75" height="75" border="0" /></a>
						  <a href="inventorystore/"><img src="images/btn_manageinventory.jpg" alt="Manage Inventory" width="75" height="75" border="0"/></a>
						  <a href="dailystatistics.asp"><img src="images/btn_dailystatistics.jpg" alt="Daily Statistics" width="75" height="75" border="0" /></a>
						  <a href="makedeposit.asp"><img src="images/btn_makedeposit.jpg" alt="Make Deposit" width="75" height="75" border="0" /></a>
						  <a href="viewdeposits.asp"><img src="images/btn_viewdeposits.jpg" alt="View Deposits" width="75" height="75" border="0" /></a>
						  </td>
                        </tr>
                        <tr>
                          <td height="1" align="left" bgcolor="#3d96cc"></td>
                        </tr>
<%End If%>
<%If Session("SecurityID") > 2 Then%>
                        <tr>
                          <td align="left">
<%'					  <a href="manageschedules.asp"><img src="images/btn_manageschedules.jpg" alt="Manage Schedules" width="75" height="75" border="0" /></a>%>
					  <a href="reports.asp"><img src="images/btn_reports.jpg" alt="Reports" width="75" height="75" border="0" /></a>
						  <a href="storemanageaccounts.asp"><img src="images/btn_managestoreaccounts.jpg" alt="Manage Store Accounts" width="75" height="75" border="0" /></a>
						  <a href="rpt_Payroll.asp"><img src="images/btn_Viewpayroll.jpg" alt="View Payroll" width="75" height="75" border="0" /></a>
						  <a href="releasenotes.asp"><img src="images/btn_releasenotes.jpg" alt="POSSUM Release Notes" width="75" height="75" border="0" /></a>
<!--
  						  <a href="javascript: window.open('', '_self'); window.close();">Close Window</a>
-->
  						  </td>
                        </tr>
                        <tr>
                          <td height="1" align="left" bgcolor="#80993d"></td>
                        </tr>
<%End If%>
<%If Session("SecurityID") > 3 Then%>

<%
If (Left(Request.ServerVariables("remote_addr"), 9) = "10.0.254." Or Left(Request.ServerVariables("remote_addr"), 7) = "10.0.0." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.1." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.2." Or Left(Request.ServerVariables("remote_addr"), 7) = "10.0.1." Or Left(Request.ServerVariables("remote_addr"), 7) = "10.0.2.") OR Session("swipe") Then
%>
                        <tr>
                          <td align="left">
						  <a href="adminscreens/managestores.asp"><img src="images/btn_managestores.jpg" alt="Manage Stores" width="75" height="75" border="0" /></a>
						  <a href="adminscreens/managesystemroles.asp"><img src="images/btn_managesystemroles.jpg" alt="Manage System Roles" width="75" height="75" border="0" /></a>
						  <a href="adminscreens/managestoreroles.asp"><img src="images/btn_managestoreroles.jpg" alt="Manage Store Roles" width="75" height="75" border="0" /></a>
							<a href="adminscreens/managecoupons.asp"><img src="images/btn_managecoupons.jpg" alt="Manage Coupons" width="75" height="75"  border="0" /></a>
							<a href="menumaintenance/default.asp"><img src="images/btn_managemenu.jpg" alt="Manage Menu" width="75" height="75" border="0" /></a>
							<a href="adminscreens/manageaddresses.asp"><img src="images/btn_manageaddresses.jpg" alt="Manage Addresses" width="75" height="75" border="0" /></a>
							<a href="inventorymaster/"><img src="images/btn_masterinventory.jpg" alt="Master Inventory" width="75" height="75" border="0" ></a>
							<a href="adminscreens/changestore.asp"><img src="images/btn_changestore.jpg" alt="Change Store" width="75" height="75" border="0" ></a>
							<a href="adminscreens/manageaccounts.asp"><img src="images/btn_manageaccounts.jpg" alt="Manage Accounts" width="75" height="75" border="0" ></a>
							<a href="rpt_Payroll.asp"><img src="images/btn_ViewPayrollReports.jpg" alt="Payroll" width="75" height="75" border="0"></a>
						 <a href="http://extraordinary.vitos.com/vowners.asp" target="_new"><img src="images/btn_extraordinaryaccountinfo.jpg" alt="Owners Portal Account Info" width="75" height="75" border="0"></a>
							<a href="adminscreens/managemailers.asp"><img src="images/btn_managemailers.jpg" alt="Manage Mailers" width="75" height="75" border="0"></a>
							<a href="adminscreens/managescroller.asp"><img src="images/btn_managescroller.jpg" alt="Manage Scroller" width="75" height="75" border="0"></a>
							<a href="sw/"><img src="images/btn_downloadsoftware.jpg" alt="Download Software" width="75" height="75" border="0"></a>
				  <a href="scheduling/scheduling/scheduling.aspx?EmployeeID=<%=Session("EmployeeID")%>&StoreID=<%=Session("StoreID")%>"><img src="images/btn_scheduling.jpg" alt="Scheduling" width="75" height="75" border="0" /></a>
						  <a href="manageemployees.asp"><img src="images/btn_manageemployees.jpg" alt="Manage Employees" width="75" height="75" border="0" /></a>
						  </td>
                        </tr>
                        <tr>
                          <td bgcolor="#af0808" height="1"></td>
                        </tr>
<%
End If 'In office OR used Swipe card
%>

<%End If%>
                        <tr>
                          <td><div align="center"><a href="help.asp?HelpPage=<%=Replace(Request.ServerVariables("script_name"),"/","")%>"><img src="images/btn_help.jpg" alt="Help" width="75" height="75" border="0" /></a><a href="default.asp"><img src="images/btn_signoff.jpg" alt="Sign Off" width="75" height="75" border="0" /></a></div></td>
                        </tr>
                    </table></td>
                    <td width="300" align="center" valign="top">
<%'---------------------------------------- ISSUES BOX -------------------------------------------------------%>
<%
Dim sqlIssues, RSIssues
sqlIssues = "SELECT * FROM tblIssues WHERE (StoreID = "&Session("StoreID")&") AND (RADRAT > CONVERT(datetime, CONVERT(varchar, GETDATE() - 14, 101))) Order by RADRAT Desc;"
Set RSIssues=Conn.Execute(sqlIssues)
If RSIssues.BOF And RSIssues.EOF Then
Else
Dim iCount
%>
					<table cellspacing="0" cellpadding="1" width="300" bgcolor="#006d31" border="0" align="center">
                      <tbody>
                        <tr>
                          <td><table width="100%" border="0" cellpadding="1" cellspacing="3" bgcolor="fbf3c5">
                              <tr>
                                <td align="left" bgcolor="#006D31"><b><font color="#ffffff"><a href="issues.asp"><font color="#ffffff">Issues</font></a></b></span></td>
                              </tr>
                              <tr>
                                <td><marquee>
								<%iCount = 1
								Do While Not RSIssues.EOF%>
								&nbsp;<%=iCount%>) <%=Replace(RSIssues("Issue"),"*Edited*<br>","")%>
								<%RSIssues.MoveNext
								iCount = iCount + 1
								Loop%></marquee>
								</td>
                              </tr>
                          </table>
						  </td>
                        </tr>
                      </tbody>
                    </table>
<%End If%>
<%'------------------------------------ END ISSUES BOX ---------------------------------------------------%>

<%If Session("SecurityID") > 1 Then%>

<%
Dim sqlHoldOrders, RSHoldOrders

sqlHoldOrders = "SELECT OrderID, SessionID, IPAddress, EmpID, RefID, TransactionDate, SubmitDate, ReleaseDate, StoreID, CustomerID, CustomerName, CustomerPhone, AddressID, OrderTypeID, IsPaid, PaymentTypeID, PaymentReference, PaidDate, DeliveryCharge, DriverMoney, (Tax + Tax2) as Tax, Tip, OrderStatusID, OrderNotes, RADRAT FROM tblOrders WHERE (StoreID = "&Session("StoreID")&") AND OrderStatusID = 2 AND ReleaseDate <= dateadd(day, 7, getdate()) ORDER BY ReleaseDate"

Set RSHoldOrders = Conn.Execute(sqlHoldOrders)

If RSHoldOrders.BOF And RSHoldOrders.EOF Then 'None found
Else
%>
<%'------------------------------------ ON HOLD ORDERS BOX ---------------------------------------------------%>
					<table cellspacing="0" cellpadding="1" width="300" bgcolor="#c30a0d" border="0" align="center">
                      <tbody>
                        <tr>
                          <td><table width="100%" border="0" cellpadding="1" cellspacing="3" bgcolor="fbf3c5">
                              <tr>
                                <td align="left" bgcolor="#c30a0d"><b><font color="#ffffff">Hold Orders</b></span></td>
                              </tr>
<%Do While Not RSHoldOrders.EOF%>
                              <tr>
                                <td><input type="button" value="<%=RSHoldOrders("OrderID")%>" class="redbuttonthinwide"  onclick="document.location = 'ticket.asp?OrderID=<%=RSHoldOrders("OrderID")%>';"> <%=RSHoldOrders("ReleaseDate")%>
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

                      <table cellspacing="0" cellpadding="1" width="300" bgcolor="#006d31" border="0" align="center">
                        <tbody>
                          <tr>
                            <td>
							
							<table width="100%" border="0" cellpadding="1" cellspacing="3" bgcolor="fbf3c5">
                                <tr>
                                  <td align="left" bgcolor="#006D31"><b><font color="#ffffff">Information</b></span></td>
                                </tr>
                                <tr>
                                  <td align="left"><b>Driver Status - <%=RSEmp("DriverStatus")%></b></td>
                                </tr>
<%If Not RSPunch.EOF Then
If session("intAge") < 18 And Session("CurrentShiftID") <>"" Then 'reminder for break time if under the age of 18%>
                                <tr>
                                  <td align="left"><b><font color=red>You need to take a break by: <%=FormatDateTime(DateAdd("h",5,rspunch("RADRAT")),3)%></font></b></td>
                                </tr>
<%End If
End If
%>

<% If Session("SecurityID") > 1 Then 'need to loop through those that need break reminders

Dim sqlBreakReminder, RSBreakReminder
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

                        <table cellspacing="0" cellpadding="1" width="300" bgcolor="#006d31" border="0" align="center">
                          <tbody>
                            <tr>
                              <td>
                                <table width="100%" border="0" cellpadding="1" cellspacing="3" bgcolor="fbf3c5">
                                  <tr>
                                    <td align="left" bgcolor="#006D31"><font color="#ffffff"><b>Messages</b></font></td>
                                  </tr>

<%
Do While Not RSMessages.EOF%>
<%
If RSMessages("Status") = "Unread" Or RSMessages("Status") = "Replied" Then 
	Response.Redirect("readmessage.asp?MessageID="&RSMessages("MessageID"))
End If
%>

                                  <tr>
                                    <td align="left"><span class="smallblack10">

									<img src="images/icon_<%=RSMessages("Status")%>.gif" alt="Icon" width="16" height="15" hspace="3" border="0" align="absmiddle" />

									<a href="readmessage.asp?MessageID=<%=RSMessages("MessageID")%>" class="smallblack11">Message from <%=RSMessages("SenderName")%></a>, <%=FormatDateTime(RSMessages("RADRAT"),2)%></span></td>
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
                                            <td width="33%" class="style3"><div align="center"><img src="images/icon_unread.gif" alt="Unread Icon" width="16" height="15" hspace="3" vspace="3" align="absmiddle" />Unread</div></td>
                                            <td width="33%" class="style3"><div align="center"><img src="images/icon_read.gif" alt="Read Icon" width="16" height="15" hspace="3" vspace="3" align="absmiddle" />Read</div></td>
                                            <td width="33%" class="style3"><div align="center"><img src="images/icon_replied.gif" alt="Replied Icon" width="16" height="15" hspace="3" vspace="3" align="absmiddle" />Replied</div></td>
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

<%'------------------------------------ TESTS BOX ---------------------------------------------------%>

                      <iframe src="Litmos/LitmosIFrame.aspx?EmpID=<%=Session("EmpID")%>&StoreID=<%=Session("StoreID")%>" id="litmosiframe" name="litmosiframe" width="300" height="440" scrolling="no" onload="iframeclick()"/>
<%
'Add any new boxes here before the next TD
%>

</td>
                  </tr>
                  
              </table></td>
            </tr>
          </table></td>
        </tr>
      </table></td>
    </tr>
  </tbody>
</table>

</body>
</html>
<%
CloseConn 'Close Database Connection
%>
<!-- #Include Virtual="include2/db-disconnect.asp" -->
