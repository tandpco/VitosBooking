<%
Option Explicit
Response.buffer = True
%>
<!-- #INCLUDE FILE="include2/adovbs.asp" -->
<!-- #INCLUDE FILE="include2/utility.asp" -->
<%
'-----------------------------------Security Check (Gen Employee Screen)----------------------------------------
If  Session("SecurityID") < 1 Then Response.Redirect("default.asp")
'------------------------------------------------------------------------------------------------------------------------

OpenSQLConn 'Open Test SQL Database Connection
Dim NotificationMsg

If Request.Form("Submit")  <> "" Then

	If request.Form("Submit") = "Open" Then 'open
	Dim sqlOpen
		sqlOpen = "Insert Into tblStoreReportDate(StoreID, ReportDate, CurrentStatus) Values("&Session("StoreID")&", '"&Request.Form("ReportDate")&"','Open')"
		Conn.execute(SQLOpen)
		Response.Redirect("main.asp")
	End If
End If 'Form Submitted

Dim Action, i

If Request.QueryString("Action") = "Open" Then
	Action = "Open"
ElseIf Request.QueryString("Action") = "Close" Then
	Action="Close"
	'Get current Open Date
	Dim SQLCurrentDate, RSCurrentDate
'	SQLCurrentDate="Select * from tblStoreReportDate where StoreID = "&Session("StoreID")
	sqlCurrentDate = "SELECT TOP (1) StoreReportID, StoreID, ReportDate, CurrentStatus, RADRAT FROM tblStoreReportDate WHERE (StoreID = "&Session("StoreID")&") AND (CurrentStatus = 'Open') ORDER BY ReportDate DESC;"
	Set RSCurrentDate=Conn.Execute(SQLCurrentDate)
Else
	Response.Redirect("default.asp")
End If

If Session("TransactionDate")="" Then
	Session("TransactionDate") =  FormatDateTime(Now(),2)
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Vitos - Open/Close Store
</title>
<LINK href="css/vitosmain.css" type=text/css rel=stylesheet>
<script src="include2/redirect.js" type="text/javascript"></script>

<!-- #Include File="include2/clock-server.asp" -->
</head>
<body onLoad="clockInit(clockLocalStartTime, clockServerStartTime); clockOnLoad(); redirect('default.asp')" onUnload="clockOnUnload()">
<table cellspacing="0" cellpadding="1" width="1010" bgcolor="#006d31" border="0" align="center">
    <tr>
      <td>
	  
	  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#fbf3c5" style="WIDTH: 100%">
        <tr>
          <td valign="top">
		  
		  
		  <table width="100%" border="0" cellspacing="10" cellpadding="0">
            <tr>
              <td width="50%" align="center"><div align="left"><img src="images/logo_sm.jpg" width="200" height="96"/></div></td>
              <td width="50%" align="center"><div align="right"><%=Session("name")%><br>
Store: <%=Session("StoreID")%><br /><div id="ClockDate"><%=clockDateString(gDate)%></div>
                      <div id="ClockTime" onClick="clockToggleSeconds()"><%=clockTimeString(Hour(gDate), Minute(gDate), Second(gDate))%></div>
                      <div class="counter" id="redirect">900</div></div></td>
            </tr>
			</table>

<form id="form" name="form" method="post" action="store.asp">
<%
'---------------------------- PROCESS PAYROLL BUTTON/INFORMATION WAS HERE
%>
<br>
&nbsp;<span class="PageHeader"><%=Action%> Store</span>


<%'-------------------------------------USER WANTS TO OPEN THE STORE
If Action = "Open" Then%>
<table width="100%" border="0" cellpadding="0" cellspacing="5">
  <tr><td>
				  <input type="image" name="Submit" src="images/btn_openstore.jpg" alt="Open" value="Open" width="75" height="75" border="0">
				  <input type=hidden name="ReportDate" value="<%=FormatDateTime(Now(),2)%>"><br />
                  Report Date: <%=FormatDateTime(Now(),2)%></td></tr>
<%If Session("SecurityID") > 3 Then%>
                        <tr>
                          <td bgcolor="#af0808" height="1"></td>
                        </tr>
						<tr>
                          <td align="left"><span class="PageHeader">Reports</span><br>
						<a href="rpt_SalesRevenue.asp"><img src="images/btn_salesrevenuereport.jpg" alt="Sales Reports" width="75" height="75" border="0" /></a>						
						<a href="dailystatistics.asp"><img src="images/btn_dailystatistics.jpg" alt="Daily Statistics" width="75" height="75" border="0" /></a>
					</td>
						</tr>
<%
If (Left(Request.ServerVariables("remote_addr"), 9) = "10.0.254." Or Left(Request.ServerVariables("remote_addr"), 7) = "10.0.0." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.1." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.2.") OR Session("swipe") Then
%>
                        <tr>
                          <td bgcolor="#af0808" height="1"></td>
                        </tr>
                        <tr>
                          <td align="left"><span class="PageHeader">Admin Functions</span><br>
						  <a href="adminscreens/managestores.asp"><img src="images/btn_managestores.jpg" alt="Manage Stores" width="75" height="75" border="0" /></a>
						  <a href="adminscreens/managesystemroles.asp"><img src="images/btn_managesystemroles.jpg" alt="Manage System Roles" width="75" height="75" border="0" /></a>
						  <a href="adminscreens/managestoreroles.asp"><img src="images/btn_managestoreroles.jpg" alt="Manage Store Roles" width="75" height="75" border="0" /></a>
<%'						  <a href="adminscreens/manageuniforms.asp"><img src="images/btn_manageuniforms.jpg" alt="Manage Uniforms" width="75" height="75" border="0" /></a>%>
							<a href="adminscreens/managecoupons.asp"><img src="images/btn_managecoupons.jpg" alt="Manage Coupons" width="75" height="75"  border="0" /></a>
						  <a href="testing/test-tests.asp"><img src="images/btn_managetests.jpg" alt="Manage Tests" width="75" height="75"  border="0" /></a>
							<a href="menumaintenance/default.asp"><img src="images/btn_managemenu.jpg" alt="Manage Menu" width="75" height="75" border="0" /></a>
							<a href="#"><img src="images/btn_manageaddresses.jpg" alt="Manage Addresses" width="75" height="75" border="0" /></a>
							<a href="inventorymaster/"><img src="images/btn_masterinventory.jpg" alt="Master Inventory" width="75" height="75" border="0" ></a>
							<a href="adminscreens/changestore.asp"><img src="images/btn_changestore.jpg" alt="Change Store" width="75" height="75" border="0" ></a>
						  <a href="adminscreens/manageaccounts.asp"><img src="images/btn_manageaccounts.jpg" alt="Manage Accounts" width="75" height="75" border="0" ></a>
						 <a href="rpt_Payroll.asp"><img src="images/btn_ViewPayrollReports.jpg" alt="Payroll" width="75" height="75" border="0"></a>
						 <a href="http://extraordinary.vitos.com/vowners.asp" target="_new"><img src="images/btn_extraordinaryaccountinfo.jpg" alt="Owners Portal Account Info" width="75" height="75" border="0"></a></td>
                        </tr>
                        <tr>
                          <td bgcolor="#af0808" height="1"></td>
                        </tr>
<%
End If 'In office OR used Swipe card
%>

		<%End If%>
<%
Else '--------------------------------USER WANTS TO CLOSE THE STORE
%>
<%
'--------------------------- CHECK TO SEE OPEN ORDERS

Dim sqlOpenOrders, RSOpenOrders, OrderStatus
sqlOpenOrders = "SELECT TOP (100) PERCENT dbo.tblOrders.OrderID, dbo.tblOrders.CustomerID, dbo.tblOrders.IsPaid, dbo.tblOrders.CustomerName, dbo.tblOrders.CustomerPhone, dbo.tblAddresses.AddressLine1, dbo.tblOrders.SubmitDate, dbo.tblOrders.OrderTypeID, dbo.tblOrders.StoreID, dbo.tblOrders.TransactionDate FROM dbo.tblOrders INNER JOIN dbo.tblAddresses ON dbo.tblOrders.AddressID = dbo.tblAddresses.AddressID WHERE (dbo.tblOrders.OrderStatusID > 2 and dbo.tblOrders.OrderStatusID < 10) AND (dbo.tblOrders.StoreID = "&Session("StoreID")&")"

Set RSOpenOrders = Conn.Execute(sqlOpenOrders)

If RSOpenOrders.BOF And RSOpenOrders.EOF Then 'No Open Orders, proceed
OrderStatus = "Clear"
%>

<div align="center">
  <table width="100%" border="0" cellspacing="3" cellpadding="0">
    <tr>
      <td width="275" valign="top"><b>Steps To Close</b><ol>
           <li>Close Open Orders <b>(done)</b></li>
           <li>Punch People Out</li>
           <li>Make a Deposit</li>
           <li>Settle Current CC Batch</li>
           <li>Print Daily Reports</li>
           <li>Punch People Out - Final</li>
           <li>Close Store</li>
           </ol></td>
      <td bgcolor="#C30A0D" width="1"></td>
      <td valign="top"><div align="center"><strong>There are no Incomplete Orders. On to Deposits...</strong><br>
          <br>
          <input type="button" value="Continue" class="redbuttonwide"  onClick="document.location = 'makedeposit.asp?Action=Close';">
      </div></td>
    </tr>
  </table>
</div>
<%
Else 'Open orders exist and need to be dealt with
%>
			<br><b>You cannot close a store while Open Orders exist. Please review the list below.</b><br><br>

  <table width="975" border="0" cellspacing="3" cellpadding="0">
										  <tr>
										  <td width="275" valign="top"><b>Steps To Close</b><ol>
											   <li>Close Open Orders <b>(here)</b></li>
										    <li>Punch People Out</li>
											   <li>Make a Deposit</li>
											   <li>Settle Current CC Batch</li>
											   <li>Print Daily Reports</li>
											   <li>Punch People Out - Final</li>
											   <li>Close Store</li>
											   </ol></td>
										  <td bgcolor="#C30A0D" width="1"></td>
										  <td valign="top">
  <table width="700" border="0" cellspacing="3" cellpadding="0">
<%Do While Not RSOpenOrders.EOF 'Loop through the recordset%>
										  <tr>
											<td width="100" valign="top"><input type="button" value="<%=RSOpenOrders("OrderID")%>" class="orangebuttonwide"  onclick="document.location = 'ticket.asp?OrderID=<%=RSOpenOrders("OrderID")%>';"></td>
											<td valign="top">
											<b>Order Taken: <%=FormatDateTime(RSOpenOrders("SubmitDate"),1)%></b><br> 
											<%If Len(RSOpenOrders("CustomerName"))>0 Then Response.Write RSOpenOrders("CustomerName") & ", "%>
											<%If Len(RSOpenOrders("CustomerPhone"))>0 Then Response.Write MkPhoneNum(RSOpenOrders("CustomerPhone")) & ", "%>
											<%=RSOpenOrders("AddressLine1")%><br><b>Paid Status:  <%=RSOpenOrders("IsPaid")%></b>
											</td>
<%If RSOpenOrders("IsPaid")="True" Then%>
											<td colspan=2 valign="top"><input type="button" value="Complete" class="redbuttonwide"  onclick="document.location = 'updateorderstatus.asp?OrderID=<%=RSOpenOrders("OrderID")%>&OrderStatusID=10';"></td>
<%Else%>
											<td width="75" valign="top"><input type="button" value="Mark Paid" class="redbuttonwide"  onclick="document.location = 'updateorderstatus.asp?OrderID=<%=RSOpenOrders("OrderID")%>&IsPaid=True&OrderStatusID=10';"></td>
											<td width="75" valign="top"><input type="button" value="Mark UnPaid" class="redbuttonwide"  onclick="document.location = 'updateorderstatus.asp?OrderID=<%=RSOpenOrders("OrderID")%>&OrderStatusID=12';"></td>
<%End If%>
										  </tr>
						<%
						RSOpenOrders.MoveNext
						Loop
						%>
			</table>
</tr></table>
<%End If 'End If Open Orders %>

<% End If 'End Deal with Open Orders%>
</form>
</td>
              </tr>
            <tr>
              <td colspan="2"><div align="center">
                          <a href="main.asp"><img src="images/btn_mainmenu.jpg" alt="Main Menu" width="75" height="75" border="0" /></a><a href="help.asp?HelpPage=<%=Replace(Request.ServerVariables("script_name"),"/","")%>"><img src="images/btn_help.jpg" alt="Help" width="75" height="75" border="0" /></a><a href="default.asp"><img src="images/btn_signoff.jpg" alt="Sign Off" width="75" height="75" border="0" /></a>
              </div></td>
              </tr>
          </table></td>
        </tr>
      </table></td>
    </tr>
</table>
</body>
</html>
<%
CloseConn 'Close Database Connection%>

