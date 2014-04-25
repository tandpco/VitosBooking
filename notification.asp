<%
Option Explicit
Response.buffer = True
%>

<!-- #INCLUDE Virtual="/include2/adovbs.asp" -->
<!-- #INCLUDE Virtual="/include2/utility.asp" -->
<%

'-----------------------------------------Security Check (XXXX Screen)-------------------------------------------
'If  Session("SecurityID") < 1 Then Response.Redirect("default.asp")
'------------------------------------------------------------------------------------------------------------------------

Dim msg

msg = Request.QueryString("msg")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Vitos - Notification</title>
<LINK href="css/vitosmain.css" type=text/css rel=stylesheet>
<script src="include2/redirect.js" type="text/javascript"></script>
<!-- #Include File="include2/clock-server.asp" -->
</head>
<body onLoad="clockInit(clockLocalStartTime, clockServerStartTime); clockOnLoad(); redirect('default.asp')" onUnload="clockOnUnload()">
<table cellspacing="0" cellpadding="1" width="1010" bgcolor="#006d31" border="0" align="center">
  <tbody>
    <tr>
      <td><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#fbf3c5" style="WIDTH: 100%">
        <tr>
          <td valign="top"><table width="100%" border="0" cellspacing="10" cellpadding="0">
            <tr>
              <td width="50%" align="center"><div align="left"><img src="images/logo_sm.jpg" width="200" height="96"/></div></td>
              <td width="50%" align="center"><div align="right"><%=Session("name")%><br>
Store: <%=Session("StoreID")%><br /><div id="ClockDate"><%=clockDateString(gDate)%></div>
                      <div id="ClockTime" onClick="clockToggleSeconds()"><%=clockTimeString(Hour(gDate), Minute(gDate), Second(gDate))%></div>
                      <div class="counter" id="redirect">900</div></div></td>
            </tr>
            <tr>
              <td colspan="2">
			<form id="form1" name="form1" method="post" action="<%=Request.QueryString("NextPage")%>">
			  <p align="center" class="largetitle"><%=msg%></p> <div align="center">

<%
If Request.QueryString("NextPage") = "Back2" Then
%>
<input type="button" name="Try Again!" value="Try Again!" class="orangebuttonwide"  onclick="javascript:window.history.back(2)"/>
<%Else%>

			 
			    <INPUT TYPE="submit" Name="OK" class="orangebuttonwide"  Value="OK">

<%End If%>
			 			    </div>
							</FORM> 
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
  </tbody>
</table>
</body>
</html>
<%
'CloseConn 'Close Database Connection%>

