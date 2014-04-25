<%
Option Explicit
Response.buffer = True
%>
<!-- #Include Virtual="/include2/globals.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Vitos - Error</title>
<link href="/css/vitosmain.css" type=text/css rel=stylesheet />
<script src="/include2/redirect.js" type="text/javascript"></script>
<!-- #Include Virtual="/include2/clock-server.asp" -->
</head>
<body onload="clockInit(clockLocalStartTime, clockServerStartTime); clockOnLoad(); redirect('default.asp')" onunload="clockOnUnload()">
<table cellspacing="0" cellpadding="1" width="1010" bgcolor="#006d31" border="0" align="center">
  <tbody>
    <tr>
      <td><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#fbf3c5" style="WIDTH: 100%">
        <tr>
          <td valign="top"><table width="100%" border="0" cellspacing="10" cellpadding="0">
            <tr>
              <td width="50%" align="center"><div align="left"><img src="images/logo_sm.jpg" width="200" height="96"/></div></td>
              <td width="50%" align="center"><div align="right">
<%
If gbTestMode Then
	If gbDevMode Then
%>
				<strong>DEV SYSTEM</strong><br/>
<%
	Else
%>
				<strong>TEST SYSTEM</strong><br/>
<%
	End If
End If
%>
				<b><%=Session("name")%></b>
				<div id="ClockDate"><%=clockDateString(gDate)%></div>
				<div id="ClockTime" onclick="clockToggleSeconds()"><%=clockTimeString(Hour(gDate), Minute(gDate), Second(gDate))%></div>
				<div class="counter" id="redirect">900</div>
              </div></td>
            </tr>
<%
If Request("err").Count > 0 Then
%>
            <tr>
              <td colspan="2"><div align="center"><span class="PageHeader">We are sorry 
				  but an unexpected error has occured. 
<%
	If Len(Request("err")) > 0 Then
%>
				  The error was: <br />
                <%=Request("err")%><br />
<%
	End If
	
	If Request("o").Count > 0 Then
		If IsNumeric(Request("o")) Then
			Session("OrderID") = 0
%>
				<a href="ordering/unitselect.asp?o=<%=Request("o")%>"><img src="/images/btn_returntoorder.jpg" alt="Return To Order" border="0" /></a><br />
<%
		End If
	End If
%>
                <br />Please call 419.720.5050 for further assistance.  </tr>
<%
Else
%>
            <tr>
              <td colspan="2"><div align="center"><span class="PageHeader">We are sorry but you have reached this page in error. <br />
                Most likely because you are not accessing this system from an authorized location. <br />
                Please call 419.720.5050 if you feel you have reached this page in error.</span></div></td>
              </tr>
<%
End If
%>
            <tr>
              <td colspan="2"><div align="center">
<%
If Session("SecurityID") <> "" Then
%>
              <a href="/main.asp"><img src="/images/btn_mainmenu.jpg" alt="Main Menu" border="0" /></a>
<%
End If
%>
			  <a href="/default.asp"><img src="/images/btn_signoff.jpg" alt="Sign Off" border="0" /></a>
              <br /><span class="small">For technical assistance, please call 419.720.5050</span>
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


