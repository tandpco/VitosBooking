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


Dim sql, RS

SQL="SELECT * FROM tblMessages WHERE (((tblMessages.RecipientID)="&Session("EmpID")&") AND ((tblMessages.MessageID)="&Request.QueryString("MessageID")&"));"

Set RS=Conn.Execute(SQL)

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Vitos - Read Message</title>
<LINK href="css/vitosmain.css" type=text/css rel=stylesheet>
<script src="include2/redirect.js" type="text/javascript"></script>
<!-- #Include File="include2/clock-server.asp" -->
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
              <td width="50%" align="center"><div align="right"><%=Session("name")%><br>
Store: <%=Session("StoreID")%><br /><div id="ClockDate"><%=clockDateString(gDate)%></div>
                      <div id="ClockTime" onclick="clockToggleSeconds()"><%=clockTimeString(Hour(gDate), Minute(gDate), Second(gDate))%></div>
                      <div class="counter" id="redirect">900</div></div></td>
            </tr>
            <tr>
              <td colspan="2"><span class="PageHeader">READ MESSAGE</span>
                <table cellspacing="0" cellpadding="1" width="100%" bgcolor="#006d31" border="0" align="center">
                  <tbody>
                    <tr>
                      <td><table width="100%" border="0" cellpadding="3" cellspacing="3" bgcolor="fbf3c5">
                          <tr>
                            <td bgcolor="#006D31"><strong><font color="#ffffff">Message</font></strong></td>
                            <td width="100" bgcolor="#006D31" align="center"><strong><font color="#ffffff">Actions</font></strong></td>
                          </tr>
                         
                          <tr>
                            <td valign="top">You have received a message from <%=RS("SenderName")%>; <br><br><font size=+1><%=(RS("Message"))%></font><hr></td>

                            <td width="90" valign="top"><div align="center">

<%If RS("Status")="Unread" Then%>							
							<a href="processmessage.asp?MarkAsRead=True&amp;MessageID=<%=RS("MessageID")%>"><img src="images/btn_msgmarkasread.jpg" alt="Mark As Read" width="75" height="75" border="0" /></a><br />
<%Else%>
							<a href="processmessage.asp?MarkAsRead=False&amp;MessageID=<%=RS("MessageID")%>"><img src="images/btn_msgmarkasunread.jpg" alt="Mark As Unread" width="75" height="75" border="0" /></a><br />

<%End If%>
                                <a href="replytomessage.asp?MessageID=<%=RS("MessageID")%>"><img src="images/btn_msgreplytomessage.jpg" alt="Reply To Message" width="75" height="75" border="0" /></a><br />
                                <a href="processmessage.asp?DelMessage=True&amp;MessageID=<%=RS("MessageID")%>"><img src="images/btn_msgdeletemessage.jpg" alt="Delete Message" width="75" height="75" border="0" /></a></div></td>
                          </tr>
                      </table></td>
                    </tr>
                  </tbody>
                </table>
                <p>&nbsp;</p></td>
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
CloseConn 'Close Database Connection%>

