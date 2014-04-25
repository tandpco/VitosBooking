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


Dim SQL

If Request("MarkAsRead") = "True" And Request.Querystring("MessageID") <> "" Then 'Mark As Read

		SQL = "UPDATE tblMessages SET Status= 'Read' WHERE (MessageID = " & Request.Querystring("MessageID") & ")"
		Conn.Execute(SQL)

ElseIf Request("MarkAsRead") = "False" And Request.Querystring("MessageID") <> "" Then 'Mark As Read

		SQL = "UPDATE tblMessages SET Status= 'Unread' WHERE (MessageID = " & Request.Querystring("MessageID") & ")"
		Conn.Execute(SQL)

ElseIf Request.QueryString("DelMessage")="True" And Request.Querystring("MessageID") <> "" Then

		SQL = "UPDATE tblMessages SET Status= 'Deleted' WHERE (MessageID = " & Request.Querystring("MessageID") & ")"

		Conn.Execute(SQL)

End If

Response.Redirect("main.asp")


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Vitos - Process Message</title>
<LINK href="css/vitosmain.css" type=text/css rel=stylesheet>
<script src="include2/redirect.js" type="text/javascript"></script>
</head>

<body>
<table cellspacing="0" cellpadding="1" width="1010" bgcolor="#006d31" border="0" align="center">
  <tbody>
    <tr>
      <td><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#fbf3c5" style="WIDTH: 100%">
        <tr>
          <td valign="top"><table width="100%" border="0" cellspacing="10" cellpadding="0">
            <tr>
              <td width="50%" align="center"><div align="left"><img src="images/logo_sm.jpg" width="200" height="96"/></div></td>
              <td width="50%" align="center">&nbsp;</td>
            </tr>
            <tr>
              <td colspan="2"><span class="PageHeader">PROCESS MESSAGE</span>
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

