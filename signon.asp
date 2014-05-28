<%
Option Explicit
Response.buffer = True
%>
<!-- #Include Virtual="include2/globals.asp" -->
<!-- #INCLUDE FILE="include2/adovbs.asp" -->
<!-- #INCLUDE FILE="include2/utility.asp" -->
<%

OpenSQLConn

Dim SQL, RS

If Left(Request.ServerVariables("remote_addr"), 9) = "10.0.254." Or Left(Request.ServerVariables("remote_addr"), 7) = "10.0.0." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.1." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.2." Then
	SQL = "Select * from tblStores where storeid = 1"
Else
	If Left(Request.ServerVariables("remote_addr"), 7) = "10.0.1." Or Left(Request.ServerVariables("remote_addr"), 7) = "10.0.2." Then
		SQL = "Select * from tblStores where storeid = 9"
	Else
		SQL = "Select * from tblStores where networkip = '" & Left(Request.ServerVariables("remote_addr"), InStrRev(Request.ServerVariables("remote_addr"), ".")) & "0'"
	End If
End If

Set RS=Conn.Execute(SQL)

If RS.EOF And RS.BOF Then
Response.Redirect("Error.asp")
Else

Session("StoreID") = RS("StoreID")
Session("StorePostalCode") = Trim(RS("PostalCode"))
Session("TZOffset") = RS("TZOffset")

End If

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Vito's Management System</title>
<LINK href="css/vitosmain.css" type=text/css rel=stylesheet>
<script src="include2/redirect2.js" type="text/javascript"></script>
<!-- #Include File="include2/clock-server.asp" -->

<script LANGUAGE="JavaScript">
<!--
var ie4=document.all;

function resetRedirect() {
	var loRedirectDiv;
	
	loRedirectDiv = ie4? eval("document.all.redirect") : document.getElementById("redirect");
	loRedirectDiv.innerHTML = <%=gnRedirectTime%>;
}

function disableEnterKey() {
	var loText, loDiv;
	
	if (event.keyCode == 13) {
		event.cancelBubble = true;
		event.returnValue = false;
		return false;
	}
}

function backspace() {
	var loText, lsText;
	
	loText = ie4? eval("document.all.EmployeeID") : document.getElementById('EmployeeID');
	lsText = loText.value;
	if (lsText.length > 0) {
		lsText = lsText.substr(0, (lsText.length - 1));
		loText.value = lsText;
	}
}


function ValidateForm(theForm)
	{

if (theForm.EmployeeID.value == "")
{
alert("Please enter your Employee ID.");
theForm.EmployeeID.focus();
return (false);
}

 return (true);
	}
//-->
</script>
</head>
<body onLoad="document.form1.EmployeeID.focus(); clockInit(clockLocalStartTime, clockServerStartTime); clockOnLoad(); redirect('/default.asp')" onUnload="clockOnUnload()">


<table cellspacing="0" cellpadding="1" width="1010" bgcolor="#006d31" border="0" align="center">
<tbody>
            <tr>
              <td>
			  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#fbf3c5" style="WIDTH: 100%">
                  <tr>
                    <td valign="top">
					<table width="100%" border="0" cellspacing="10" cellpadding="0">
                    <tr>
                          <td align=center>
<form name="form1" method="post" action="main.asp" onSubmit="return ValidateForm(this);">
  <img src="images/logo.jpg" alt="Vito's Logo" width="400" height="191" border="0"><br />
  <table width="600" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%" valign="top"><table width="100%" border="0" align="center" cellpadding="0">
        <tr>
          <td><div align="center" class="PageHeader">Employee ID</div></td>
        </tr>
        <%If request.QueryString("ID")="No" Then%>
        <tr>
          <td><div align="center" class="PageHeader">That UserID Was Not Found.<br /> 
            Please Try Again.</div></td>
        </tr>
        <%End If%>
        <tr>
          <td align="center"><input type="password"  name="EmployeeID" id="EmployeeID"  autofocus autocomplete="off" onkeydown="//disableEnterKey();"/></td>
        </tr>
        <tr>
          <td align="center"><table border="0" cellspacing="3" cellpadding="0">
              <tr>
                <td><div align="center"><INPUT TYPE="button" NAME="one" VALUE=" 1 " class="orangebutton75" OnClick="form1.EmployeeID.value += '1'; resetRedirect();"></div></td>
                <td><div align="center"><INPUT TYPE="button" NAME="two" VALUE=" 2 " class="orangebutton75" OnClick="form1.EmployeeID.value += '2'; resetRedirect();"></div></td>
                <td><div align="center"><INPUT TYPE="button" NAME="three" VALUE=" 3 " class="orangebutton75" OnClick="form1.EmployeeID.value += '3'; resetRedirect();"></div></td>
              </tr>
              <tr>
                <td><div align="center"><INPUT TYPE="button" NAME="four" VALUE=" 4 " class="orangebutton75" OnClick="form1.EmployeeID.value += '4'; resetRedirect();"></div></td>
                <td><div align="center"><INPUT TYPE="button" NAME="five" VALUE=" 5 " class="orangebutton75" OnClick="form1.EmployeeID.value += '5'; resetRedirect();"></div></td>
                <td><div align="center"><INPUT TYPE="button" NAME="six" VALUE=" 6 " class="orangebutton75" OnClick="form1.EmployeeID.value += '6'; resetRedirect();"></div></td>
              </tr>
              <tr>
                <td><div align="center"><INPUT TYPE="button" NAME="seven" VALUE=" 7 " class="orangebutton75" OnClick="form1.EmployeeID.value += '7'; resetRedirect();"></div></td>
                <td><div align="center"><INPUT TYPE="button" NAME="eight" VALUE=" 8 " class="orangebutton75" OnClick="form1.EmployeeID.value += '8'; resetRedirect();"></div></td>
                <td><div align="center"><INPUT TYPE="button" NAME="nine" VALUE=" 9 " class="orangebutton75" OnClick="form1.EmployeeID.value += '9'; resetRedirect();"></div></td>
              </tr>
              <tr>
                <td><div align="center"><INPUT TYPE="button" NAME="Bksp" VALUE="Bksp" class="orangebutton75" OnClick="backspace(); resetRedirect();"></div></td>
                <td><div align="center"><INPUT TYPE="button" NAME="zero" VALUE=" 0 " class="orangebutton75" OnClick="form1.EmployeeID.value += '0'; resetRedirect();"></div></td>
                <td><div align="center"><input type="submit" name="Submit" class="orangebutton75" id="Submit" value="OK" /></div></td>
                </div></td>
              </tr>
          </table></td>
        </tr>
        
        
      </table>
        </td>
      <td width="50%" align="center" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><div align="center"><span class="PageHeader">Instructions</span></div></td>
        </tr>
        <tr>
          <td><div align="left">Enter your Employee ID in the box provided using the keypad, a keyboard, or your swipecard.</div></td>
        </tr>
        <tr>
          <td bgcolor="#B30D11" height="1"></td>
        </tr>
        <tr>
          <td align="center"><p><a href="driverdispatch.asp"><img src="images/btn_driverdispatch.jpg" alt="Driver Dispatch" width="75" height="75" border="0" align="absmiddle" /></a><br>
            For technical assistance, please call 419.720.5050<br><br>
            <strong>
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
            </strong><br>
  Copyright &copy; 2011-2013
  <div class="counter" id="redirect"><%=gnRedirectTime%></div>
<a href="default.asp"><img src="images/btn_signoff.jpg" alt="Sign Off" width="75" height="75" border="0" align="absmiddle" /></a>
  </p></td>
        </tr>
      </table>
      </td>
    </tr>
  </table>
  </form>
</td>
                        </tr>
                      </table> </td>
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
