<%
Option Explicit

Response.buffer = TRUE
Dim gnCustomerID

If Session("SecurityID") = "" Then
  Response.Redirect("/default.asp")
End If

If Request("p").Count = 0 Then
  Response.Redirect("neworder.asp")
End If

If Not IsNumeric(Request("p")) Then
  Response.Redirect("neworder.asp")
End If

If Request("t").Count = 0 Then
  Response.Redirect("neworder.asp")
End If

If Not IsNumeric(Request("t")) Then
  Response.Redirect("neworder.asp")
End If

If Request("c").Count <> 0 Then
  If IsNumeric(Request("c")) Then
    gsAssignVisible = "hidden"
    gsPostalVisible = "visible"
        gnCustomerID = Request("c")
  Else
        gnCustomerID = 0
  End If
Else
    gnCustomerID = 0
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
Dim gbShowMenuButtons
Dim gnOrderTypeID, gsPhone, ganAddressIDs(), ganStoreIDs(), gasAddresses(), i, rowCnt
Dim ganOrderIDs(), gsLocalErrorMsg
Dim gsAssignVisible, gsPostalVisible, gsPhoneVisible
Dim gasPostalCodes(), ganAreaCodes()

If Session("OrderID") = 0 Then
  gbShowMenuButtons = TRUE
Else
  gbShowMenuButtons = FALSE
End If

gsAssignVisible = "visible"
gsPostalVisible = "hidden"
gsPhoneVisible = "hidden"
Dim gbNeedPrinterAlert

gbNeedPrinterAlert = FALSE

gnOrderTypeID = CLng(Request("t"))

gsPhone = Request("p")

If (Request("r")) Then
    rowCnt = Request("r")
Else
    rowCnt = 3
End If

Session("CustomerPhone") = gsPhone

Session("ReturnURL") = "/ordering/customerfind-new.asp?t=" & gnOrderTypeID & "&p=" & gsPhone
Session("SaveURL") = "/ordering/addressfind.asp?t=" & gnOrderTypeID & "&p=" & gsPhone

If GetAddressesByPhone(gsPhone, rowCnt, ganAddressIDs, ganStoreIDs, gasAddresses) Then
  If ganAddressIDs(0) = 0 Then
    gsAssignVisible = "hidden"
    gsPhoneVisible = "visible"
  End If
Else
  Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If Not GetStorePostalCodes(Session("StoreID"), gasPostalCodes) Then
  Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If Not GetStoreAreaCodes(Session("StoreID"), ganAreaCodes) Then
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
<!-- #Include Virtual="ordering/customerfind-js.asp" -->

</head>

<body onload="clockInit(clockLocalStartTime, clockServerStartTime); clockOnLoad();" onunload="clockOnUnload()">

<div id="mainwindow" style="position: absolute; top: 0px; left: 0px; width=810PX; height: 968px; overflow: hidden;">
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
                    <div align="center">
                        <br />
                        <ol id="tabs">
                <li><a onclick="back2Delivery();" title="Delivery">Delivery</a></li>
                <li><a onclick="back2Phone();" title="Phone">Phone</a></li>
                <li class="active">Address</li>
                <li>Customer Name</li>
                <li>Order</li>
                <li>Notes</li>
            </ol>           
                    </div>
        </td>
      </tr>
<%
'If ganCustomerIDs(0) <> 0 Then
' Dim adxCnt
'    adxCnt = UBound(ganCustomerIDs) + 1
'else
'    adxCnt = 1
'End If
'If gnOrderTypeID <> 1 Then
'    adxCnt = adxCnt + 1
'End If
%>
      <tr>
        <td valign="top" width="1010">
          <div id="content" align="center" style="position: relative; width: 1010PX;  overflow: auto;">
            <div id="assigndiv" align="center" style="position: relative; top: 0px; left: 0px; width: 810PX; visibility: <%=gsAssignVisible%>;">
<%
If gnOrderTypeID = 1 Then
%>
              <div align="center"><strong>SELECT DELIVERY LOCATION</strong></div><br/>
<%
Else
%>
              <div align="center"><strong>SELECT CUSTOMER FOR PICKUP</strong></div><br/>
<%
End If
%>
                            <div align="center"><strong>Most Recent Addresses</strong></div><br/>
<%

If ganAddressIDs(0) <> 0 Then

  For i = 0 To UBound(ganAddressIDs)
    
    If ganAddressIDs(i) <> 0 Then
        If gnOrderTypeID = 1 And ganStoreIDs(i) <> Session("StoreID") Then
          %><button style="width: 810px; " onclick="window.location='otherstore.asp?t=<%=gnOrderTypeID%>&s=<%=ganStoreIDs(i)%>&a=<%=ganAddressIDs(i)%>'"><%=gasAddresses(i)%></button><%
        Else
					%><button style="width: 810px; " onclick="window.location='customerselect-new.asp?t=<%=gnOrderTypeID%>&a=<%=ganAddressIDs(i)%>'"><%=gasAddresses(i)%></button><%
        End If
    End If
  Next
End If
%>
                        <br /><br />
                        <button style="width: 400PX;" onclick="window.location='customerfind-new.asp?t=<%=gnOrderTypeID %>&p=<%=gsPhone %>&r=99999'">All Addresses</button>
                        <button style="width: 400PX;" onclick="window.location='../custmaint/newaddress.asp?o=0&a=0'">Add New Address</button><br />


            <span class="orangetext">For technical assistance, please call 419.720.5050</span>
          </div>

            </div>
            <div id="postalcodediv" style="position: absolute; top: 0px; left: 0px; width: 810PX; visibility: <%=gsPostalVisible%>;">
              <table align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td colspan="3"><div align="center">
                    <strong>ENTER POSTAL CODE</strong></div></td>
                  <td width="25">&nbsp;</td>
                  <td colspan="3"><div align="center">
                    <strong>ENTER STREET #</strong></div></td>
                </tr>
                <tr>
                  <td colspan="3"><div align="center">
                    <input type="text" id="postalcode" style="width: 200px" autocomplete="off" onkeydown="disableEnterKey();" value="<%=Request("z")%>" /></div></td>
                  <td width="25">&nbsp;</td>
                  <td colspan="3"><div align="center">
                    <input type="text" id="streetnumber" style="width: 200px" autocomplete="off" onkeydown="disableEnterKey();" value="<%=Request("y")%>" /></div></td>
                </tr>
                <tr>
                  <td><button onclick="addToPostalCode('1')">1</button></td>
                  <td><button onclick="addToPostalCode('2')">2</button></td>
                  <td><button onclick="addToPostalCode('3')">3</button></td>
                  <td width="25">&nbsp;</td>
                  <td><button onclick="addToStreetNumber('1')">1</button></td>
                  <td><button onclick="addToStreetNumber('2')">2</button></td>
                  <td><button onclick="addToStreetNumber('3')">3</button></td>
                </tr>
                <tr>
                  <td><button onclick="addToPostalCode('4')">4</button></td>
                  <td><button onclick="addToPostalCode('5')">5</button></td>
                  <td><button onclick="addToPostalCode('6')">6</button></td>
                  <td width="25">&nbsp;</td>
                  <td><button onclick="addToStreetNumber('4')">4</button></td>
                  <td><button onclick="addToStreetNumber('5')">5</button></td>
                  <td><button onclick="addToStreetNumber('6')">6</button></td>
                </tr>
                <tr>
                  <td><button onclick="addToPostalCode('7')">7</button></td>
                  <td><button onclick="addToPostalCode('8')">8</button></td>
                  <td><button onclick="addToPostalCode('9')">9</button></td>
                  <td width="25">&nbsp;</td>
                  <td><button onclick="addToStreetNumber('7')">7</button></td>
                  <td><button onclick="addToStreetNumber('8')">8</button></td>
                  <td><button onclick="addToStreetNumber('9')">9</button></td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                  <td><button onclick="addToPostalCode('0')">0</button></td>
                  <td><button onclick="backspacePostalCode()">Bksp</button></td>
                  <td width="25">&nbsp;</td>
                  <td>&nbsp;</td>
                  <td><button onclick="addToStreetNumber('0')">0</button></td>
                  <td><button onclick="backspaceStreetNumber()">Bksp</button></td>
                </tr>
                <tr>
<%
Dim gnMaxI

If Len(gasPostalCodes(0)) > 0 Then
  If UBound(gasPostalCodes) > 2 Then
    gnMaxI = 2
  Else
    gnMaxI = UBound(gasPostalCodes)
  End If
  
  For i = 0 To gnMaxI
%>
                  <td><button onclick="setPostalCode('<%=gasPostalCodes(i)%>')"><%=gasPostalCodes(i)%></button></td>
<%
  Next
  
  If gnMaxI < UBound(gasPostalCodes) Then
    gnMaxI = gnMaxI + 1
    
    Response.Write("<td width=""25"">&nbsp;</td>")
    
    For i = gnMaxI To UBound(gasPostalCodes)
      If i > gnMaxI And i Mod 3 = 0 Then
        Response.Write("<td width=""25"">&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr><tr>")
      End If
%>
                  <td><button onclick="setPostalCode('<%=gasPostalCodes(i)%>')"><%=gasPostalCodes(i)%></button></td>
<%
    Next
  End If
  
  If (UBound(gasPostalCodes) + 1) Mod 3 > 0 Then
    For i = ((UBound(gasPostalCodes) + 1) Mod 3) To 2
%>
                  <td>&nbsp;</td>
<%
    Next
  End If
Else
%>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
<%
End If

If UBound(gasPostalCodes) < 3 Then
%>
                  <td width="25">&nbsp;</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
<%
End If
%>
                </tr>
                <tr>
                  <td colspan="7">&nbsp;</td>
                </tr>
                <tr>
                  <td colspan="7" align="center"><button onclick="cancelPostalCode();">Cancel</button>&nbsp;<button onclick="getStreetLetter(true);">1/2 Address</button>&nbsp;<button onclick="getStreetLetter(false);">Next</button></td>
                </tr>
              </table>
            </div>
            <div id="addressdiv" style="position: absolute; top: 0px; left: 0px; width: 810PX; visibility: hidden;">
              <table align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td colspan="11"><div align="center">
                    <span id="postalstreetspan"></span><strong>ENTER FIRST LETTER OF STREET NAME</strong></div></td>
                </tr>
                <tr>
                  <td><button onclick="goNext('1')">1</button></td>
                  <td><button onclick="goNext('2')">2</button></td>
                  <td><button onclick="goNext('3')">3</button></td>
                  <td><button onclick="goNext('4')">4</button></td>
                  <td><button onclick="goNext('5')">5</button></td>
                  <td><button onclick="goNext('6')">6</button></td>
                  <td><button onclick="goNext('7')">7</button></td>
                  <td><button onclick="goNext('8')">8</button></td>
                  <td><button onclick="goNext('9')">9</button></td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr>
                  <td><button onclick="goNext('Q')">Q</button></td>
                  <td><button onclick="goNext('W')">W</button></td>
                  <td><button onclick="goNext('E')">E</button></td>
                  <td><button onclick="goNext('R')">R</button></td>
                  <td><button onclick="goNext('T')">T</button></td>
                  <td><button onclick="goNext('Y')">Y</button></td>
                  <td><button onclick="goNext('U')">U</button></td>
                  <td><button onclick="goNext('I')">I</button></td>
                  <td><button onclick="goNext('O')">O</button></td>
                  <td><button onclick="goNext('P')">P</button></td>
                  <td>&nbsp;</td>
                </tr>
                <tr>
                  <td><button onclick="goNext('A')">A</button></td>
                  <td><button onclick="goNext('S')">S</button></td>
                  <td><button onclick="goNext('D')">D</button></td>
                  <td><button onclick="goNext('F')">F</button></td>
                  <td><button onclick="goNext('G')">G</button></td>
                  <td><button onclick="goNext('H')">H</button></td>
                  <td><button onclick="goNext('J')">J</button></td>
                  <td><button onclick="goNext('K')">K</button></td>
                  <td><button onclick="goNext('L')">L</button></td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr>
                  <td><button onclick="goNext('Z')">Z</button></td>
                  <td><button onclick="goNext('X')">X</button></td>
                  <td><button onclick="goNext('C')">C</button></td>
                  <td><button onclick="goNext('V')">V</button></td>
                  <td><button onclick="goNext('B')">B</button></td>
                  <td><button onclick="goNext('N')">N</button></td>
                  <td><button onclick="goNext('M')">M</button></td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td><button onclick="cancelAddress()">Back</button></td>
                </tr>
              </table>
            </div>
            <div id="namediv" style="position: absolute; top: 0px; left: 0px; width: 810PX; visibility: hidden;">
              <table align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td colspan="11"><div align="center">
                    <strong>ENTER CUSTOMER NAME</strong></div></td>
                </tr>
                <tr>
                  <td colspan="11"><div align="center">
                    <input type="text" id="name" style="width: 800px" autocomplete="off" onkeydown="disableEnterKey();" /></div></td>
                </tr>
                <tr>
                  <td><button onclick="addToName('1')">1</button></td>
                  <td><button onclick="addToName('2')">2</button></td>
                  <td><button onclick="addToName('3')">3</button></td>
                  <td><button onclick="addToName('4')">4</button></td>
                  <td><button onclick="addToName('5')">5</button></td>
                  <td><button onclick="addToName('6')">6</button></td>
                  <td><button onclick="addToName('7')">7</button></td>
                  <td><button onclick="addToName('8')">8</button></td>
                  <td><button onclick="addToName('9')">9</button></td>
                  <td><button onclick="addToName('0')">0</button></td>
                  <td><button onclick="backspaceName()">Bksp</button></td>
                </tr>
                <tr>
                  <td><button onclick="addToName('Q')">Q</button></td>
                  <td><button onclick="addToName('W')">W</button></td>
                  <td><button onclick="addToName('E')">E</button></td>
                  <td><button onclick="addToName('R')">R</button></td>
                  <td><button onclick="addToName('T')">T</button></td>
                  <td><button onclick="addToName('Y')">Y</button></td>
                  <td><button onclick="addToName('U')">U</button></td>
                  <td><button onclick="addToName('I')">I</button></td>
                  <td><button onclick="addToName('O')">O</button></td>
                  <td><button onclick="addToName('P')">P</button></td>
                  <td><button onclick="addToName('#')">#</button></td>
                </tr>
                <tr>
                  <td><button onclick="addToName('A')">A</button></td>
                  <td><button onclick="addToName('S')">S</button></td>
                  <td><button onclick="addToName('D')">D</button></td>
                  <td><button onclick="addToName('F')">F</button></td>
                  <td><button onclick="addToName('G')">G</button></td>
                  <td><button onclick="addToName('H')">H</button></td>
                  <td><button onclick="addToName('J')">J</button></td>
                  <td><button onclick="addToName('K')">K</button></td>
                  <td><button onclick="addToName('L')">L</button></td>
                  <td><button onclick="addToName(''')">'</button></td>
                  <td><button onclick="cancelName()">Back</button></td>
                </tr>
                <tr>
                  <td><button onclick="addToName('Z')">Z</button></td>
                  <td><button onclick="addToName('X')">X</button></td>
                  <td><button onclick="addToName('C')">C</button></td>
                  <td><button onclick="addToName('V')">V</button></td>
                  <td><button onclick="addToName('B')">B</button></td>
                  <td><button onclick="addToName('N')">N</button></td>
                  <td><button onclick="addToName('M')">M</button></td>
                  <td><button onclick="addToName('.')">.</button></td>
                  <td><button onclick="addToName(',')">,</button></td>
                  <td><button onclick="addToName(' ')">Space</button></td>
                  <td><button onclick="goPickupNoCustomer()">OK</button></td>
                </tr>
              </table>
            </div>
            <div id="phonediv" style="position: absolute; top: 128px; left: 0px; width: 810PX; visibility: <%=gsPhoneVisible%>;">
              <table align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="top">
                    <table align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td colspan="3"><div align="center">
                          <strong>AREA CODE</strong></div></td>
                      </tr>
                      <tr>
                        <td colspan="3"><div align="center">
                          <input type="text" id="areacode" autocomplete="off" onkeydown="disableEnterKey();" onfocus="setFocusAreaCode(true);" style="width: 100px; text-align: center; background-color: #cccccc;" value="<%=Left(gsPhone, 3)%>" /></div></td>
                      </tr>
<%
If ganAreaCodes(0) = 0 Then
%>
                      <tr>
                        <td><button onclick="setAreaCode('419')">419</button></td>
                        <td><button onclick="setAreaCode('567')">567</button></td>
                        <td><button onclick="setAreaCode('734')">734</button></td>
                      </tr>
<%
Else
%>
                      <tr>
<%
  For i = 0 To UBound(ganAreaCodes)
    If i > 0 And i Mod 3 = 0 Then
%>
                      </tr>
                      <tr>
<%
    End If
%>
                        <td><button onclick="setAreaCode('<%=ganAreaCodes(i)%>')"><%=ganAreaCodes(i)%></button></td>
<%
'   If i = 11 Then
'     Exit For
'   End If
  Next
  
' If UBound(ganAreaCodes) Mod 3 <> 0 Then
    For i = (UBound(ganAreaCodes) Mod 3) To 1
%>
                        <td>&nbsp;</td>
<%
    Next
' End If
%>
                      </tr>
<%
End If
%>
                      <tr>
                        <td colspan="3"><button style="width: 235px;" onclick="clearAreaCode()">Clear Area Code</button></td>
                      </tr>
                    </table>
                  </td>
                  <td valign="top" width="75"></td>
                  <td valign="top">
                    <table align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td colspan="3"><div align="center">
                          <strong>ENTER PHONE NUMBER</strong></div></td>
                      </tr>
                      <tr>
                        <td colspan="3"><div align="center">
                          <input type="text" id="phone" autocomplete="off" onkeydown="disableEnterKey();" onfocus="setFocusAreaCode(false);" style="width: 200px; text-align: center;" value="<%=Mid(gsPhone, 4, 3) & "-" & Mid(gsPhone, 7)%>"/></div></td>
                      </tr>
                      <tr>
                        <td><button onclick="addToPhone('1')">1</button></td>
                        <td><button onclick="addToPhone('2')">2</button></td>
                        <td><button onclick="addToPhone('3')">3</button></td>
                      </tr>
                      <tr>
                        <td><button onclick="addToPhone('4')">4</button></td>
                        <td><button onclick="addToPhone('5')">5</button></td>
                        <td><button onclick="addToPhone('6')">6</button></td>
                      </tr>
                      <tr>
                        <td><button onclick="addToPhone('7')">7</button></td>
                        <td><button onclick="addToPhone('8')">8</button></td>
                        <td><button onclick="addToPhone('9')">9</button></td>
                      </tr>
                      <tr>
                        <td><button onclick="cancelPhone()">Cancel</button></td>
                        <td><button onclick="addToPhone('0')">0</button></td>
                        <td><button onclick="backspacePhone()">Bksp</button></td>
                      </tr>
                      <tr>
                        <td colspan="3"><button style="width: 235px;" onclick="goNewPhone()">OK</button></td>
                      </tr>
                    </table>
                  </td>
                  <td valign="top" width="75">&nbsp;</td>
                  <td valign="top" width="235">&nbsp;</td>
                </tr>
              </table>
            </div>
            <div id="phoneconfirmdiv" style="position: absolute; top: 50px; left: 75px; width: 860px; height: 400px; z-index: 20; background: #fbf3c5; border: medium #000000 solid; visibility: <%=gsPhoneVisible%>;">
              <table align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="top" width="25">&nbsp;</td>
                  <td valign="top">&nbsp;</td>
                  <td valign="top" width="25">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" width="25">&nbsp;</td>
                  <td valign="top"><strong>This phone number does not match anything in the database. Ask the customer if they are sure that this is the correct number. Is this the correct number?</strong></td>
                  <td valign="top" width="25">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" width="25">&nbsp;</td>
                  <td valign="top">&nbsp;</td>
                  <td valign="top" width="25">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" width="25">&nbsp;</td>
                  <td valign="top">&nbsp;</td>
                  <td valign="top" width="25">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" width="25">&nbsp;</td>
                  <td valign="top" align="center"><button style="width: 235px;" onclick="cancelNewPhone()">Yes</button>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;<button style="width: 235px;" onclick="getNewPhone()">No</button></td>
                  <td valign="top" width="25">&nbsp;</td>
                </tr>
              </table>
            </div>
          </div>
        </td>
      </tr>
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
