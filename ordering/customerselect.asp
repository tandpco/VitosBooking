<%
Option Explicit
Response.buffer = TRUE

If Session("SecurityID") = "" Then
  Response.Redirect("/default.asp")
End If


%>
<!-- #Include Virtual="include2/globals.asp" --> <!-- #Include Virtual="include2/math.asp" --> <!-- #Include Virtual="include2/db-connect.asp" --> <!-- #Include Virtual="include2/customer.asp" --> <!-- #Include Virtual="include2/employee.asp" -->
<%
Dim gnOrderTypeID,  gsPhone,      ganCustomerIDs(), gasNames(),   i,                gnCustomerID, gnAddressID
Dim gasEMails(),    extensions(), gasPhones(),      gnAddressZip, gnAddressString,  currentTab

currentTab              = "customer-name"
gnOrderTypeID           = CLng(Request("t"))
gsPhone                 = Session("CustomerPhone")
gnAddressID             = Request("a")
Session("AddressID")    = gnAddressID
Session("OrderTypeID")  = gnOrderTypeID
Session("ReturnURL")    = "/ordering/customerfind.asp?t=" & gnOrderTypeID & "&p=" & gsPhone
Session("SaveURL")      = "/ordering/addressfind.asp?t=" & gnOrderTypeID & "&p=" & gsPhone

Call GetAddressDetails2(gnAddressID, gnAddressString, gnAddressZip)
Call GetCustomerPrimaryAddressDetails(gnAddressID, gasNames, ganCustomerIDs, gasEMails,extensions,gasPhones)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <meta content="en-us" http-equiv="Content-Language" />
    <meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
    <title>Vito's Point of Sale</title>
    <link rel="stylesheet" href="/css/vitos.css" type="text/css" />
    <!-- #Include Virtual="include2/clock-server.asp" -->
    <script type="text/javascript" src="http://code.jquery.com/jquery-latest.js"></script>
    <link rel="stylesheet" href="/Scripts/keyboard/css/jsKeyboard.css" type="text/css" media="screen"/>
    <script type="text/javascript" src="/Scripts/keyboard/jsKeyboard.js"></script>
    <script type="text/javascript">
      function showAll(el) {
        if(el.innerHTML == "Show All") {
          document.getElementById("addressList").className = 'showAll'
          el.innerHTML = 'Top 3 Only'
        } else {
          document.getElementById("addressList").className = ''
          el.innerHTML = 'Show All'
          $(el).closest('#content-wrapper').scrollTop('0')
        }
      }

      $(function () {
        jsKeyboard.init("virtualKeyboard");
        jsKeyboard.hide();

        var $firstInput = $('#livesearch').focus();
        jsKeyboard.currentElement = $firstInput;
        jsKeyboard.currentElementCursorPosition = 0;

        var Phone = "<%=gsPhone %>";

        $("#addressList .buttonLine").each(function(){
          console.log($(this).data('phones').toString())
          if($(this).data('phones').toString().indexOf(Phone) !== -1) {
            $(this).prependTo("#addressList").find('button').css('background-color','green')
          }
          $(this).data('text',$(this).find('.nameButton .name').text().toUpperCase())
        })
        $("#livesearch").on('change',function(){
          $("#content-wrapper").scrollTop(0)
          var $val = $(this).val()
          console.log('changed',$val)
          $("#addressList .buttonLine").each(function(){
            if($(this).data('text').indexOf($val) === -1) {
              $(this).find('.nameButton .name').html($(this).data('text'))
              $(this).hide()
            }
            else {
              $(this).find('.nameButton .name').html($(this).data('text').replace($val, '<span class="highlight">'+$val+'</span>'))
              $(this).show()
            }
          })

        })
      })

    </script>
  </head>
  <body style="padding-bottom:0">
    <div id="mainwindow" style="position: absolute; top: 0px; left: 0px; width=810PX; height: 768px; overflow: hidden;">
      <table cellspacing="0" cellpadding="0" width="1010" height="764" border="1">
        <tr>
          <td valign="top" width="1010" height="764">
            <table cellspacing="0" cellpadding="0" width="1010">
              <!-- #Include Virtual="ordering/top-header.asp" -->
              <tr>
                <td valign="top" width="1010">
                  <div id="content-wrapper">
                    <div id="content" align="center" style="position: relative;padding-bottom:400px">
                      <div id="assigndiv" align="center" style="position: relative; top: 0px; left: 0px; width: 810PX;">
                        <% If gnOrderTypeID = 1 Then %>
                          <div align="center"><strong>SELECT CUSTOMER FOR DELIVERY</strong></div><br/>
                        <% Else %>
                          <div align="center"><strong>SELECT CUSTOMER FOR PICKUP</strong></div><br/>
                        <% End If %>
                        <input type="text" id="livesearch" style="position:absolute;left:-1000px" />
                        <div id="virtualKeyboard"></div>
                        <div align="center"><strong>Most Recent Names <span style="color:green">*Green Items Match Phone Number Calling</span></strong></div>
                      <br/>
                        <div id="addressList">
                          <% If ganCustomerIDs(0) <> 0 Then
                            For i = 0 to UBound(ganCustomerIDs) %>
                            <div class="buttonLine" style="<%=IIf(extensions(i) <> "" and Session("Extension") <> "" and extensions(i) <> Session("Extension"),"opacity:0.5","") %>" data-phones="<%=gasPhones(i)%>">
                              <button style="width: 730px; text-align:left;"  onclick="window.location='unitselect.asp?t=<%=gnOrderTypeID%>&amp;c=<%=ganCustomerIDs(i)%>&amp;a=<%=gnAddressID%>'" class="nameButton">
                                <span class="name"><%=gasNames(i)%></span>
                                <%=IIf(extensions(i) <> ""," (ext. / bld. "+extensions(i)+")","") %>
                                <span style="float:right;display:inline-block;margin-right:10px;font-size:14px"><%=IIf(gasEmails(i) <> "",gasEmails(i),"No Email Yet") %></span>
                              </button>
                              <button style="width: 20px;" onclick="window.location='../custmaint/editcustomer.asp?c=<%=ganCustomerIDs(i)%>&amp;a=<%=gnAddressID%>&amp;o=0&amp;afterEdit=<%=Server.URLEncode(Request.ServerVariables("SCRIPT_NAME")& "?" & Request.QueryString)%>'" >Edit</button>
                            </div>
                          <% Next
                          End If %>
                        </div>
                        <div>
                          <% If UBound(ganCustomerIDs) > 3 Then %>
                          <button style="width:300px" onclick="showAll(this)">Show All</button>
                          <% End If %>
                          <button style="width:300px" onclick="window.location='addressfind.asp?t=<%=gnOrderTypeID%>&amp;z=<%=gnAddressZip%>&amp;b=<%=Server.URLEncode(gnAddressString)%>&amp;nn=yes'">Add New Customer Name</button>
                        </div>
                      </div>
                    </div>
                  </div>
                </td>
              </tr>
              <tr>
                <td valign="top" colspan="2" width="1010">
                  <div align="center">
                    <span class="orangetext">For technical assistance, please call 419.720.5050</span>
                  </div>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </div>
  </body>
</html>
<!-- #Include Virtual="include2/db-disconnect.asp" -->