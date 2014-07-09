<%
Response.buffer = TRUE

If Session("SecurityID") = "" Then
  Response.Redirect("/default.asp")
End If

currentTab = "customer-name"
%>
<!-- #Include Virtual="include2/globals.asp" -->
<!-- #Include Virtual="include2/math.asp" -->
<!-- #Include Virtual="include2/db-connect.asp" -->
<!-- #Include Virtual="include2/customer.asp" -->
<!-- #Include Virtual="include2/address.asp" -->
<!-- #Include Virtual="include2/order.asp" -->
<%
Dim gsPostalCode, gsAddress1, gsAddress2, gsCity, gsState, gdDeliveryCharge, gdDriverMoney
gnAddressID     = CLng(Request("a"))
gnOrderID       = Session("OrderID")
gnOrderTypeID   = Session("OrderTypeID")
gsCustomerPhone = Session("CustomerPhone")
If Request("action") = "SaveAddress" Then
  gsAddress1 = Request("b")
  gsPostalCode = Request("z")
  If Request("Manual").Count > 0 Then
    If Request("Manual") = "Yes" Then
      gbIsManual = TRUE
    Else
      gbIsManual = FALSE
    End If
  Else
    gbIsManual = FALSE
  End If
  
  gnStoreID = GetStoreByAddress(gsPostalCode, gsAddress1, gsAddress2, gsCity, gsState, gdDeliveryCharge, gdDriverMoney)
  If gnStoreID = -1 Then
    Response.Redirect("/error.asp?err=" & Server.URLEncode("Unable to narrow down a specific address."))
  Else
    If gsCity = "UNKNOWN CITY" And gsState = "US" Then
      Response.Redirect("/error.asp?err=" & Server.URLEncode("No city/state data is available for that zip code."))
    Else
      If Not LookupAddress(gsAddress1, gsAddress2, gsCity, gsState, gsPostalCode, gnAddressID, gnStoreID2, gsAddressNotes) Then
        Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
      End If
      
      If gnAddressID = 0 Then
        gnStoreID2 = gnStoreID
        gnAddressID = AddAddress(gnStoreID, gsAddress1, gsAddress2, gsCity, gsState, gsPostalCode, "", gbIsManual)
        If gnAddressID = 0 Then
          Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
        End If
      End If
      
    End If
  End If
End If
If Request("action") = "savecustomer" Then
  gsEMail       = Request("saveemail")
  gsFirstName   = Request("savefirstname")
  gsLastName    = Request("savelastname")
  gsExtension   = Request("saveextension")
  gdtBirthdate  = Request("savebirthdate")
  gsHomePhone   = Request("savehomephone")
  gsCellPhone   = Request("savecellphone")
  gsWorkPhone   = Request("saveworkphone")
  gsFAXPhone    = Request("savefaxphone")
  gbIsEMailList = Iif(Request("saveisemaillist") = "yes", TRUE,FALSE)
  gbIsTextList  = Iif(Request("saveistextlist") = "yes",  TRUE,FALSE)
  gbNoChecks    = Iif(Request("savenochecks") = "yes",    TRUE,FALSE)
  gnCustomerID  = AddCustomer(gsEMail, "", gsFirstName, gsLastName, gdtBirthdate, 1, gsHomePhone, gsCellPhone, gsWorkPhone, gsFAXPhone, gbIsEMailList, gbIsTextList) 'gsExtension / noChecks
  If gnAddressID <> 0 Then
    call AddCustomerAddress(gnCustomerID, gnAddressID, "Primary Address")
  End If
  If Session("optRedirect") <> "" Then
    'Response.Redirect(Session("optRedirect"))
  Else
    If gnAddressID <> 0 Then
      Response.Redirect("/ordering/unitselect.asp?t="&gnOrderTypeID&"&c="&gnCustomerID&"&a="&gnAddressID)
    Else
      Response.Redirect("/ordering/unitselect.asp?t="&gnOrderTypeID&"&c="&gnCustomerID)
    End If
  End If
Else
  Session("optRedirect") = Iif(optRedirect <> "", optRedirect, "")
  
End If

  gsEMail = ""
  gsFirstName = ""
  gsLastName = ""
  gdtBirthdate = ""
  gnPrimaryAddressID = 1
  gsHomePhone = ""
  gsCellPhone = gsCustomerPhone
  gsWorkPhone = ""
  gsFAXPhone = ""
  gbIsEMailList = TRUE
  gbIsTextList = TRUE

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

  <head>
    <meta content="en-us" http-equiv="Content-Language" />
    <meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
    <title>Vito's Point of Sale</title>
    <link rel="stylesheet" href="/css/vitos.css" type="text/css" />
    <!-- #Include Virtual="include2/clock-server.asp" -->
    <script src="/include2/isDate.js" type="text/javascript"></script>
    <script src="http://code.jquery.com/jquery-latest.js"></script>
    <link rel="stylesheet" href="/Scripts/keyboard/css/jsKeyboard.css" type="text/css" media="screen"/>
    <script type="text/javascript" src="/Scripts/keyboard/jsKeyboard.js"></script>
    <script type="text/javascript">

      function saveCustomer() {
        $("#formCustomer").submit();
      }
      $(function(){
        $("#addAddress").click(function(){
          saveCustomer();
        })
        jsKeyboard.init("virtualKeyboard");
        $("input[type=text]").on('focus',function(){
          jsKeyboard.show()
          if($(this).is('.number'))
            jsKeyboard.changeToNumber();
          else
            jsKeyboard.changeToCapitalLetter();
        })
        $(".toggleButton").each(function(){
          var $toggle = $(this)
          if(!$(this).data('req'))
            return;
          var $req = $("#"+$(this).data('req'));
          if($req.val() == '')
            $(this).addClass('disabled')
          $req.on('keyup change blur',function(){
            if($(this).val() == '')
              $(".toggleButton[data-req="+$(this).attr('id')+"]").addClass('disabled')
            else
              $(".toggleButton[data-req="+$(this).attr('id')+"]").removeClass('disabled')
          })
        })
        $(".toggleButton").on('click',function(){
          if($("#"+$(this).data('req')).val() == '')
            return false;
          if($(this).is('.active')) {
            $(this).removeClass('active')
            $(this).next().val($(this).data('off'))
          } else {
            $(this).addClass('active')
            $(this).next().val($(this).data('on'))
          }
        })
        $("#phoneEditShow").click(function(){
          $("#phoneEdit").show()
          $("#extraEdit").hide()
        })
        $("#extraEditShow").click(function(){
          $("#phoneEdit").hide()
          $("#extraEdit").show()
        })
      })
           
    </script>
  </head>

<body style="padding-bottom:0">

<div id="mainwindow" style="position: absolute; top: 0px; left: 0px; width=1010px; height: 768px; overflow: hidden;">
<table cellspacing="0" cellpadding="0" width="1010" height="764" border="1">
  <tr>
    <td valign="top" width="1010" height="764">
    <table cellspacing="0" cellpadding="0" width="1010">
      <!-- #Include Virtual="ordering/top-header.asp" -->
      <tr height="733">
        <td valign="top" width="1010">
          <div id="content-wrapper" style="top:0">
          <div id="content">
            <div class="row">
              <div class="col width-50" style="width:70%"><button class="tabs active" id="addAddress">Add Name And Return To Order</button></div>
              <div class="col width-50" style="width:30%"><button class="tabs" id="addressTab" onclick="window.location='/ordering/customerfind.asp?t=<%=gnOrderTypeID%>&p=<%=gsCustomerPhone%>'">Cancel</button></div>
            </div>
            <div id="orderlinenotes">
              <form id="formCustomer" name="formCustomer" method="post" action="addcustomer.asp?action=savecustomer">
                <input type="hidden" id="action" name="action" value="savecustomer" />
                <input type="hidden" id="a" name="a" value="<%=gnAddressID%>" />
                <div class="row">
                  <div class="col width-25">
                    <label>First Name</label>
                    <input type="text" id="savefirstname" name="savefirstname" value="<%=gsFirstName%>" class="newInput" />
                  </div>
                  <div class="col width-25">
                    <label>Last Name</label>
                    <input type="text" id="savelastname" name="savelastname" value="<%=gsLastName%>" class="newInput" />
                  </div>
                  <div class="col width-25">
                    <label>Email</label>
                    <input type="text" name="saveemail" id="saveemail" value="<%=gsEMail%>" />
                  </div>
                  <div class="col width-25" style="width:25%">
                    <label>Extension / Building</label>
                    <input type="text" id="saveextension" name="saveextension" value="<%=gsExtension %>" class="newInput number" />
                  </div>
                </div>
                <div class="row" id="extraEdit">
                  <div class="col width-25">
                    <label>&nbsp;</label>
                    <input type="button" value="Edit Phones" class="newInput" id="phoneEditShow" />
                  </div>
                  <div class="col width-25">
                    <label>Put on Email List?</label>
                    <div class="toggleButton<%=Iif(gbIsEmailList," active", "")%>" data-on="yes" data-off="no" data-req="saveemail"></div>
                    <input type="hidden" id="saveisemaillist" name="saveisemaillist" value="<%If gbIsEmailList Then Response.Write("yes") Else Response.Write("no") End If%>" />
                  </div>
                  <div class="col width-25">
                    <label>Text Me?</label>
                    <div class="toggleButton<%=Iif(gbIsTextList," active", "")%>" data-on="yes" data-off="no" data-req="savecellphone"></div>
                    <input type="hidden" id="saveistextlist" name="saveistextlist" value="<%If gbIsTextList Then Response.Write("yes") Else Response.Write("no") End If%>" />
                  </div>
                  <div class="col width-25">
                    <label>No Checks</label>
                    <div class="toggleButton<%=Iif(gbNoChecks," active", "")%>" data-on="yes" data-off="no"></div>
                    <input type="hidden" id="savenochecks" name="savenochecks" value="<%If gbNoChecks Then Response.Write("yes") Else Response.Write("no") End If%>" />
                  </div>
                </div>
                <div class="row" style="display:none" id="phoneEdit">
                  <div class="col width-25" style="width:15%">
                    <label>&nbsp;</label>
                    <input type="button" value="&laquo; Back" class="newInput" id="extraEditShow"/>
                  </div>
                  <div class="col width-25" style="width:20%">
                    <label>Home Phone</label>
                    <input type="text" id="savehomephone" name="savehomephone" value="<%=gsHomePhone%>" class="newInput number" />
                  </div>
                  <div class="col width-25" style="width:20%">
                    <label>Cell Phone</label>
                    <input type="text" id="savecellphone" name="savecellphone" value="<%=gsCellPhone%>" class="newInput number" />
                  </div>
                  <div class="col width-25" style="width:20%">
                    <label>Work Phone</label>
                    <input type="text" id="saveworkphone" name="saveworkphone" value="<%=gsWorkPhone%>" class="newInput number" />
                  </div>
                  <div class="col width-25">
                    <label>Birth Date</label>
                    <input type="text" id="savebirthdate" name="savebirthdate" value="<%=gdtBirthdate%>" class="newInput number" />
                  </div>
                </div>
              </form>
            </div>

            <div id="virtualKeyboard"></div>
            
          </div>
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
