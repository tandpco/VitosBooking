<%
Option Explicit
Response.buffer = TRUE

If Session("SecurityID") = "" Then
	Response.Redirect("/default.asp")
End If

If Request("o").Count > 0 Then
	If Not IsNumeric(Request("o")) Then
		Response.Redirect("../main.asp")
	End If
Else
	Response.Redirect("../main.asp")
End If

If Request("c").Count > 0 Then
	If Not IsNumeric(Request("c")) Then
		Response.Redirect("../ticket.asp?OrderID=" & Request("o"))
	End If
Else
	Response.Redirect("../ticket.asp?OrderID=" & Request("o"))
End If

'If Request("a").Count > 0 Then
'	If Not IsNumeric(Request("a")) Then
'		Response.Redirect("../ticket.asp?OrderID=" & Request("o"))
'	End If
'Else
'	Response.Redirect("../ticket.asp?OrderID=" & Request("o"))
'End If

If Request("action") = "SaveAddress" Then
	If Request("b").Count = 0 Then
		Response.Redirect("../ticket.asp?OrderID=" & Request("o"))
	Else
		If Len(Request("b")) = 0 Then
			Response.Redirect("../ticket.asp?OrderID=" & Request("o"))
		Else
			If Request("z").Count = 0 Then
				Response.Redirect("../ticket.asp?OrderID=" & Request("o"))
			Else
				If Len(Request("z")) = 0 Then
					Response.Redirect("../ticket.asp?OrderID=" & Request("o"))
				End If
			End If
		End If
	End If
End If



' Highlight the currently selected tab.
Dim currentTab
currentTab = "customer-name"
%>
<!-- #Include Virtual="include2/globals.asp" -->
<!-- #Include Virtual="include2/math.asp" -->
<!-- #Include Virtual="include2/db-connect.asp" -->
<!-- #Include Virtual="include2/customer.asp" -->
<!-- #Include Virtual="include2/address.asp" -->
<!-- #Include Virtual="include2/order.asp" -->
<%
Dim gnOrderID, gnCustomerID, gnAddressID
Dim gsEMail, gsFirstName, gsLastName, gdtBirthdate, gnPrimaryAddressID, gsHomePhone, gsCellPhone, gsWorkPhone, gsFAXPhone, gbIsEMailList, gbIsTextList
Dim gbNoChecks,gsExtension
Dim gnOrderStoreID, gsOrderAddress1, gsOrderAddress2, gsOrderCity, gsOrderState, gsOrderPostalCode, gsOrderAddressNotes
Dim ganAddressIDs(), gasAddresses()
Dim i,req
Dim gsTitle,optRedirect
Dim gnStoreID, gsPostalCode, gsAddress1, gsAddress2, gsCity, gsState, gdDeliveryCharge, gdDriverMoney, gnNewAddressID, gnStoreID2, gsAddressNotes, gbIsManual
Dim gnSessionID, gsIPAddress, gnEmpID, gsRefID, gdtTransactionDate, gdtSubmitDate, gdtReleaseDate, gdtExpectedDate, gnStoreID3, gnCustomerID2, gsCustomerName, gsCustomerPhone, gnAddressID2, gnOrderTypeID, gbIsPaid, gnPaymentTypeID, gsPaymentReference, gnAccountID, gdTax, gdTax2, gdTip, gnOrderStatusID, gsOrderNotes

gnOrderID = CLng(Request("o"))
gnCustomerID = CLng(Request("c"))
gnAddressID = CLng(Request("a"))
optRedirect = CStr(Request("afterEdit"))

If Request("action") = "savecustomer" Then
	gsTitle = "Edit Customer"
	
	gsEMail = Request("saveemail")
	gsFirstName = Request("savefirstname")
	gsLastName = Request("savelastname")
    gsExtension = Request("saveextension")

 	gdtBirthdate = Request("savebirthdate")
 	gsHomePhone = Request("savehomephone")
 	gsCellPhone = Request("savecellphone")
 	gsWorkPhone = Request("saveworkphone")
 	gsFAXPhone = Request("savefaxphone")
 	If Request("saveisemaillist") = "yes" Then
 		gbIsEMailList = TRUE
 	Else
 		gbIsEMailList = FALSE
 	End If
 	If Request("saveistextlist") = "yes" Then
 		gbIsTextList = TRUE
 	Else
 		gbIsTextList = FALSE
 	End If
 	If Request("savenochecks") = "yes" Then
 		gbNoChecks = TRUE
 	Else
 		gbNoChecks = FALSE
 	End If
	
	If gnCustomerID = 1 Then
		gnCustomerID = AddCustomer(gsEMail, "", gsFirstName, gsLastName, gdtBirthdate, 1, gsHomePhone, gsCellPhone, gsWorkPhone, gsFAXPhone, gbIsEMailList, gbIsTextList)
		
		If gnCustomerID = 0 Then
			Response.Redirect("/error.asp?nocustomer&err=" & Server.URLEncode(gsDBErrorMessage))
		Else
			If Not SetOrderCustomerID(gnOrderID, gnCustomerID) Then
				Response.Redirect("/error.asp?notsetorder&err=" & Server.URLEncode(gsDBErrorMessage))
			End If
		End If
	Else
		If Not UpdateCustomer_New2(gnCustomerID,gsEMail, gsFirstName, gsLastName,gdtBirthdate,gsHomePhone,gsCellPhone,gsWorkPhone,gsFAXPhone,gbIsEMailList,gbIsTextList,gbNoChecks, gsExtension) Then
			Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
		Else
			If Session("optRedirect") <> "" Then
				Response.Redirect(Session("optRedirect"))
			End If
		End If

	End If
	
	Session("ReturnURL") = "editcustomer.asp?o=" & gnOrderID & "&c=" & gnCustomerID & "&a=" & gnAddressID
	Session("SaveURL") = "editcustomer.asp?o=" & gnOrderID & "&c=" & gnCustomerID & "&a=" & gnAddressID
Else
	Session("optRedirect") = ""
	If(optRedirect <> "") Then
		Session("optRedirect") = optRedirect
	End if
	If gnCustomerID = 1 Then
		gsTitle = "<font color=""red"">Add Customer - Please Verify All Information</font>"
	Else
		gsTitle = "Edit Customer"
	End If
	
	If Request("action") = "setprimary" Then
		If Request("primary").Count > 0 Then
			If IsNumeric(Request("primary")) Then
				If Not SetPrimaryAddress(gnCustomerID, CLng(Request("primary"))) Then
					Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
				End If
			End If
		End If
	Else
		If Request("action") = "deleteaddress" Then
			If Request("delete").Count > 0 Then
				If IsNumeric(Request("delete")) Then
					If Not DeleteCustomerAddress(gnCustomerID, CLng(Request("delete"))) Then
						Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
					End If
				End If
			End If
		Else
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
						If Not LookupAddress(gsAddress1, gsAddress2, gsCity, gsState, gsPostalCode, gnNewAddressID, gnStoreID2, gsAddressNotes) Then
							Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
						End If
						
						If gnNewAddressID = 0 Then
							gnStoreID2 = gnStoreID
							gnNewAddressID = AddAddress(gnStoreID, gsAddress1, gsAddress2, gsCity, gsState, gsPostalCode, "", gbIsManual)
							If gnAddressID = 0 Then
								Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
							End If
						End If
						
						If Not AddCustomerAddress(gnCustomerID, gnNewAddressID, "Alternate Address") Then
							Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
						End If
					End If
				End If
			Else
				Session("ReturnURL") = "editcustomer.asp?o=" & gnOrderID & "&c=" & gnCustomerID & "&a=" & gnAddressID
				Session("SaveURL") = "editcustomer.asp?o=" & gnOrderID & "&c=" & gnCustomerID & "&a=" & gnAddressID
			End If
		End If
	End If
End If

If gnCustomerID = 1 Then
	If GetOrderDetails(gnOrderID, gnSessionID, gsIPAddress, gnEmpID, gsRefID, gdtTransactionDate, gdtSubmitDate, gdtReleaseDate, gdtExpectedDate, gnStoreID3, gnCustomerID2, gsCustomerName, gsCustomerPhone, gnAddressID2, gnOrderTypeID, gbIsPaid, gnPaymentTypeID, gsPaymentReference, gnAccountID, gdDeliveryCharge, gdDriverMoney, gdTax, gdTax2, gdTip, gnOrderStatusID, gsOrderNotes) Then
		gsEMail = ""
		If InStr(gsCustomerName, " ") = 0 Then
			gsFirstName = gsCustomerName
			gsLastName = ""
		Else
			gsFirstName = Left(gsCustomerName, (InStr(gsCustomerName, " ") - 1))
			gsLastName = Mid(gsCustomerName, (InStr(gsCustomerName, " ") + 1))
		End If
		gdtBirthdate = ""
		gnPrimaryAddressID = 1
		gsHomePhone = ""
		gsCellPhone = gsCustomerPhone
		gsWorkPhone = ""
		gsFAXPhone = ""
		gbIsEMailList = TRUE
		gbIsTextList = TRUE
	End If
Else
	If Not GetCustomerDetails(gnCustomerID, gsEMail, gsFirstName, gsLastName, gdtBirthdate, gnPrimaryAddressID, gsHomePhone, gsCellPhone, gsWorkPhone, gsFAXPhone, gbIsEMailList, gbIsTextList) Then
		Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
	End If
	
	If gdtBirthdate = "1/1/1900" Then
		gdtBirthdate = ""
	End If
	
	gbNoChecks = Not IsCustomerCheckOK(gnCustomerID)
End If

If Not GetAddressDetails(gnAddressID, gnOrderStoreID, gsOrderAddress1, gsOrderAddress2, gsOrderCity, gsOrderState, gsOrderPostalCode, gsOrderAddressNotes) Then
	'Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If Not GetCustomerAddresses(gnCustomerID, ganAddressIDs, gasAddresses) Then
	'Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
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
<script src="/include2/isDate.js" type="text/javascript"></script>
<script src="http://code.jquery.com/jquery-latest.js"></script>
<script type="text/javascript">
<!--
var ie4=document.all;

function resetRedirect() {
//	var loRedirectDiv;
//	
//	loRedirectDiv = ie4? eval("document.all.redirect") : document.getElementById("redirect");
//	loRedirectDiv.innerHTML = <%=gnRedirectTime%>;
}

function disableEnterKey() {
	var loText, loDiv;
	
	if (event.keyCode == 13) {
		event.cancelBubble = true;
		event.returnValue = false;
		return false;
	}
}

function toggleDivs(psHideDiv, psShowDiv) {
	var loHideDiv, loShowDiv;
	
	loHideDiv = ie4? eval("document.all." + psHideDiv) : document.getElementById(psHideDiv);
	loShowDiv = ie4? eval("document.all." + psShowDiv) : document.getElementById(psShowDiv);
	
	loHideDiv.style.visibility = "hidden";
	loShowDiv.style.visibility = "visible";
	
	resetRedirect();
}
//-->
</script>
    <link rel="stylesheet" href="/Scripts/keyboard/css/jsKeyboard.css" type="text/css" media="screen"/>
    <script type="text/javascript" src="/Scripts/keyboard/jsKeyboard.js"></script>
<script type="text/javascript">
var gsDeleteForm = "";


function GotoConfirmDelete(psTargetForm, psAddress) {
	gsDeleteForm = psTargetForm;
	
	$("#dispaddr").html(psAddress)
	$("#confirmdiv").show()
	$("#menudiv").hide()
	resetRedirect();
}

function CancelDelete() {
	$("#confirmdiv").hide()
	$("#menudiv").show()	
	resetRedirect();
}

function ConfirmDelete() {
	$("#"+gsDeleteForm).submit();
}

function goEditInfo() {
	
  jsKeyboard.show()
  $("#editTab").text("Save and Return to Order")
	$("#confirmdiv").hide()
	$("#orderlinenotes").show()
	$("#menudiv").hide()
	resetRedirect();
}

function cancelEditInfo() {
  jsKeyboard.hide()
  $("#editTab").text("Edit Details")
<%
If gnCustomerID = 1 Then
%>
	window.location = '../ticket.asp?OrderID=<%=gnOrderID%>';
<%
End If
%>
	
	$("#confirmdiv").hide()
	$("#orderlinenotes").hide()
	$("#menudiv").show()
	resetRedirect();

}


function properNotes() {
	resetRedirect();
}


function saveCustomer() {
	$("#formCustomer").submit();
}
$(function(){
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
	$("#editTab").click(function(){
		if($(this).is('.active'))
			return saveCustomer();
		goEditInfo()
		$("button.tabs").removeClass('active')
		$(this).addClass('active')
	})
	$("#addressTab").click(function(){
		cancelEditInfo()
		$("button.tabs").removeClass('active')
		$(this).addClass('active')
	})
	$("#editTab").click()
})
     
</script>
</head>

<body onload="clockInit(clockLocalStartTime, clockServerStartTime); clockOnLoad();" onunload="clockOnUnload()" style="padding-bottom:0">

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
							<div class="col width-50"><button class="tabs" id="editTab">Edit Details</button></div>
							<div class="col width-50"><button class="tabs" id="addressTab">Manage Addresses</button></div>
						</div>
						<div id="orderlinenotes">
							<form id="formCustomer" name="formCustomer" method="post" action="<%=Session("SaveURL")%>">
								<input type="hidden" id="action" name="action" value="savecustomer" />
								<input type="hidden" id="o" name="o" value="<%=gnOrderID%>" />
								<input type="hidden" id="c" name="c" value="<%=gnCustomerID%>" />
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
									<div class="col width-25">
										<label>Birth Date</label>
										<input type="text" id="savebirthdate" name="savebirthdate" value="<%=gdtBirthdate%>" class="newInput number" />
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
									<div class="col width-25" style="width:25%">
										<label>Extension / Building</label>
                		<input type="text" id="saveextension" name="saveextension" value="<%=gsExtension %>" class="newInput number" />
									</div>
								</div>
								<!-- <div class="row">
									<div class="col width-25">
										<label>Fax</label>
										<input type="text" id="savefaxphone" name="savefaxphone" value="<%=gsFAXPhone%>" class="newInput" />
									</div>
								</div> -->
							</form>
						</div>
						<div id="menudiv">
							<table width="925" align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="4">
<%
If gnAddressID = 1 Then
%>
										<strong>No address associated with this order.</strong>
<%
Else
%>
										<strong>Address on order: <%=gsOrderAddress1%>&nbsp;<%=gsOrderAddress2%>, <%=gsOrderCity%>, <%=gsOrderState%>&nbsp;<%=gsOrderPostalCode%></strong>
<%
End If
%>
									</td>
								</tr>
<%
If ganAddressIDs(0) <> 0 Then
%>
								<tr>
									<td colspan="4"><p><strong>Customer's Addresses:</strong></p></td>
								</tr>
<%
	If UBound(ganAddressIDs) = 0 And ganAddressIDs(0) = 1 Then
%>
								<tr>
									<td colspan="4"><p><strong>No addresses defined.</strong></p></td>
								</tr>
<%
	Else
		For i = 0 To UBound(ganAddressIDs)
			If ganAddressIDs(i) = gnPrimaryAddressID Then
%>
								<tr>
									<td><strong><%=gasAddresses(i)%></strong></td>
									<td width="250" colspan="2"><strong>&lt;&lt; Primary Address</strong></td>
									<td width="125" align="right"><button style="width: 100px;" onclick="window.location = 'addressnotes.asp?CustomerID=<%=gnCustomerID%>&AddressID=<%=ganAddressIDs(i)%>&OrderID=<%=gnOrderID%>'">Edit Notes</button></td>
								</tr>
<%
			End If
		Next
		
		For i = 0 To UBound(ganAddressIDs)
			If ganAddressIDs(i) <> gnPrimaryAddressID Then
%>
								<tr>
									<td><strong><%=gasAddresses(i)%></strong></td>
									<td width="125">
										<form id="formSetPrimary<%=i%>" name="formSetPrimary<%=i%>" method="post" action="editcustomer.asp">
											<input type="hidden" id="action" name="action" value="setprimary" />
											<input type="hidden" id="o" name="o" value="<%=gnOrderID%>" />
											<input type="hidden" id="c" name="c" value="<%=gnCustomerID%>" />
											<input type="hidden" id="a" name="a" value="<%=gnAddressID%>" />
											<input type="hidden" id="primary" name="primary" value="<%=ganAddressIDs(i)%>" />
											<button style="width: 100px;">Set As Primary</button>
										</form>
									</td>
									<td width="125" align="center">
										<form id="formDeleteAddress<%=i%>" name="formDeleteAddress<%=i%>" method="post" action="editcustomer.asp">
											<input type="hidden" id="action" name="action" value="deleteaddress" />
											<input type="hidden" id="o" name="o" value="<%=gnOrderID%>" />
											<input type="hidden" id="c" name="c" value="<%=gnCustomerID%>" />
											<input type="hidden" id="a" name="a" value="<%=gnAddressID%>" />
											<input type="hidden" id="delete" name="delete" value="<%=ganAddressIDs(i)%>" />
										</form>
										<button style="width: 100px;" onclick="GotoConfirmDelete('formDeleteAddress<%=i%>', '<%=gasAddresses(i)%>');">Remove From Customer</button>
									</td>
									<td width="125" align="right"><button style="width: 100px;" onclick="window.location = 'addressnotes.asp?CustomerID=<%=gnCustomerID%>&AddressID=<%=ganAddressIDs(i)%>&OrderID=<%=gnOrderID%>'">Edit Notes</button></td>
								</tr>
<%
			End If
		Next
	End If
End If
%>
								<tr>
									<td colspan="4"><button style="width: 925px;" onclick="window.location = 'newaddress.asp?c=<%=gnCustomerID%>&a=<%=gnAddressID%>&o=<%=gnOrderID%>';">Add a New Address For This Customer</button></td>
								</tr>
							</table>
						</div>

            <div id="virtualKeyboard"></div>
						
						<div id="confirmdiv" style="position: absolute; top: 150px; left: 0px; width: 1010px; background-color: #fbf3c5;display:none">
							<p align="center"><strong>Are you sure you want to delete the following address from this customer?</strong><br/><br/><strong><span id="dispaddr" name="dispaddr">Here</span></strong><br/><br/>
							<button onclick="ConfirmDelete();">Confirm</button>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<button onclick="CancelDelete();">Cancel</button></p>
						</div>
					</div>
					</div>
				</td>
			</tr>
		</table>
		</td>
	</tr>
</table>
</div>

<%
If gnCustomerID = 1 Then
%>
<script type="text/javascript">
<!--
goEditInfo();
//-->
</script>
<%
End If
%>
</body>

</html>
<!-- #Include Virtual="include2/db-disconnect.asp" -->
