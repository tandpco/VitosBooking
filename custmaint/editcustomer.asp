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

If Request("a").Count > 0 Then
	If Not IsNumeric(Request("a")) Then
		Response.Redirect("../ticket.asp?OrderID=" & Request("o"))
	End If
Else
	Response.Redirect("../ticket.asp?OrderID=" & Request("o"))
End If

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
Dim i
Dim gsTitle
Dim gnStoreID, gsPostalCode, gsAddress1, gsAddress2, gsCity, gsState, gdDeliveryCharge, gdDriverMoney, gnNewAddressID, gnStoreID2, gsAddressNotes, gbIsManual
Dim gnSessionID, gsIPAddress, gnEmpID, gsRefID, gdtTransactionDate, gdtSubmitDate, gdtReleaseDate, gdtExpectedDate, gnStoreID3, gnCustomerID2, gsCustomerName, gsCustomerPhone, gnAddressID2, gnOrderTypeID, gbIsPaid, gnPaymentTypeID, gsPaymentReference, gnAccountID, gdTax, gdTax2, gdTip, gnOrderStatusID, gsOrderNotes

gnOrderID = CLng(Request("o"))
gnCustomerID = CLng(Request("c"))
gnAddressID = CLng(Request("a"))

If Request("action") = "savecustomer" Then
	gsTitle = "Edit Customer"
	
	gsEMail = Request("saveemail")
	gsFirstName = Request("savefirstname")
	gsLastName = Request("savelastname")
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
			Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
		Else
			If Not SetOrderCustomerID(gnOrderID, gnCustomerID) Then
				Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
			End If
		End If
	Else
		If Not UpdateCustomer(gnCustomerID, gsEMail, gsFirstName, gsLastName, gdtBirthdate, gsHomePhone, gsCellPhone, gsWorkPhone, gsFAXPhone, gbIsEMailList, gbIsTextList, gbNoChecks) Then
			Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
		End If
	End If
	
	Session("ReturnURL") = "editcustomer.asp?o=" & gnOrderID & "&c=" & gnCustomerID & "&a=" & gnAddressID
	Session("SaveURL") = "editcustomer.asp?o=" & gnOrderID & "&c=" & gnCustomerID & "&a=" & gnAddressID
Else
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
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If Not GetCustomerAddresses(gnCustomerID, ganAddressIDs, gasAddresses) Then
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
<script src="/include2/isDate.js" type="text/javascript"></script>
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
<script type="text/javascript">
<!--
var gbNotesInLowerCase = false;
var gsCurrentField = "firstname";
var gsDeleteForm = "";

function setCurrentField(psCurrentField) {
	var loField;
	
	loField = ie4? eval("document.all." + gsCurrentField) : document.getElementById(gsCurrentField);
	loField.style.backgroundColor = "#CCCCCC";
	
	loField = ie4? eval("document.all." + psCurrentField) : document.getElementById(psCurrentField);
	loField.style.backgroundColor = "#FFFFFF";
	
	gsCurrentField = psCurrentField;
	
	resetRedirect();
}

function GotoConfirmDelete(psTargetForm, psAddress) {
	var loSpan, loDiv;
	
	gsDeleteForm = psTargetForm;
	
	loSpan = ie4? eval("document.all.dispaddr") : document.getElementById('dispaddr');
	loSpan.innerHTML = psAddress;
	
	loDiv = ie4? eval("document.all.menudiv") : document.getElementById('menudiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.kbdiv") : document.getElementById('kbdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.confirmdiv") : document.getElementById('confirmdiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function CancelDelete() {
	var loDiv;

	loDiv = ie4? eval("document.all.kbdiv") : document.getElementById('kbdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.confirmdiv") : document.getElementById('confirmdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.menudiv") : document.getElementById('menudiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function ConfirmDelete() {
	var loForm;

	loForm = ie4? eval("document.all." + gsDeleteForm) : document.getElementById(gsDeleteForm);
	loForm.submit();
}

function goEditInfo() {
	var loField, loDiv;

	loField = ie4? eval("document.all.firstname") : document.getElementById("firstname");
	loField.disabled = false;
	loField = ie4? eval("document.all.lastname") : document.getElementById("lastname");
	loField.disabled = false;
	loField = ie4? eval("document.all.birthdate") : document.getElementById("birthdate");
	loField.disabled = false;
	loField = ie4? eval("document.all.email") : document.getElementById("email");
	loField.disabled = false;
	loField = ie4? eval("document.all.homephone") : document.getElementById("homephone");
	loField.disabled = false;
	loField = ie4? eval("document.all.cellphone") : document.getElementById("cellphone");
	loField.disabled = false;
	loField = ie4? eval("document.all.workphone") : document.getElementById("workphone");
	loField.disabled = false;
	loField = ie4? eval("document.all.faxphone") : document.getElementById("faxphone");
	loField.disabled = false;
	loField = ie4? eval("document.all.isemaillist") : document.getElementById("isemaillist");
	loField.disabled = false;
	loField = ie4? eval("document.all.istextlist") : document.getElementById("istextlist");
	loField.disabled = false;
	loField = ie4? eval("document.all.nochecks") : document.getElementById("nochecks");
	loField.disabled = false;
	
	setCurrentField("firstname");
	
	loDiv = ie4? eval("document.all.menudiv") : document.getElementById('menudiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.confirmdiv") : document.getElementById('confirmdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.kbdiv") : document.getElementById('kbdiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function cancelEditInfo() {
<%
If gnCustomerID = 1 Then
%>
	window.location = '../ticket.asp?OrderID=<%=gnOrderID%>';
<%
Else
%>
	var loField, loDiv;

	loField = ie4? eval("document.all.firstname") : document.getElementById("firstname");
	loField.disabled = true;
	loField.value = "<%=gsFirstName%>";
	loField = ie4? eval("document.all.lastname") : document.getElementById("lastname");
	loField.disabled = true;
	loField.value = "<%=gsLastName%>";
	loField = ie4? eval("document.all.birthdate") : document.getElementById("birthdate");
	loField.disabled = true;
	loField.value = "<%=gdtBirthdate%>";
	loField = ie4? eval("document.all.email") : document.getElementById("email");
	loField.disabled = true;
	loField.value = "<%=gsEMail%>";
	loField = ie4? eval("document.all.homephone") : document.getElementById("homephone");
	loField.disabled = true;
	loField.value = "<%=gsHomePhone%>";
	loField = ie4? eval("document.all.cellphone") : document.getElementById("cellphone");
	loField.disabled = true;
	loField.value = "<%=gsCellPhone%>";
	loField = ie4? eval("document.all.workphone") : document.getElementById("workphone");
	loField.disabled = true;
	loField.value = "<%=gsWorkPhone%>";
	loField = ie4? eval("document.all.faxphone") : document.getElementById("faxphone");
	loField.disabled = true;
	loField.value = "<%=gsFAXPhone%>";
	loField = ie4? eval("document.all.isemaillist") : document.getElementById("isemaillist");
	loField.disabled = true;
<%
If gbIsEMailList Then
%>
	loField.checked = true;
<%
Else
%>
	loField.checked = false;
<%
End If
%>
	loField = ie4? eval("document.all.istextlist") : document.getElementById("istextlist");
	loField.disabled = true;
<%
If gbIsTextList Then
%>
	loField.checked = true;
<%
Else
%>
	loField.checked = false;
<%
End If
%>
	loField = ie4? eval("document.all.nochecks") : document.getElementById("nochecks");
	loField.disabled = true;
<%
If gbNoChecks Then
%>
	loField.checked = true;
<%
Else
%>
	loField.checked = false;
<%
End If
%>
	
	loField = ie4? eval("document.all." + gsCurrentField) : document.getElementById(gsCurrentField);
	loField.style.backgroundColor = "#CCCCCC";
	
	loDiv = ie4? eval("document.all.kbdiv") : document.getElementById('kbdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.confirmdiv") : document.getElementById('confirmdiv');
	loDiv.style.visibility = "hidden";
	loDiv = ie4? eval("document.all.menudiv") : document.getElementById('menudiv');
	loDiv.style.visibility = "visible";
	
	resetRedirect();
<%
End If
%>
}

function addToNotes(psDigit) {
	var loNotes, lsNotes;
	
	loNotes = ie4? eval("document.all." + gsCurrentField) : document.getElementById(gsCurrentField);
	lsNotes = loNotes.value;
	
	if (psDigit.length > 1) {
		if (lsNotes.length > 0) {
			lsNotes = lsNotes + " ";
		}
	}
	
	if (gbNotesInLowerCase) {
		lsNotes += psDigit.toLowerCase();
	}
	else {
		lsNotes += psDigit;
	}
	
	if (lsNotes.length > 255) {
		lsNotes = lsNotes.substr(0, 255);
	}
	
	loNotes.value = lsNotes;
	
	resetRedirect();
}

function backspaceNotes() {
	var loNotes, lsNotes;
	
	loNotes = ie4? eval("document.all." + gsCurrentField) : document.getElementById(gsCurrentField);
	lsNotes = loNotes.value;
	if (lsNotes.length > 0) {
		lsNotes = lsNotes.substr(0, (lsNotes.length - 1));
		loNotes.value = lsNotes;
	}
	
	resetRedirect();
}

function clearCurrentField() {
	var loNotes;
	
	loNotes = ie4? eval("document.all." + gsCurrentField) : document.getElementById(gsCurrentField);
	loNotes.value = "";
	
	resetRedirect();
}

function properNotes() {
	var loNotes, i, lsText, lbDoUpper;
	
	loNotes = ie4? eval("document.all." + gsCurrentField) : document.getElementById(gsCurrentField);
	if (loNotes.value.length > 0) {
		lsText = loNotes.value.substr(0, 1).toUpperCase();
		lbDoUpper = false;
		for (i = 1; i < loNotes.value.length; i++) {
			if (loNotes.value.substr(i, 1) == " ") {
				lsText += " ";
				lbDoUpper = true;
			}
			else {
				if (lbDoUpper) {
					lsText += loNotes.value.substr(i, 1).toUpperCase();
					lbDoUpper = false;
				}
				else {
					lsText += loNotes.value.substr(i, 1).toLowerCase();
				}
			}
		}
		loNotes.value = lsText;
	}
	
	resetRedirect();
}

function shiftNotes() {
	var loObj;
	
	if (gbNotesInLowerCase) {
		loObj = ie4? eval("document.all.key-q") : document.getElementById('key-q');
		loObj.innerHTML = "Q";
		loObj = ie4? eval("document.all.key-w") : document.getElementById('key-w');
		loObj.innerHTML = "W";
		loObj = ie4? eval("document.all.key-e") : document.getElementById('key-e');
		loObj.innerHTML = "E";
		loObj = ie4? eval("document.all.key-r") : document.getElementById('key-r');
		loObj.innerHTML = "R";
		loObj = ie4? eval("document.all.key-t") : document.getElementById('key-t');
		loObj.innerHTML = "T";
		loObj = ie4? eval("document.all.key-y") : document.getElementById('key-y');
		loObj.innerHTML = "Y";
		loObj = ie4? eval("document.all.key-u") : document.getElementById('key-u');
		loObj.innerHTML = "U";
		loObj = ie4? eval("document.all.key-i") : document.getElementById('key-i');
		loObj.innerHTML = "I";
		loObj = ie4? eval("document.all.key-o") : document.getElementById('key-o');
		loObj.innerHTML = "O";
		loObj = ie4? eval("document.all.key-p") : document.getElementById('key-p');
		loObj.innerHTML = "P";
		loObj = ie4? eval("document.all.key-a") : document.getElementById('key-a');
		loObj.innerHTML = "A";
		loObj = ie4? eval("document.all.key-s") : document.getElementById('key-s');
		loObj.innerHTML = "S";
		loObj = ie4? eval("document.all.key-d") : document.getElementById('key-d');
		loObj.innerHTML = "D";
		loObj = ie4? eval("document.all.key-f") : document.getElementById('key-f');
		loObj.innerHTML = "F";
		loObj = ie4? eval("document.all.key-g") : document.getElementById('key-g');
		loObj.innerHTML = "G";
		loObj = ie4? eval("document.all.key-h") : document.getElementById('key-h');
		loObj.innerHTML = "H";
		loObj = ie4? eval("document.all.key-j") : document.getElementById('key-j');
		loObj.innerHTML = "J";
		loObj = ie4? eval("document.all.key-k") : document.getElementById('key-k');
		loObj.innerHTML = "K";
		loObj = ie4? eval("document.all.key-l") : document.getElementById('key-l');
		loObj.innerHTML = "L";
		loObj = ie4? eval("document.all.key-z") : document.getElementById('key-z');
		loObj.innerHTML = "Z";
		loObj = ie4? eval("document.all.key-x") : document.getElementById('key-x');
		loObj.innerHTML = "X";
		loObj = ie4? eval("document.all.key-c") : document.getElementById('key-c');
		loObj.innerHTML = "C";
		loObj = ie4? eval("document.all.key-v") : document.getElementById('key-v');
		loObj.innerHTML = "V";
		loObj = ie4? eval("document.all.key-b") : document.getElementById('key-b');
		loObj.innerHTML = "B";
		loObj = ie4? eval("document.all.key-n") : document.getElementById('key-n');
		loObj.innerHTML = "N";
		loObj = ie4? eval("document.all.key-m") : document.getElementById('key-m');
		loObj.innerHTML = "M";
	}
	else {
		loObj = ie4? eval("document.all.key-q") : document.getElementById('key-q');
		loObj.innerHTML = "q";
		loObj = ie4? eval("document.all.key-w") : document.getElementById('key-w');
		loObj.innerHTML = "w";
		loObj = ie4? eval("document.all.key-e") : document.getElementById('key-e');
		loObj.innerHTML = "e";
		loObj = ie4? eval("document.all.key-r") : document.getElementById('key-r');
		loObj.innerHTML = "r";
		loObj = ie4? eval("document.all.key-t") : document.getElementById('key-t');
		loObj.innerHTML = "t";
		loObj = ie4? eval("document.all.key-y") : document.getElementById('key-y');
		loObj.innerHTML = "y";
		loObj = ie4? eval("document.all.key-u") : document.getElementById('key-u');
		loObj.innerHTML = "u";
		loObj = ie4? eval("document.all.key-i") : document.getElementById('key-i');
		loObj.innerHTML = "i";
		loObj = ie4? eval("document.all.key-o") : document.getElementById('key-o');
		loObj.innerHTML = "o";
		loObj = ie4? eval("document.all.key-p") : document.getElementById('key-p');
		loObj.innerHTML = "p";
		loObj = ie4? eval("document.all.key-a") : document.getElementById('key-a');
		loObj.innerHTML = "a";
		loObj = ie4? eval("document.all.key-s") : document.getElementById('key-s');
		loObj.innerHTML = "s";
		loObj = ie4? eval("document.all.key-d") : document.getElementById('key-d');
		loObj.innerHTML = "d";
		loObj = ie4? eval("document.all.key-f") : document.getElementById('key-f');
		loObj.innerHTML = "f";
		loObj = ie4? eval("document.all.key-g") : document.getElementById('key-g');
		loObj.innerHTML = "g";
		loObj = ie4? eval("document.all.key-h") : document.getElementById('key-h');
		loObj.innerHTML = "h";
		loObj = ie4? eval("document.all.key-j") : document.getElementById('key-j');
		loObj.innerHTML = "j";
		loObj = ie4? eval("document.all.key-k") : document.getElementById('key-k');
		loObj.innerHTML = "k";
		loObj = ie4? eval("document.all.key-l") : document.getElementById('key-l');
		loObj.innerHTML = "l";
		loObj = ie4? eval("document.all.key-z") : document.getElementById('key-z');
		loObj.innerHTML = "z";
		loObj = ie4? eval("document.all.key-x") : document.getElementById('key-x');
		loObj.innerHTML = "x";
		loObj = ie4? eval("document.all.key-c") : document.getElementById('key-c');
		loObj.innerHTML = "c";
		loObj = ie4? eval("document.all.key-v") : document.getElementById('key-v');
		loObj.innerHTML = "v";
		loObj = ie4? eval("document.all.key-b") : document.getElementById('key-b');
		loObj.innerHTML = "b";
		loObj = ie4? eval("document.all.key-n") : document.getElementById('key-n');
		loObj.innerHTML = "n";
		loObj = ie4? eval("document.all.key-m") : document.getElementById('key-m');
		loObj.innerHTML = "m";
	}
	
	gbNotesInLowerCase = !gbNotesInLowerCase;
	
	resetRedirect();
}

function saveCustomer() {
	var loNotes, loFormNotes, loForm, lbHasPhone;
	
	lbHasPhone = false;
	
	loNotes = ie4? eval("document.all.email") : document.getElementById('email');
	loFormNotes = ie4? eval("document.all.saveemail") : document.getElementById('saveemail');
	loFormNotes.value = loNotes.value;
	loNotes = ie4? eval("document.all.firstname") : document.getElementById('firstname');
	loFormNotes = ie4? eval("document.all.savefirstname") : document.getElementById('savefirstname');
	loFormNotes.value = loNotes.value;
	loNotes = ie4? eval("document.all.lastname") : document.getElementById('lastname');
	loFormNotes = ie4? eval("document.all.savelastname") : document.getElementById('savelastname');
	loFormNotes.value = loNotes.value;
	loNotes = ie4? eval("document.all.birthdate") : document.getElementById('birthdate');
	if (loNotes.value.length > 0) {
		if (!isDate(loNotes.value)) {
			setCurrentField("birthdate");
			return false;
		}
	}
	loFormNotes = ie4? eval("document.all.savebirthdate") : document.getElementById('savebirthdate');
	loFormNotes.value = loNotes.value;
	loNotes = ie4? eval("document.all.homephone") : document.getElementById('homephone');
	if (isNaN(loNotes.value)) {
		alert("Please enter numbers only for home phone.");
		setCurrentField("homephone");
		return false;
	}
	if (loNotes.value.length > 0) {
		if (loNotes.value.length == 10) {
			lbHasPhone = true;
		}
		else {
			alert("Home phone must be 10 digits.");
			setCurrentField("homephone");
			return false;
		}
	}
	loFormNotes = ie4? eval("document.all.savehomephone") : document.getElementById('savehomephone');
	loFormNotes.value = loNotes.value;
	loNotes = ie4? eval("document.all.cellphone") : document.getElementById('cellphone');
	if (isNaN(loNotes.value)) {
		alert("Please enter numbers only for cell phone.");
		setCurrentField("cellphone");
		return false;
	}
	if (loNotes.value.length > 0) {
		if (loNotes.value.length == 10) {
			lbHasPhone = true;
		}
		else {
			alert("Cell phone must be 10 digits.");
			setCurrentField("cellphone");
			return false;
		}
	}
	loFormNotes = ie4? eval("document.all.savecellphone") : document.getElementById('savecellphone');
	loFormNotes.value = loNotes.value;
	loNotes = ie4? eval("document.all.workphone") : document.getElementById('workphone');
	if (isNaN(loNotes.value)) {
		alert("Please enter numbers only for work phone.");
		setCurrentField("workphone");
		return false;
	}
	if (loNotes.value.length > 0) {
		if (loNotes.value.length == 10) {
			lbHasPhone = true;
		}
		else {
			alert("Work phone must be 10 digits.");
			setCurrentField("workphone");
			return false;
		}
	}
//	if (!lbHasPhone) {
//		alert("A home, cell or work phone is required.");
//		setCurrentField("homephone");
//		return false;
//	}
	loFormNotes = ie4? eval("document.all.saveworkphone") : document.getElementById('saveworkphone');
	loFormNotes.value = loNotes.value;
	loNotes = ie4? eval("document.all.faxphone") : document.getElementById('faxphone');
	if (isNaN(loNotes.value)) {
		alert("Please enter numbers only for FAX phone.");
		setCurrentField("faxphone");
		return false;
	}
	if (loNotes.value.length > 0) {
		if (loNotes.value.length != 10) {
			alert("FAX phone must be 10 digits.");
			setCurrentField("faxphone");
			return false;
		}
	}
	loFormNotes = ie4? eval("document.all.savefaxphone") : document.getElementById('savefaxphone');
	loFormNotes.value = loNotes.value;
	loNotes = ie4? eval("document.all.isemaillist") : document.getElementById('isemaillist');
	loFormNotes = ie4? eval("document.all.saveisemaillist") : document.getElementById('saveisemaillist');
	if (loNotes.checked) {
		loFormNotes.value = "yes";
	} else {
		loFormNotes.value = "no";
	}
	loNotes = ie4? eval("document.all.istextlist") : document.getElementById('istextlist');
	loFormNotes = ie4? eval("document.all.saveistextlist") : document.getElementById('saveistextlist');
	if (loNotes.checked) {
		loFormNotes.value = "yes";
	} else {
		loFormNotes.value = "no";
	}
	loNotes = ie4? eval("document.all.nochecks") : document.getElementById('nochecks');
	loFormNotes = ie4? eval("document.all.savenochecks") : document.getElementById('savenochecks');
	if (loNotes.checked) {
		loFormNotes.value = "yes";
	} else {
		loFormNotes.value = "no";
	}
	
	loForm = ie4? eval("document.all.formCustomer") : document.getElementById("formCustomer");
	loForm.submit();
}

//-->
</script>
    <style type="text/css">
        .auto-style2 {
            width: 439px;
        }
        .auto-style3 {
            width: 556px;
        }
        .auto-style6 {
            width: 74px;
        }
    </style>
</head>

<body onload="clockInit(clockLocalStartTime, clockServerStartTime); clockOnLoad();" onunload="clockOnUnload()">

<div id="mainwindow" style="position: absolute; top: 0px; left: 0px; width=1010px; height: 768px; overflow: hidden;">
<table cellspacing="0" cellpadding="0" width="1010" height="764" border="1">
	<tr>
		<td valign="top" width="1010" height="764">
		<table cellspacing="0" cellpadding="5" width="1010">
			<tr height="31">
				<td valign="top" width="1010">
					<div align="center">
                        <ol id="tabs">
						    <li><a onclick="back2Delivery();" title="Delivery">Delivery</a></li>
						    <li><a onclick="back2Phone();" title="Phone">Phone</a></li>
						    <li><a onclick="back2Adx();" title="Phone">Address</a></li>
						    <li class="active">Customer Name</li>
						    <li>Order</li>
						    <li>Notes</li>
						</ol>						

<%
If gbTestMode Then
	If gbDevMode Then
%>
						<strong>DEV SYSTEM <%
	Else
%>
						<strong>TEST SYSTEM
<%
	End If
End If
%>
						Store <%=Session("StoreID")%></strong> |
						<b><%=Session("name")%></b> |
						<span id="ClockDate"><%=clockDateString(gDate)%></span> 
						|
						<span id="ClockTime" onclick="clockToggleSeconds()"><%=clockTimeString(Hour(gDate), Minute(gDate), Second(gDate))%></span>
					</div>
				</td>
			</tr>
			<tr height="733">
				<td valign="top" width="1010">
					<div id="content" style="position: relative; width: 1010px; height: 723px; overflow: auto;">
						<div id="orderlinenotes" style="position: absolute; left: 0px; top: 0px; width: 1010px; height: 723px; background-color: #fbf3c5;">
							<form id="formCustomer" name="formCustomer" method="post" action="editcustomer.asp">
								<input type="hidden" id="action" name="action" value="savecustomer" />
								<input type="hidden" id="o" name="o" value="<%=gnOrderID%>" />
								<input type="hidden" id="c" name="c" value="<%=gnCustomerID%>" />
								<input type="hidden" id="a" name="a" value="<%=gnAddressID%>" />
								<input type="hidden" id="saveemail" name="saveemail" value="<%=gsEMail%>" />
								<input type="hidden" id="savefirstname" name="savefirstname" value="<%=gsFirstName%>" />
								<input type="hidden" id="savelastname" name="savelastname" value="<%=gsLastName%>" />
								<input type="hidden" id="savebirthdate" name="savebirthdate" value="<%=gdtBirthdate%>" />
								<input type="hidden" id="savehomephone" name="savehomephone" value="<%=gsHomePhone%>" />
								<input type="hidden" id="savecellphone" name="savecellphone" value="<%=gsCellPhone%>" />
								<input type="hidden" id="saveworkphone" name="saveworkphone" value="<%=gsWorkPhone%>" />
								<input type="hidden" id="savefaxphone" name="savefaxphone" value="<%=gsFAXPhone%>" />
								<input type="hidden" id="saveisemaillist" name="saveisemaillist" value="<%If gbIsEmailList Then Response.Write("yes") Else Response.Write("no") End If%>" />
								<input type="hidden" id="saveistextlist" name="saveistextlist" value="<%If gbIsTextList Then Response.Write("yes") Else Response.Write("no") End If%>" />
								<input type="hidden" id="savenochecks" name="savenochecks" value="<%If gbNoChecks Then Response.Write("yes") Else Response.Write("no") End If%>" />
								<p align="center"><strong><%=gsTitle%></strong></p>
							</form>
							<table width="1000" align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td>
										<table align="left" cellpadding="0" cellspacing="0" style="width: 995px">
											<tr>
												<td align="left" class="auto-style2"><strong>First Name</strong></td>
												<td align="left" class="auto-style3" colspan="3"><strong>Last Name:</strong></td>											</tr>
											<tr>
												<td align="left" class="auto-style2"><input type="text" value="<%=gsFirstName%>" id="firstname" disabled autocomplete="off" onkeydown="disableEnterKey();" onfocus="setCurrentField('firstname');" style="width: 427px; background-color: #cccccc; margin-left: 0px;" /></td>
												<td align="left" class="auto-style3" colspan="3"><input type="text" value="<%=gsLastName%>" id="lastname" disabled autocomplete="off" onkeydown="disableEnterKey();" onfocus="setCurrentField('lastname');" style="width: 552px; background-color: #cccccc;" /></td>
                                            </tr>
											<tr>
												<td align="left" class="auto-style2"><strong>E-Mail</strong></td>
                                                <td align="left" class="auto-style3" colspan="3"><strong>Extension or Department</strong></td>
											</tr>
											<tr>
												<td align="left" class="auto-style2"><input type="text" value="<%=gsEMail%>" id="email" disabled autocomplete="off" onkeydown="disableEnterKey();" onfocus="setCurrentField('firstname');" style="width: 427px; background-color: #cccccc; margin-left: 0px;" /></td>
												<td align="left" class="auto-style3" colspan="3"><input type="text" value="<%=gsExtension%>" id="extension" disabled autocomplete="off" onkeydown="disableEnterKey();" onfocus="setCurrentField('lastname');" style="width: 552px; background-color: #cccccc;" /></td>
                                            </tr>
                                            <tr>
                                                <td>&nbsp</td>
                                                <td class="auto-style6">Text Me:</td>
                                                <td class="auto-style3"><input id="txtYes" type="checkbox" />Yes<input id="txtNo" type="checkbox" />No</td>
                                            </tr>
                                        </table>
									</td>
								</tr>
							</table>
						</div>
						<div id="menudiv" style="position: absolute; top: 150px; left: 0px; width: 1010px;">
							<table width="450" align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td><button style="width: 100px;" onclick="window.location = '../ticket.asp?OrderID=<%=gnOrderID%>'">Done</button></td>
									<td align="right"><button style="width: 100px;" onclick="goEditInfo();">Edit Customer Info</button></td>
								</tr>
							</table>
							<p>&nbsp;</p>
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
						<div id="kbdiv" style="position: absolute; top: 150px; left: 0px; width: 1010px; visibility: hidden;">
							<table width="925" align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td><button style="width: 100px;" onclick="cancelEditInfo();">Cancel</button></td>
									<td align="center"><button style="width: 100px;" onclick="clearCurrentField()">Clear</button></td>
									<td align="right"><button style="width: 100px;" onclick="saveCustomer()">Done</button></td>
								</tr>
							</table>
							<table align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td><div align="center">
										<button onclick="addToNotes('+')">+</button><button onclick="addToNotes('!')">!</button><button onclick="addToNotes('@')">@</button><button onclick="addToNotes('#')">#</button><button onclick="addToNotes('$')">$</button><button onclick="addToNotes('%')">%</button><button onclick="addToNotes('^')">^</button><button onclick="addToNotes('&')">&amp;</button><button onclick="addToNotes('*')">*</button><button onclick="addToNotes('(')">(</button><button onclick="addToNotes(')')">)</button><button onclick="addToNotes(':')">:</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="addToNotes('=')">=</button><button onclick="addToNotes('1')">1</button><button onclick="addToNotes('2')">2</button><button onclick="addToNotes('3')">3</button><button onclick="addToNotes('4')">4</button><button onclick="addToNotes('5')">5</button><button onclick="addToNotes('6')">6</button><button onclick="addToNotes('7')">7</button><button onclick="addToNotes('8')">8</button><button onclick="addToNotes('9')">9</button><button onclick="addToNotes('0')">0</button><button onclick="addToNotes('?')">?</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="addToNotes('\'')">'</button><button name="key-q" id="key-q" onclick="addToNotes('Q')">Q</button><button name="key-w" id="key-w" onclick="addToNotes('W')">W</button><button name="key-e" id="key-e" onclick="addToNotes('E')">E</button><button name="key-r" id="key-r" onclick="addToNotes('R')">R</button><button name="key-t" id="key-t" onclick="addToNotes('T')">T</button><button name="key-y" id="key-y" onclick="addToNotes('Y')">Y</button><button name="key-u" id="key-u" onclick="addToNotes('U')">U</button><button name="key-i" id="key-i" onclick="addToNotes('I')">I</button><button name="key-o" id="key-o" onclick="addToNotes('O')">O</button><button name="key-p" id="key-p" onclick="addToNotes('P')">P</button><button onclick="addToNotes('/')">/</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="addToNotes('.')">.</button><button name="key-a" id="key-a" onclick="addToNotes('A')">A</button><button name="key-s" id="key-s" onclick="addToNotes('S')">S</button><button name="key-d" id="key-d" onclick="addToNotes('D')">D</button><button name="key-f" id="key-f" onclick="addToNotes('F')">F</button><button name="key-g" id="key-g" onclick="addToNotes('G')">G</button><button name="key-h" id="key-h" onclick="addToNotes('H')">H</button><button name="key-j" id="key-j" onclick="addToNotes('J')">J</button><button name="key-k" id="key-k" onclick="addToNotes('K')">K</button><button name="key-l" id="key-l" onclick="addToNotes('L')">L</button><button onclick="addToNotes(',')">,</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="shiftNotes()">Shift</button><button onclick="addToNotes('<')">&lt;</button><button name="key-z" id="key-z" onclick="addToNotes('Z')">Z</button><button name="key-x" id="key-x" onclick="addToNotes('X')">X</button><button name="key-c" id="key-c" onclick="addToNotes('C')">C</button><button name="key-v" id="key-v" onclick="addToNotes('V')">V</button><button name="key-b" id="key-b" onclick="addToNotes('B')">B</button><button name="key-n" id="key-n" onclick="addToNotes('N')">N</button><button name="key-m" id="key-m" onclick="addToNotes('M')">M</button><button onclick="addToNotes('>')">&gt;</button><button onclick="properNotes()">Proper</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button style="width: 150px;" onclick="addToNotes('1/2')">1/2</button>&nbsp;<button style="width: 600px;" onclick="addToNotes(' ')">Space</button>&nbsp;<button style="width: 150px;" onclick="backspaceNotes()">Bksp</button></div></td>
								</tr>
							</table>
						</div>
						<div id="confirmdiv" style="position: absolute; top: 150px; left: 0px; width: 1010px; background-color: #fbf3c5; visibility: hidden;">
							<p align="center"><strong>Are you sure you want to delete the following address from this customer?</strong><br/><br/><strong><span id="dispaddr" name="dispaddr">Here</span></strong><br/><br/>
							<button onclick="ConfirmDelete();">Confirm</button>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<button onclick="CancelDelete();">Cancel</button></p>
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
