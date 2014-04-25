<%
Option Explicit
Response.buffer = TRUE

If Session("SecurityID") = "" Then
	Response.Redirect("/default.asp")
End If

If Request("l").Count = 0 And Request("u").Count = 0 Then
	Response.Redirect("unitselect.asp")
End If

If Request("l").Count > 0 Then
	If Not IsNumeric(Request("l")) Then
		Response.Redirect("unitselect.asp")
	End If
Else
	If Not IsNumeric(Request("u")) Then
		Response.Redirect("unitselect.asp")
	End If
End If

If Request("s").Count > 0 Then
	If Not IsNumeric(Request("s")) Then
		Response.Redirect("unitselect.asp")
	End If
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
Dim gnOrderID, gnSessionID, gsIPAddress, gsRefID, gdtTransactionDate, gdtSubmitDate, gdtReleaseDate, gdtExpectedDate, gnStoreID, gnCustomerID, gsCustomerName, gsCustomerPhone, gnAddressID, gnOrderTypeID, gbIsPaid, gnPaymentTypeID, gsPaymentReference, gdDeliveryCharge, gdDriverMoney, gdTax, gdTax2, gdTip, gnOrderStatusID, gsOrderNotes
Dim gbQuickMode
Dim gsOrderTypeDescription
Dim gsEMail, gsFirstName, gsLastName, gdtBirthdate, gnPrimaryAddressID, gsHomePhone, gsCellPhone, gsWorkPhone, gsFAXPhone, gbIsEMailList, gbIsTextList
Dim gsAddress1, gsAddress2, gsCity, gsState, gsPostalCode, gsAddressNotes
Dim gsAddressDescription, gsCustomerNotes
Dim gdOrderTotal
Dim gnOrderLineID, gnUnitID, gnSpecialtyID, gnSizeID, gnStyleID, gnHalf1SauceID, gnHalf2SauceID, gnHalf1SauceModifierID, gnHalf2SauceModifierID, gsOrderLineNotes, gnQuantity, gdOrderLineCost, gdOrderLineDiscount, gnCouponID
Dim gsUnitDescription, gsSpecialtyDescription, gsSizeDescription, gsStyleDescription, gsHalf1SauceDescription, gsHalf2SauceDescription, gsHalf1SauceModifierDescription, gsHalf2SauceModifierDescription
Dim ganOrderLineItemIDs(), ganItemIDs(), ganHalfIDs(), gasItemDescriptions(), gasItemShortDescriptions()
Dim ganOrderLineTopperIDs(), ganTopperIDs(), ganTopperHalfIDs(), gasTopperDescriptions(), gasTopperShortDescriptions()
Dim ganOrderLineFreeSideIDs(), ganFreeSideIDs(), gasFreeSideDescriptions(), gasFreeSideShortDescriptions()
Dim ganOrderLineAddSideIDs(), ganAddSideIDs(), gasAddSideDescriptions(), gasAddSideShortDescriptions()
Dim ganUnitSizeIDs(), gasUnitSizeDescriptions(), gasUnitSizeShortDescriptions(), gadUnitSizeStandardBasePrice(), ganUnitSizeStandardNumberIncludedItems(), gadUnitSizeSpecialtyBasePrice(), ganUnitSizeSpecialtyNumberIncludedItems(), ganUnitSizePercentSpecialtyItemVariances(), gadPerAdditionalItemPrices(), gabIsTaxable()
Dim ganSizeStyleIDs(), gasSizeStyleDescriptions(), gasSizeStyleShortDescriptions(), gasSizeStyleSpecialMessage(), ganSizeStyleSizeIDs(), gadSizeStyleSurcharges()
Dim ganSauceIDs(), gasSauceDescriptions(), gasSauceShortDescriptions()
Dim ganSauceModifierIDs(), gasSauceModifierDescriptions(), gasSauceModifierShortDescriptions()
Dim ganUnitItemIDs(), gasUnitItemDescriptions(), gasUnitItemShortDescriptions(), gadUnitItemOnSidePrice(), ganUnitItemCounts(), gabUnitFreeItemFlags(), gabUnitItemIsCheeses(), gabUnitItemIsBaseCheeses(), gabUnitItemIsExtraCheeses()
Dim ganUnitTopperIDs(), gasUnitTopperDescriptions(), gasUnitTopperShortDescriptions()
Dim ganUnitSideIDs(), gasUnitSideDescriptions(), gasUnitSideShortDescriptions(), gadUnitSidePrices()
Dim ganUnitSpecialtyIDs(), gasUnitSpecialtyDescriptions(), gasUnitSpecialtyShortDescriptions(), ganUnitSpecialtySauceID(), ganUnitSpecialtyStyleID(), gabSpecialtyNoBaseCheese(), ganUnitSpecialtyItemIDs(), ganUnitSpecialtyItemQuantity()
Dim ganUpchargeSizeIDs(), ganUpchargeItemIDs(), gadUpchargePrice()
Dim ganSideGroupSpecialtyIDs(), ganSideGroupSizeIDs(), ganSideGroupSideGroupIDs(), gadSideGroupQuantity()
Dim ganUnitGroupSizeIDs(), ganUnitGroupSideGroupIDs(), gadUnitGroupQuantity()
Dim ganSideGroupIDs(), gasSideGroupDescriptions(), gasSideGroupShortDescriptions(), ganSideGroupSideIDs(), gasSideGroupSideDescriptions(), gasSideGroupSideShortDescriptions(), gabSideGroupSideIsDefault()
Dim i, j, k, l, m, n, lsTmp, lnTop, lnLeft, gnPos, gnCount, gbFound
Dim ganOrderIDs(), gsLocalErrorMsg
Dim gbNeedPrinterAlert
Dim ganEmpIDs(), ganEmployeeIDs(), gasCardIDs()
Dim gasMPOReasons()

' 2012-10-01 TAM: Don't release hold orders during ordering process in case of stuck hold order
'If Not ReleaseHoldOrders(Session("StoreID"), Session("TransactionDate"), ganOrderIDs) Then
'	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
'End If
'
'If ganOrderIDs(0) > 0 Then
'	For i = 0 To UBound(ganOrderIDs)
'		If Not PrintOrder(Session("StoreID"), ganOrderIDs(i), TRUE) Then
'			ResetHoldOrder ganOrderIDs(i)
'			gbNeedPrinterAlert = TRUE
'			gsLocalErrorMsg = "PRINT FAILURE, CANNOT RELEASE HOLD ORDERS!"
'		End If
'	Next
'Else
	gbNeedPrinterAlert = FALSE
'End If

gnStoreID = Session("StoreID")
gnOrderID = Session("OrderID")
gnSessionID = Session("SessionID")
gsIPAddress = Session("IPAddress")
gsRefID = Session("RefID")
gdtSubmitDate = Session("SubmitDate")
gdtReleaseDate = Session("ReleaseDate")
gdtExpectedDate = Session("ExpectedDate")
gbIsPaid = Session("IsPaid")
gnPaymentTypeID = Session("PaymentTypeID")
gsPaymentReference = Session("PaymentReference")
gdDeliveryCharge = Session("DeliveryCharge")
gdDriverMoney = Session("DriverMoney")
gdTax = Session("Tax")
gdTax2 = Session("Tax2")
gdTip = Session("Tip")
gnOrderStatusID = Session("OrderStatusID")
gsOrderNotes = Session("OrderNotes")
gsCustomerPhone = Session("CustomerPhone")
gbQuickMode = Session("QuickMode")
gnOrderTypeID = CLng(Session("OrderTypeID"))
gsOrderTypeDescription = Session("OrderTypeDescription")
gnCustomerID = CLng(Session("CustomerID"))
gsEMail = Session("EMail")
gsFirstName = Session("FirstName")
gsLastName = Session("LastName")
gdtBirthdate = Session("Birthdate")
gnPrimaryAddressID = CLng(Session("PrimaryAddressID"))
gsHomePhone = Session("HomePhone")
gsCellPhone = Session("CellPhone")
gsWorkPhone = Session("WorkPhone")
gsFAXPhone = Session("FAXPhone")
gbIsEMailList = Session("IsEmailList")
gbIsTextList = Session("IsTextList")
gsCustomerName = Session("CustomerName")
gsCustomerName = Session("CustomerName")
gnAddressID = CLng(Session("AddressID"))
gsAddress1 = Session("Address1")
gsAddress2 = Session("Address2")
gsCity = Session("City")
gsState = Session("State")
gsPostalCode = Session("PostalCode")
gsAddressNotes = Session("AddressNotes")
gsAddressDescription = Session("AddressDescription")
gsCustomerNotes = Session("CustomerNotes")
gdOrderTotal = Session("OrderTotal")

If Request("l").Count > 0 Then
	gnOrderLineID = CLng(Request("l"))
	If Not GetOrderLineDetails(gnOrderLineID, gnUnitID, gnSpecialtyID, gnSizeID, gnStyleID, gnHalf1SauceID, gnHalf2SauceID, gnHalf1SauceModifierID, gnHalf2SauceModifierID, gsOrderLineNotes, gnQuantity, gdOrderLineCost, gdOrderLineDiscount, gnCouponID) Then
		If Len(gsDBErrorMessage) > 0 Then
			Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
		Else
			Response.Redirect("/error.asp?err=" & Server.URLEncode("Invalid Order Line Specified"))
		End If
	End If
	
	gdOrderTotal = gdOrderTotal - (gnQuantity * (gdOrderLineCost - gdOrderLineDiscount))
	gsUnitDescription = GetUnitShortDescription(gnUnitID)
	gsSpecialtyDescription = GetSpecialtyDescription(gnSpecialtyID)
	gsSizeDescription = GetSizeDescription(gnSizeID)
	gsStyleDescription = GetStyleDescription(gnStyleID)
	gsHalf1SauceDescription = GetSauceDescription(gnHalf1SauceID)
	gsHalf2SauceDescription = GetSauceDescription(gnHalf2SauceID)
	gsHalf1SauceModifierDescription = GetSauceModifierDescription(gnHalf1SauceModifierID)
	gsHalf2SauceModifierDescription = GetSauceModifierDescription(gnHalf2SauceModifierID)
	
	If Not GetOrderLineItems(gnOrderLineID, gnUnitID, ganOrderLineItemIDs, ganItemIDs, ganHalfIDs, gasItemDescriptions, gasItemShortDescriptions) Then
		Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
	End If
	
	If Not GetOrderLineToppers(gnOrderLineID, ganOrderLineTopperIDs, ganTopperIDs, ganTopperHalfIDs, gasTopperDescriptions, gasTopperShortDescriptions) Then
		Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
	End If
	
	If Not GetOrderLineFreeSides(gnOrderLineID, ganOrderLineFreeSideIDs, ganFreeSideIDs, gasFreeSideDescriptions, gasFreeSideShortDescriptions) Then
		Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
	End If
	
	If Not GetOrderLineAddSides(gnOrderLineID, ganOrderLineAddSideIDs, ganAddSideIDs, gasAddSideDescriptions, gasAddSideShortDescriptions) Then
		Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
	End If
Else
	gnOrderLineID = 0
	gnUnitID = CLng(Request("u"))
	If Request("s").Count > 0 Then
		gnSizeID = CLng(Request("s"))
		gsSizeDescription = GetSizeShortDescription(gnSizeID)
	Else
		gnSizeID = 0
		gsSizeDescription = ""
	End If
	
	gnSpecialtyID = 0
	gnStyleID = 0
	gnHalf1SauceID = 0
	gnHalf2SauceID = 0
	gnHalf1SauceModifierID = 0
	gnHalf2SauceModifierID = 0
	gsOrderLineNotes = ""
	gnQuantity = 1
	gdOrderLineCost = 0.00
	gdOrderLineDiscount = 0.00
	gnCouponID = 0
	
	gsUnitDescription = GetUnitShortDescription(gnUnitID)
	gsSpecialtyDescription = ""
	gsStyleDescription = ""
	gsHalf1SauceDescription = ""
	gsHalf2SauceDescription = ""
	gsHalf1SauceModifierDescription = ""
	gsHalf2SauceModifierDescription = ""
	
	ReDim ganItemIDs(0), ganHalfIDs(0), gasItemDescriptions(0), gasItemShortDescriptions(0)
	ganItemIDs(0) = 0
	ganHalfIDs(0) = 0
	gasItemDescriptions(0) = ""
	gasItemShortDescriptions(0) = ""
	
	ReDim ganTopperIDs(0), ganTopperHalfIDs(0), gasTopperDescriptions(0), gasTopperShortDescriptions(0)
	ganTopperIDs(0) = 0
	ganTopperHalfIDs(0) = 0
	gasTopperDescriptions(0) = ""
	gasTopperShortDescriptions(0) = ""
	
	ReDim ganFreeSideIDs(0), gasFreeSideDescriptions(0), gasFreeSideShortDescriptions(0)
	ganFreeSideIDs(0) = 0
	gasFreeSideDescriptions(0) = ""
	gasFreeSideShortDescriptions(0) = ""
	
	ReDim ganAddSideIDs(0), gasAddSideDescriptions(0), gasAddSideShortDescriptions(0)
	ganAddSideIDs(0) = 0
	gasAddSideDescriptions(0) = ""
	gasAddSideShortDescriptions(0) = ""
End If

If Not GetStoreUnitSizes(gnStoreID, gnUnitID, ganUnitSizeIDs, gasUnitSizeDescriptions, gasUnitSizeShortDescriptions, gadUnitSizeStandardBasePrice, ganUnitSizeStandardNumberIncludedItems, gadUnitSizeSpecialtyBasePrice, ganUnitSizeSpecialtyNumberIncludedItems, ganUnitSizePercentSpecialtyItemVariances, gadPerAdditionalItemPrices, gabIsTaxable) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If Not GetStoreSizeStyles(gnStoreID, gnUnitID, ganUnitSizeIDs, ganSizeStyleIDs, gasSizeStyleDescriptions, gasSizeStyleShortDescriptions, gasSizeStyleSpecialMessage, ganSizeStyleSizeIDs, gadSizeStyleSurcharges) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If Not GetStoreUnitSauces(gnStoreID, gnUnitID, ganSauceIDs, gasSauceDescriptions, gasSauceShortDescriptions) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If Not GetSauceModifiers(ganSauceModifierIDs, gasSauceModifierDescriptions, gasSauceModifierShortDescriptions) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If Not GetStoreUnitItems(gnStoreID, gnUnitID, ganUnitItemIDs, gasUnitItemDescriptions, gasUnitItemShortDescriptions, gadUnitItemOnSidePrice, ganUnitItemCounts, gabUnitFreeItemFlags, gabUnitItemIsCheeses, gabUnitItemIsBaseCheeses, gabUnitItemIsExtraCheeses) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If Not GetStoreUnitToppers(gnStoreID, gnUnitID, ganUnitTopperIDs, gasUnitTopperDescriptions, gasUnitTopperShortDescriptions) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If Not GetStoreUnitSides(gnStoreID, gnUnitID, ganUnitSideIDs, gasUnitSideDescriptions, gasUnitSideShortDescriptions, gadUnitSidePrices) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If Not GetStoreUnitSpecialties(gnStoreID, gnUnitID, ganUnitSpecialtyIDs, gasUnitSpecialtyDescriptions, gasUnitSpecialtyShortDescriptions, ganUnitSpecialtySauceID, ganUnitSpecialtyStyleID, gabSpecialtyNoBaseCheese, ganUnitSpecialtyItemIDs, ganUnitSpecialtyItemQuantity) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If Not GetUpchargeItems(gnStoreID, gnUnitID, ganUpchargeSizeIDs, ganUpchargeItemIDs, gadUpchargePrice) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If Not GetSpecialtySizeSideGroups(gnStoreID, gnUnitID, ganSideGroupSpecialtyIDs, ganSideGroupSizeIDs, ganSideGroupSideGroupIDs, gadSideGroupQuantity) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If Not GetUnitSizeSideGroups(gnStoreID, gnUnitID, ganUnitGroupSizeIDs, ganUnitGroupSideGroupIDs, gadUnitGroupQuantity) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If Not GetSideGroups(ganSideGroupIDs, gasSideGroupDescriptions, gasSideGroupShortDescriptions, ganSideGroupSideIDs, gasSideGroupSideDescriptions, gasSideGroupSideShortDescriptions, gabSideGroupSideIsDefault) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If Not GetAllStoreManagers(gnStoreID, ganEmpIDs, ganEmployeeIDs, gasCardIDs) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If Request("l").Count = 0 Then
	If Request("s").Count = 0 Then
		gnSizeID = ganUnitSizeIDs(0)
		gsSizeDescription = gasUnitSizeShortDescriptions(0)
	End If
	gnStyleID = ganSizeStyleIDs(0)
	gsStyleDescription = gasSizeStyleShortDescriptions(0)
	gbFound = FALSE
	For i = 0 To UBound(ganSizeStyleIDs)
		For j = 0 To UBound(ganSizeStyleSizeIDs, 2)
			If ganSizeStyleSizeIDs(i, j) = gnSizeID Then
				gnStyleID = ganSizeStyleIDs(i)
				gsStyleDescription = gasSizeStyleShortDescriptions(i)
				gbFound = TRUE
				Exit For
			End If
		Next
		If gbFound Then
			Exit For
		End If
	Next
	
	If gnUnitID = 1 Then
		gnHalf1SauceID = ganSauceIDs(0)
		gnHalf2SauceID = ganSauceIDs(0)
		gsHalf1SauceDescription = gasSauceShortDescriptions(0)
		gsHalf2SauceDescription = gasSauceShortDescriptions(0)
	Else
		gnHalf1SauceID = 0
		gnHalf2SauceID = 0
		gsHalf1SauceDescription = ""
		gsHalf2SauceDescription = ""
	End If
	
	gdOrderLineCost = gadUnitSizeStandardBasePrice(0)
	
	For i = 0 To UBound(ganUnitItemIDs)
		If gabUnitItemIsBaseCheeses(i) Then
			ganItemIDs(0) = ganUnitItemIDs(i)
			ganHalfIDs(0) = 0
			gasItemDescriptions(0) = gasUnitItemDescriptions(i)
			gasItemShortDescriptions(0) = gasUnitItemShortDescriptions(i)
			
			Exit For
		End If
	Next
	
	gnPos = 0
	For i = 0 To UBound(ganUnitGroupSizeIDs)
		If ganUnitGroupSizeIDs(i) = gnSizeID Then
			If ganUnitGroupSideGroupIDs(i, 0) <> 0 Then
				For j = 0 To UBound(ganUnitGroupSideGroupIDs, 2)
					For k = 0 To UBound(ganSideGroupIDs)
						If ganSideGroupIDs(k) = ganUnitGroupSideGroupIDs(i, j) Then
							If ganSideGroupSideIDs(k, 0) <> 0 Then
								For l = 0 To (gadUnitGroupQuantity(i, j) - 1)
									gnCount = 0
									gbFound = FALSE
									For m = 0 To UBound(ganSideGroupSideIDs, 2)
										If gabSideGroupSideIsDefault(k, m) Then
											If gnCount = gnPos Then
												ReDim Preserve ganFreeSideIDs(gnPos), gasFreeSideDescriptions(gnPos), gasFreeSideShortDescriptions(gnPos)
												
												ganFreeSideIDs(gnPos) = ganSideGroupSideIDs(k, m)
												gasFreeSideDescriptions(gnPos) = gasSideGroupSideDescriptions(k, m)
												gasFreeSideShortDescriptions(gnPos) = gasSideGroupSideShortDescriptions(k, m)
												
												gnPos = gnPos + 1
												gbFound = TRUE
												
												Exit For
											End If
											
											gnCount = gnCount + 1
										End If
									Next
									If Not gbFound Then
										For m = 0 To UBound(ganSideGroupSideIDs, 2)
											If gabSideGroupSideIsDefault(k, m) Then
												ReDim Preserve ganFreeSideIDs(gnPos), gasFreeSideDescriptions(gnPos), gasFreeSideShortDescriptions(gnPos)
												
												ganFreeSideIDs(gnPos) = ganSideGroupSideIDs(k, m)
												gasFreeSideDescriptions(gnPos) = gasSideGroupSideDescriptions(k, m)
												gasFreeSideShortDescriptions(gnPos) = gasSideGroupSideShortDescriptions(k, m)
												
												gnPos = gnPos + 1
												gbFound = TRUE
												
												Exit For
											End If
										Next
									End If
									If Not gbFound Then
										ReDim Preserve ganFreeSideIDs(gnPos), gasFreeSideDescriptions(gnPos), gasFreeSideShortDescriptions(gnPos)
										
										ganFreeSideIDs(gnPos) = ganSideGroupSideIDs(k, 0)
										gasFreeSideDescriptions(gnPos) = gasSideGroupSideDescriptions(k, 0)
										gasFreeSideShortDescriptions(gnPos) = gasSideGroupSideShortDescriptions(k, 0)
										
										gnPos = gnPos + 1
									End If
								Next
							End If
							
							Exit For
						End If
					Next
				Next
			End If
			
			Exit For
		End If
	Next
End If

If Not GetMPOReasons(gasMPOReasons) Then
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
<script type="text/javascript">
<!--
var ie4=document.all;

function resetRedirect() {
//	var loRedirectDiv;
//	
//	loRedirectDiv = ie4? eval("document.all.redirect") : document.getElementById("redirect");
//	loRedirectDiv.innerHTML = <%=gnRedirectTime%>;
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
<!-- #Include Virtual="include2/jsmenu.asp" -->
<script type="text/javascript">
<!--
var gnHalfID = 0;
var gbDeleteMode = false;
var gbNeedIncludedSides = true;
var gdQuantity = <%=gnQuantity%>;
var gbHasQuantityPrice = false;
var gbNotesInLowerCase = false;
var gbManualStyle = false;
var gbIncludedSidesDoneIsDone = false;
var gbAddAnother = false;
var gbGetPrice = false;
var gbNeedStyle = true;
var gbNeedStyleDoneIsDone = false;
var gbNeedToppers = false;
var gbNeedTopperDoneIsDone = false;
var gbDonePressed = false;
var gbMPOReasonInLowerCase = false;
var gsMPOReason = "";

<%
If Request("l").Count <> 0 Then
%>
gbNeedIncludedSides = false;
gbNeedStyle = false;
gbNeedToppers = false;
<%
Else
%>
// TODO: Make this based on a unit flag
if (((gnUnitID == 1) || (gnUnitID == 21)) && (ganUnitTopperIDs[0] != 0)) {
	gbNeedToppers = true;
}
if (ganSizeStyleIDs.length == 1) {
	gbNeedStyle = false;
}
<%
End If
%>

var i;

i = <%=UBound(ganEmpIDs) + 1%>;
var ganEmpIDs = new Array(i);
var ganEmployeeIDs = new Array(i);
var gasCardIDs = new Array(i);
<%
For i = 0 To UBound(ganEmpIDs)
%>
	ganEmpIDs[<%=i%>] = <%=ganEmpIDs(i)%>;
	ganEmployeeIDs[<%=i%>] = <%=ganEmployeeIDs(i)%>;
	gasCardIDs[<%=i%>] = "<%=gasCardIDs(i)%>";
<%
Next
%>

function setSize(pnUnitSizeID) {
	var i, j, s, loSizeDiv, lbFound, lbSizeStyleFound;
	
	gnSizeID = pnUnitSizeID;
	gsSizeDescription = "";
	for (i = 0; i < ganUnitSizeIDs.length; i++) {
		if (ganUnitSizeIDs[i] == pnUnitSizeID) {
			gsSizeDescription = gasUnitSizeShortDescriptions[i];
			i = ganUnitSizeIDs.length;
		}
	}
	
	s = gsSizeDescription + " " + gsUnitDescription;
	loSizeDiv = ie4? eval("document.all.editunit") : document.getElementById("editunit");
	loSizeDiv.innerHTML = s;
	loSizeDiv.style.height = "15px";
	loSizeDiv.style.visibility = "visible";
	
	if (ganSizeStyleIDs[0] != 0) {
		lbFound = false;
		for (i = 0; i < ganSizeStyleIDs.length; i++) {
			if (ganSizeStyleIDs[i] == gnStyleID) {
				for (j = 0; j < ganSizeStyleSizeIDs[i].length; j++) {
					if (ganSizeStyleSizeIDs[i][j] == pnUnitSizeID) {
						lbFound = true;
						j = ganSizeStyleSizeIDs[i].length;
					}
				}
				
				i = ganSizeStyleIDs.length;
			}
		}
		if (!lbFound) {
			lbSizeStyleFound = false;
			for (i = 0; i < ganSizeStyleIDs.length; i++) {
				for (j = 0; j < ganSizeStyleSizeIDs[i].length; j++) {
					if (ganSizeStyleSizeIDs[i][j] == pnUnitSizeID) {
						lbSizeStyleFound = true;
						
						setStyle(ganSizeStyleIDs[i]);
						
						j = ganSizeStyleSizeIDs[i].length;
						i = ganSizeStyleIDs.length - 1;
					}
				}
			}
			if (!lbSizeStyleFound) {
				setStyle(0);
			}
		}
		
		if (!gbManualStyle) {
			// if size only has one style do not prompt for specialty style
			gbManualStyle = true;
			lbFound = false;
			for (i = 0; i < ganSizeStyleIDs.length; i++) {
				for (j = 0; j < ganSizeStyleSizeIDs[i].length; j++) {
					if (ganSizeStyleSizeIDs[i][j] == pnUnitSizeID) {
						if (lbFound) {
							gbManualStyle = false;
							j = ganSizeStyleSizeIDs[i].length;
							i = ganSizeStyleIDs.length - 1;
						}
						else {
							lbFound = true;
						}
					}
				}
			}
		}
		
		for (i = 0; i < ganUnitSizeIDs.length; i++) {
			loSizeDiv = ie4? eval("document.all.sizestylediv" + ganUnitSizeIDs[i].toString()) : document.getElementById("sizestylediv" + ganUnitSizeIDs[i].toString());
			if (ganUnitSizeIDs[i] == pnUnitSizeID) {
				loSizeDiv.style.visibility = "visible";
			}
			else {
				loSizeDiv.style.visibility = "hidden";
			}
		}
	}
	
	resetIncludedSides();
	
	recalculatePrice();
	resetRedirect();
}

function setStyle(pnSizeStyleIDs) {
	var i, loStyleDiv, lsSpecialMessage;
	var lbNeedSpecialMessage = false;
	
	gnStyleID = pnSizeStyleIDs;
	gsStyleDescription = "";
	for (i = 0; i < ganSizeStyleIDs.length; i++) {
		if (ganSizeStyleIDs[i] == pnSizeStyleIDs) {
			gsStyleDescription = gasSizeStyleShortDescriptions[i];
			if (gasSizeStyleSpecialMessage[i].length > 0) {
				lsSpecialMessage = gasSizeStyleSpecialMessage[i];
				lbNeedSpecialMessage = true;
			}
			i = ganSizeStyleIDs.length;
		}
	}
	
	loStyleDiv = ie4? eval("document.all.editstyle") : document.getElementById("editstyle");
	loStyleDiv.innerHTML = gsStyleDescription;
	if (pnSizeStyleIDs == 0) {
		loStyleDiv.style.height = "0px";
		loStyleDiv.style.visibility = "hidden";
	}
	else {
		loStyleDiv.style.height = "15px";
		loStyleDiv.style.visibility = "visible";
	}
	
	recalculatePrice();
	resetRedirect();
	
	if (lbNeedSpecialMessage) {
		gotoSpecialMessage(lsSpecialMessage);
	}
	else {
		gotoUnitEditor();
	}
}

function setSauce(pnSauceID) {
	var i, s, loSauceDiv;
	
	switch (gnHalfID) {
		case 0:
			gnHalf1SauceID = pnSauceID;
			gnHalf2SauceID = pnSauceID;
			gsHalf1SauceDescription = "";
			gsHalf2SauceDescription = "";
			for (i = 0; i < ganSauceIDs.length; i++) {
				if (ganSauceIDs[i] == pnSauceID) {
					gsHalf1SauceDescription = gasSauceShortDescriptions[i];
					gsHalf2SauceDescription = gasSauceShortDescriptions[i];
					i = ganSauceIDs.length;
				}
			}
			break;
		case 1:
			gnHalf1SauceID = pnSauceID;
			gsHalf1SauceDescription = "";
			for (i = 0; i < ganSauceIDs.length; i++) {
				if (ganSauceIDs[i] == pnSauceID) {
					gsHalf1SauceDescription = gasSauceShortDescriptions[i];
					i = ganSauceIDs.length;
				}
			}
			break;
		case 2:
			gnHalf2SauceID = pnSauceID;
			gsHalf2SauceDescription = "";
			for (i = 0; i < ganSauceIDs.length; i++) {
				if (ganSauceIDs[i] == pnSauceID) {
					gsHalf2SauceDescription = gasSauceShortDescriptions[i];
					i = ganSauceIDs.length;
				}
			}
			break;
	}
	
	loSauceDiv = ie4? eval("document.all.edithalf1sauce") : document.getElementById("edithalf1sauce");
	if (gnHalf1SauceID == 0) {
		loSauceDiv.innerHTML = "";
		loSauceDiv.style.height = "0px";
		loSauceDiv.style.visibility = "hidden";
	}
	else {
		if ((gnHalf1SauceID == gnHalf2SauceID) && (gnHalf1SauceModifierID == gnHalf2SauceModifierID)) {
			s = "Whole Sauce: " + gsHalf1SauceDescription + " " + gsHalf1SauceModifierDescription ;
		}
		else {
			s = "1st Half Sauce: " + gsHalf1SauceDescription + " " + gsHalf1SauceModifierDescription ;
		}
		loSauceDiv.innerHTML = s;
		loSauceDiv.style.height = "15px";
		loSauceDiv.style.visibility = "visible";
	}
	
	loSauceDiv = ie4? eval("document.all.edithalf2sauce") : document.getElementById("edithalf2sauce");
	if ((gnHalf2SauceID > 0) && ((gnHalf1SauceID != gnHalf2SauceID) || (gnHalf1SauceModifierID != gnHalf2SauceModifierID))) {
		s = "2nd Half Sauce: " + gsHalf2SauceDescription + " " + gsHalf2SauceModifierDescription ;
		loSauceDiv.innerHTML = s;
		loSauceDiv.style.height = "15px";
		loSauceDiv.style.visibility = "visible";
	}
	else {
		loSauceDiv.innerHTML = "";
		loSauceDiv.style.height = "0px";
		loSauceDiv.style.visibility = "hidden";
	}
	
	recalculatePrice();
	resetRedirect();
}

function setSauceModifier(pnSauceModifierID) {
	var i, s, loSauceDiv;
	
	if (gbDeleteMode) {
		switch (gnHalfID) {
			case 0:
				gnHalf1SauceModifierID = 0;
				gnHalf2SauceModifierID = 0;
				gsHalf1SauceModifierDescription = "";
				gsHalf2SauceModifierDescription = "";
				break;
			case 1:
				gnHalf1SauceModifierID = 0;
				gsHalf1SauceModifierDescription = "";
				break;
			case 2:
				gnHalf2SauceModifierID = 0;
				gsHalf2SauceModifierDescription = "";
				break;
		}
	}
	else {
		switch (gnHalfID) {
			case 0:
				gnHalf1SauceModifierID = pnSauceModifierID;
				gnHalf2SauceModifierID = pnSauceModifierID;
				gsHalf1SauceModifierDescription = "";
				gsHalf2SauceModifierDescription = "";
				for (i = 0; i < ganSauceModifierIDs.length; i++) {
					if (ganSauceModifierIDs[i] == pnSauceModifierID) {
						gsHalf1SauceModifierDescription = gasSauceModifierShortDescriptions[i];
						gsHalf2SauceModifierDescription = gasSauceModifierShortDescriptions[i];
						i = ganSauceModifierIDs.length;
					}
				}
				break;
			case 1:
				gnHalf1SauceModifierID = pnSauceModifierID;
				gsHalf1SauceModifierDescription = "";
				for (i = 0; i < ganSauceModifierIDs.length; i++) {
					if (ganSauceModifierIDs[i] == pnSauceModifierID) {
						gsHalf1SauceModifierDescription = gasSauceModifierShortDescriptions[i];
						i = ganSauceModifierIDs.length;
					}
				}
				break;
			case 2:
				gnHalf2SauceModifierID = pnSauceModifierID;
				gsHalf2SauceModifierDescription = "";
				for (i = 0; i < ganSauceModifierIDs.length; i++) {
					if (ganSauceModifierIDs[i] == pnSauceModifierID) {
						gsHalf2SauceModifierDescription = gasSauceModifierShortDescriptions[i];
						i = ganSauceModifierIDs.length;
					}
				}
				break;
		}
	}
	
	loSauceDiv = ie4? eval("document.all.edithalf1sauce") : document.getElementById("edithalf1sauce");
	if (gnHalf1SauceID == 0) {
		loSauceDiv.innerHTML = "";
		loSauceDiv.style.height = "0px";
		loSauceDiv.style.visibility = "hidden";
	}
	else {
		if ((gnHalf1SauceID == gnHalf2SauceID) && (gnHalf1SauceModifierID == gnHalf2SauceModifierID)) {
			s = "Whole Sauce: " + gsHalf1SauceDescription + " " + gsHalf1SauceModifierDescription ;
		}
		else {
			s = "1st Half Sauce: " + gsHalf1SauceDescription + " " + gsHalf1SauceModifierDescription ;
		}
		loSauceDiv.innerHTML = s;
		loSauceDiv.style.height = "15px";
		loSauceDiv.style.visibility = "visible";
	}
	
	loSauceDiv = ie4? eval("document.all.edithalf2sauce") : document.getElementById("edithalf2sauce");
	if ((gnHalf2SauceID > 0) && ((gnHalf1SauceID != gnHalf2SauceID) || (gnHalf1SauceModifierID != gnHalf2SauceModifierID))) {
		s = "2nd Half Sauce: " + gsHalf2SauceDescription + " " + gsHalf2SauceModifierDescription ;
		loSauceDiv.innerHTML = s;
		loSauceDiv.style.height = "15px";
		loSauceDiv.style.visibility = "visible";
	}
	else {
		loSauceDiv.innerHTML = "";
		loSauceDiv.style.height = "0px";
		loSauceDiv.style.visibility = "hidden";
	}
	
	recalculatePrice();
	resetRedirect();
}

function toggleDelete() {
	var loDeleteButton;
	
	loDeleteButton = ie4? eval("document.all.deletebutton") : document.getElementById("deletebutton");
	gbDeleteMode = !gbDeleteMode;
	if (gbDeleteMode) {
		loDeleteButton.style.color = "#FFFFFF";
	}
	else {
		loDeleteButton.style.color = "#000000";
	}
	
	resetRedirect();
}

function toggleWhole() {
	var loWholeButton;
	
	loWholeButton = ie4? eval("document.all.wholebutton") : document.getElementById("wholebutton");
	switch (gnHalfID) {
		case 0:
			loWholeButton.innerHTML = "Half 1";
			gnHalfID = 1;
			break;
		case 1:
			loWholeButton.innerHTML = "Half 2";
			gnHalfID = 2;
			break;
		case 2:
			loWholeButton.innerHTML = "On Side";
			gnHalfID = 3;
			break;
		case 3:
			loWholeButton.innerHTML = "Whole";
			gnHalfID = 0;
			break;
	}
	
//	toggleItems();
	resetRedirect();
}

function resetWholeDelete() {
	var loWholeButton, loDeleteButton;
	
	loWholeButton = ie4? eval("document.all.wholebutton") : document.getElementById("wholebutton");
	loWholeButton.innerHTML = "Whole";
	gnHalfID = 0;
	
	loDeleteButton = ie4? eval("document.all.deletebutton") : document.getElementById("deletebutton");
	gbDeleteMode = false;
	loDeleteButton.style.color = "#000000";
	
//	toggleItems();
	resetRedirect();
}

function toggleItems() {
	var i, loItemButton, loSpecialtyButton, loToppersButton, loSidesButton, loDiv;
	
	loItemButton = ie4? eval("document.all.itemsbutton") : document.getElementById("itemsbutton");
	loItemButton.style.color = "#FFFFFF";
	loSpecialtyButton = ie4? eval("document.all.specialtybutton") : document.getElementById("specialtybutton");
	loSpecialtyButton.style.color = "#000000";
	loToppersButton = ie4? eval("document.all.toppersbutton") : document.getElementById("toppersbutton");
	loToppersButton.style.color = "#000000";
	loSidesButton = ie4? eval("document.all.sidesbutton") : document.getElementById("sidesbutton");
	loSidesButton.style.color = "#000000";
	
	if (ganUnitItemIDs[0] > 0) {
		for (i = 0; i < ganUnitItemIDs.length; i++) {
			if ((i % 20) == 0) {
				loDiv = ie4? eval("document.all.itemdiv" + (i / 20).toString()) : document.getElementById("itemdiv" + (i / 20).toString());
				if (i == 0) {
					loDiv.style.visibility = "visible";
				}
				else {
					loDiv.style.visibility = "hidden";
				}
			}
		}
	}
	if (ganUnitSpecialtyIDs[0] > 0) {
		for (i = 0; i < ganUnitSpecialtyIDs.length; i++) {
			if ((i % 20) == 0) {
				loDiv = ie4? eval("document.all.specialtydiv" + (i / 20).toString()) : document.getElementById("specialtydiv" + (i / 20).toString());
				loDiv.style.visibility = "hidden";
			}
		}
	}
	if (ganUnitTopperIDs[0] > 0) {
		for (i = 0; i < ganUnitTopperIDs.length; i++) {
			if ((i % 20) == 0) {
				loDiv = ie4? eval("document.all.topperdiv" + (i / 20).toString()) : document.getElementById("topperdiv" + (i / 20).toString());
				loDiv.style.visibility = "hidden";
			}
		}
	}
	if (ganUnitSideIDs[0] > 0) {
		for (i = 0; i < ganUnitSideIDs.length; i++) {
			if ((i % 20) == 0) {
				loDiv = ie4? eval("document.all.sidediv" + (i / 20).toString()) : document.getElementById("sidediv" + (i / 20).toString());
				loDiv.style.visibility = "hidden";
			}
		}
	}
	
	resetRedirect();
}

function toggleSpecialty() {
	var li, oItemButton, loSpecialtyButton, loToppersButton, loSidesButton, loDiv;
	
	resetWholeDelete();
	
	loItemButton = ie4? eval("document.all.itemsbutton") : document.getElementById("itemsbutton");
	loItemButton.style.color = "#000000";
	loSpecialtyButton = ie4? eval("document.all.specialtybutton") : document.getElementById("specialtybutton");
	loSpecialtyButton.style.color = "#FFFFFF";
	loToppersButton = ie4? eval("document.all.toppersbutton") : document.getElementById("toppersbutton");
	loToppersButton.style.color = "#000000";
	loSidesButton = ie4? eval("document.all.sidesbutton") : document.getElementById("sidesbutton");
	loSidesButton.style.color = "#000000";
	
	if (ganUnitItemIDs[0] > 0) {
		for (i = 0; i < ganUnitItemIDs.length; i++) {
			if ((i % 20) == 0) {
				loDiv = ie4? eval("document.all.itemdiv" + (i / 20).toString()) : document.getElementById("itemdiv" + (i / 20).toString());
				loDiv.style.visibility = "hidden";
			}
		}
	}
	if (ganUnitSpecialtyIDs[0] > 0) {
		for (i = 0; i < ganUnitSpecialtyIDs.length; i++) {
			if ((i % 20) == 0) {
				loDiv = ie4? eval("document.all.specialtydiv" + (i / 20).toString()) : document.getElementById("specialtydiv" + (i / 20).toString());
				if (i == 0) {
					loDiv.style.visibility = "visible";
				}
				else {
					loDiv.style.visibility = "hidden";
				}
			}
		}
	}
	if (ganUnitTopperIDs[0] > 0) {
		for (i = 0; i < ganUnitTopperIDs.length; i++) {
			if ((i % 20) == 0) {
				loDiv = ie4? eval("document.all.topperdiv" + (i / 20).toString()) : document.getElementById("topperdiv" + (i / 20).toString());
				loDiv.style.visibility = "hidden";
			}
		}
	}
	if (ganUnitSideIDs[0] > 0) {
		for (i = 0; i < ganUnitSideIDs.length; i++) {
			if ((i % 20) == 0) {
				loDiv = ie4? eval("document.all.sidediv" + (i / 20).toString()) : document.getElementById("sidediv" + (i / 20).toString());
				loDiv.style.visibility = "hidden";
			}
		}
	}
	
	resetRedirect();
}

function toggleToppers() {
	var i, loItemButton, loSpecialtyButton, loToppersButton, loSidesButton, loDiv;
	
	resetWholeDelete();
	
	loItemButton = ie4? eval("document.all.itemsbutton") : document.getElementById("itemsbutton");
	loItemButton.style.color = "#000000";
	loSpecialtyButton = ie4? eval("document.all.specialtybutton") : document.getElementById("specialtybutton");
	loSpecialtyButton.style.color = "#000000";
	loToppersButton = ie4? eval("document.all.toppersbutton") : document.getElementById("toppersbutton");
	loToppersButton.style.color = "#FFFFFF";
	loSidesButton = ie4? eval("document.all.sidesbutton") : document.getElementById("sidesbutton");
	loSidesButton.style.color = "#000000";
	
	if (ganUnitItemIDs[0] > 0) {
		for (i = 0; i < ganUnitItemIDs.length; i++) {
			if ((i % 20) == 0) {
				loDiv = ie4? eval("document.all.itemdiv" + (i / 20).toString()) : document.getElementById("itemdiv" + (i / 20).toString());
				loDiv.style.visibility = "hidden";
			}
		}
	}
	if (ganUnitSpecialtyIDs[0] > 0) {
		for (i = 0; i < ganUnitSpecialtyIDs.length; i++) {
			if ((i % 20) == 0) {
				loDiv = ie4? eval("document.all.specialtydiv" + (i / 20).toString()) : document.getElementById("specialtydiv" + (i / 20).toString());
				loDiv.style.visibility = "hidden";
			}
		}
	}
	if (ganUnitTopperIDs[0] > 0) {
		for (i = 0; i < ganUnitTopperIDs.length; i++) {
			if ((i % 20) == 0) {
				loDiv = ie4? eval("document.all.topperdiv" + (i / 20).toString()) : document.getElementById("topperdiv" + (i / 20).toString());
				if (i == 0) {
					loDiv.style.visibility = "visible";
				}
				else {
					loDiv.style.visibility = "hidden";
				}
			}
		}
	}
	if (ganUnitSideIDs[0] > 0) {
		for (i = 0; i < ganUnitSideIDs.length; i++) {
			if ((i % 20) == 0) {
				loDiv = ie4? eval("document.all.sidediv" + (i / 20).toString()) : document.getElementById("sidediv" + (i / 20).toString());
				loDiv.style.visibility = "hidden";
			}
		}
	}
	
	resetRedirect();
}

function toggleSides() {
	var i, loItemButton, loSpecialtyButton, loToppersButton, loSidesButton, loDiv;
	
	loItemButton = ie4? eval("document.all.itemsbutton") : document.getElementById("itemsbutton");
	loItemButton.style.color = "#000000";
	loSpecialtyButton = ie4? eval("document.all.specialtybutton") : document.getElementById("specialtybutton");
	loSpecialtyButton.style.color = "#000000";
	loToppersButton = ie4? eval("document.all.toppersbutton") : document.getElementById("toppersbutton");
	loToppersButton.style.color = "#000000";
	loSidesButton = ie4? eval("document.all.sidesbutton") : document.getElementById("sidesbutton");
	loSidesButton.style.color = "#FFFFFF";
	
	if (ganUnitItemIDs[0] > 0) {
		for (i = 0; i < ganUnitItemIDs.length; i++) {
			if ((i % 20) == 0) {
				loDiv = ie4? eval("document.all.itemdiv" + (i / 20).toString()) : document.getElementById("itemdiv" + (i / 20).toString());
				loDiv.style.visibility = "hidden";
			}
		}
	}
	if (ganUnitSpecialtyIDs[0] > 0) {
		for (i = 0; i < ganUnitSpecialtyIDs.length; i++) {
			if ((i % 20) == 0) {
				loDiv = ie4? eval("document.all.specialtydiv" + (i / 20).toString()) : document.getElementById("specialtydiv" + (i / 20).toString());
				loDiv.style.visibility = "hidden";
			}
		}
	}
	if (ganUnitTopperIDs[0] > 0) {
		for (i = 0; i < ganUnitTopperIDs.length; i++) {
			if ((i % 20) == 0) {
				loDiv = ie4? eval("document.all.topperdiv" + (i / 20).toString()) : document.getElementById("topperdiv" + (i / 20).toString());
				loDiv.style.visibility = "hidden";
			}
		}
	}
	if (ganUnitSideIDs[0] > 0) {
		for (i = 0; i < ganUnitSideIDs.length; i++) {
			if ((i % 20) == 0) {
				loDiv = ie4? eval("document.all.sidediv" + (i / 20).toString()) : document.getElementById("sidediv" + (i / 20).toString());
				if (i == 0) {
					loDiv.style.visibility = "visible";
				}
				else {
					loDiv.style.visibility = "hidden";
				}
			}
		}
	}
	
	resetRedirect();
}

function addItem(pnItemID) {
	var i, s, loDiv, loDiv2, lbOK, j, loButton, lbFound;
	
	if (gbDeleteMode) {
		for (i = 0; i < ganItemIDs.length; i++) {
			if ((ganItemIDs[i] == pnItemID) && (ganHalfIDs[i] == gnHalfID)) {
				for (j = (i + 1); j < ganItemIDs.length; j++) {
					ganItemIDs[(j - 1)] = ganItemIDs[j];
					ganHalfIDs[(j - 1)] = ganHalfIDs[j];
					gasItemDescriptions[(j - 1)] = gasItemDescriptions[j];
					gasItemShortDescriptions[(j - 1)] = gasItemShortDescriptions[j];
					
					loDiv = ie4? eval("document.all.edititem" + (j - 1).toString()) : document.getElementById("edititem" + (j- 1).toString());
					loDiv2 = ie4? eval("document.all.edititem" + j.toString()) : document.getElementById("edititem" + j.toString());
					
					loDiv.innerHTML = loDiv2.innerHTML;
				}
				loDiv = ie4? eval("document.all.edititem" + (ganItemIDs.length - 1).toString()) : document.getElementById("edititem" + (ganItemIDs.length - 1).toString());
				loDiv.innerHTML = "";
				loDiv.style.height = "0px";
				loDiv.style.visibility = "hidden";
				
				i = ganItemIDs.length;
				ganItemIDs.length = ganItemIDs.length - 1;
				ganHalfIDs.length = ganHalfIDs.length - 1;
				gasItemDescriptions.length = gasItemDescriptions.length - 1;
				gasItemShortDescriptions.length = gasItemShortDescriptions.length - 1;
			}
		}
		
		lbFound = false;
		for (i = 0; i < ganItemIDs.length; i++) {
			if (ganItemIDs[i] == pnItemID) {
				lbFound = true;
				i = ganItemIDs.length;
			}
		}
		if (!lbFound) {
			loButton = ie4? eval("document.all.item" + pnItemID.toString()) : document.getElementById("item-" + pnItemID.toString());
			loButton.style.color = "#000000";
		}
	}
	else {
		if (ganItemIDs.length < <%=gnMaxItemsPerUnit%>) {
			for (i = 0; i < ganUnitItemIDs.length; i++) {
				if (ganUnitItemIDs[i] == pnItemID) {
					lbOK = true;
					
					if (gabUnitItemIsBaseCheeses[i]) {
						for (j = 0; j < ganItemIDs.length; j++) {
							if (ganItemIDs[j] == pnItemID) {
								lbOK = false;
							}
						}
					}
					
					if (lbOK) {
						if (!(ganItemIDs.length == 1 && ganItemIDs[0] == 0)) {
							ganItemIDs.length = ganItemIDs.length + 1;
							ganHalfIDs.length = ganHalfIDs.length + 1;
							gasItemDescriptions.length = gasItemDescriptions.length + 1;
							gasItemShortDescriptions.length = gasItemShortDescriptions.length + 1;
						}
						ganItemIDs[(ganItemIDs.length - 1)] = pnItemID;
						ganHalfIDs[(ganHalfIDs.length - 1)] = gnHalfID;
						gasItemDescriptions[(gasItemDescriptions.length - 1)] = gasUnitItemDescriptions[i];
						gasItemShortDescriptions[(gasItemShortDescriptions.length - 1)] = gasUnitItemShortDescriptions[i];
						
						loDiv = ie4? eval("document.all.edititem" + (ganItemIDs.length - 1).toString()) : document.getElementById("edititem" + (ganItemIDs.length - 1).toString());
						switch (gnHalfID) {
							case 0:
								s = "Whole Item: ";
								break;
							case 1:
								s = "1st Half Item: ";
								break;
							case 2:
								s = "2nd Half Item: ";
								break;
							case 3:
								s = "On Side Item: ";
								break;
						}
						s = s + gasItemShortDescriptions[(ganItemIDs.length - 1)];
						loDiv.innerHTML = s;
						loDiv.style.height = "15px";
						loDiv.style.visibility = "visible";
						
						loButton = ie4? eval("document.all.item" + pnItemID.toString()) : document.getElementById("item-" + pnItemID.toString());
						loButton.style.color = "#FFFFFF";
					}
					
					i = ganUnitItemIDs.length
				}
			}
		}
	}
	
	recalculatePrice();
	resetRedirect();
}

function addTopper(pnTopperID) {
	var i, s, loDiv, loDiv2;
	
	gbNeedToppers = false;
	
	if (gbDeleteMode) {
		for (i = 0; i < ganTopperIDs.length; i++) {
			if ((ganTopperIDs[i] == pnTopperID) && (ganTopperHalfIDs[i] == gnHalfID)) {
				for (j = (i + 1); j < ganTopperIDs.length; j++) {
					ganTopperIDs[(j - 1)] = ganTopperIDs[j];
					ganTopperHalfIDs[(j - 1)] = ganTopperHalfIDs[j];
					gasTopperDescriptions[(j - 1)] = gasTopperDescriptions[j];
					gasTopperShortDescriptions[(j - 1)] = gasTopperShortDescriptions[j];
					
					loDiv = ie4? eval("document.all.edittopper" + (j - 1).toString()) : document.getElementById("edittopper" + (j- 1).toString());
					loDiv2 = ie4? eval("document.all.edittopper" + j.toString()) : document.getElementById("edittopper" + j.toString());
					
					loDiv.innerHTML = loDiv2.innerHTML;
				}
				loDiv = ie4? eval("document.all.edittopper" + (ganTopperIDs.length - 1).toString()) : document.getElementById("edittopper" + (ganTopperIDs.length - 1).toString());
				loDiv.innerHTML = "";
				loDiv.style.height = "0px";
				loDiv.style.visibility = "hidden";
				
				i = ganTopperIDs.length;
				ganTopperIDs.length = ganTopperIDs.length - 1;
				ganTopperHalfIDs.length = ganTopperHalfIDs.length - 1;
				gasTopperDescriptions.length = gasTopperDescriptions.length - 1;
				gasTopperShortDescriptions.length = gasTopperShortDescriptions.length - 1;
			}
		}
	}
	else {
		if (ganTopperIDs.length < <%=gnMaxToppersPerUnit%>) {
			if (!(ganTopperIDs.length == 1 && ganTopperIDs[0] == 0)) {
				ganTopperIDs.length = ganTopperIDs.length + 1;
				ganTopperHalfIDs.length = ganTopperHalfIDs.length + 1;
				gasTopperDescriptions.length = gasTopperDescriptions.length + 1;
				gasTopperShortDescriptions.length = gasTopperShortDescriptions.length + 1;
			}
			ganTopperIDs[(ganTopperIDs.length - 1)] = 0;
			ganTopperHalfIDs[(ganTopperHalfIDs.length - 1)] = 0;
			gasTopperDescriptions[(gasTopperDescriptions.length - 1)] = "";
			gasTopperShortDescriptions[(gasTopperShortDescriptions.length - 1)] = "";
			for (i = 0; i < ganUnitTopperIDs.length; i++) {
				if (ganUnitTopperIDs[i] == pnTopperID) {
					ganTopperIDs[(ganTopperIDs.length - 1)] = pnTopperID;
					ganTopperHalfIDs[(ganTopperIDs.length - 1)] = gnHalfID;
					gasTopperDescriptions[(gasTopperDescriptions.length - 1)] = gasUnitTopperDescriptions[i];
					gasTopperShortDescriptions[(gasTopperShortDescriptions.length - 1)] = gasUnitTopperShortDescriptions[i];
					
					i = ganUnitTopperIDs.length
				}
			}
			
			loDiv = ie4? eval("document.all.edittopper" + (ganTopperIDs.length - 1).toString()) : document.getElementById("edittopper" + (ganTopperIDs.length - 1).toString());
			switch (gnHalfID) {
				case 0:
					s = "Whole Topper: ";
					break;
				case 1:
					s = "1st Half Topper: ";
					break;
				case 2:
					s = "2nd Half Topper: ";
					break;
				case 3:
					s = "On Side Topper: ";
					break;
			}
			s = s + gasTopperShortDescriptions[(ganTopperIDs.length - 1)];
			loDiv.innerHTML = s;
			loDiv.style.height = "15px";
			loDiv.style.visibility = "visible";
		}
	}

// Don't recalc price for toppers	
//	recalculatePrice();
	resetRedirect();
}

function addSide(pnSideID) {
	var i, loDiv;
	
	if (gbDeleteMode) {
		for (i = 0; i < ganAddSideIDs.length; i++) {
			if (ganAddSideIDs[i] == pnSideID) {
				for (j = (i + 1); j < ganAddSideIDs.length; j++) {
					ganAddSideIDs[(j - 1)] = ganAddSideIDs[j];
					gasAddSideDescriptions[(j - 1)] = gasAddSideDescriptions[j];
					gasAddSideShortDescriptions[(j - 1)] = gasAddSideShortDescriptions[j];
					
					loDiv = ie4? eval("document.all.editaddside" + (j - 1).toString()) : document.getElementById("editaddside" + (j- 1).toString());
					loDiv2 = ie4? eval("document.all.editaddside" + j.toString()) : document.getElementById("editaddside" + j.toString());
					
					loDiv.innerHTML = loDiv2.innerHTML;
				}
				loDiv = ie4? eval("document.all.editaddside" + (ganAddSideIDs.length - 1).toString()) : document.getElementById("editaddside" + (ganAddSideIDs.length - 1).toString());
				loDiv.innerHTML = "";
				loDiv.style.height = "0px";
				loDiv.style.visibility = "hidden";
				
				i = ganAddSideIDs.length;
				ganAddSideIDs.length = ganAddSideIDs.length - 1;
				gasAddSideDescriptions.length = gasAddSideDescriptions.length - 1;
				gasAddSideShortDescriptions.length = gasAddSideShortDescriptions.length - 1;
			}
		}
	}
	else {
		if (ganAddSideIDs.length < <%=gnMaxAddSidesPerUnit%>) {
			if (!(ganAddSideIDs.length == 1 && ganAddSideIDs[0] == 0)) {
				ganAddSideIDs.length = ganAddSideIDs.length + 1;
				gasAddSideDescriptions.length = gasAddSideDescriptions.length + 1;
				gasAddSideShortDescriptions.length = gasAddSideShortDescriptions.length + 1;
			}
			ganAddSideIDs[(ganAddSideIDs.length - 1)] = 0;
			gasAddSideDescriptions[(gasAddSideDescriptions.length - 1)] = "";
			gasAddSideShortDescriptions[(gasAddSideShortDescriptions.length - 1)] = "";
			for (i = 0; i < ganUnitSideIDs.length; i++) {
				if (ganUnitSideIDs[i] == pnSideID) {
					ganAddSideIDs[(ganAddSideIDs.length - 1)] = pnSideID;
					gasAddSideDescriptions[(gasAddSideDescriptions.length - 1)] = gasUnitSideDescriptions[i];
					gasAddSideShortDescriptions[(gasAddSideShortDescriptions.length - 1)] = gasUnitSideShortDescriptions[i];
					
					i = ganUnitSideIDs.length
				}
			}
			
			loDiv = ie4? eval("document.all.editaddside" + (ganAddSideIDs.length - 1).toString()) : document.getElementById("editaddside" + (ganAddSideIDs.length - 1).toString());
			loDiv.innerHTML = "Add Side: " + gasAddSideShortDescriptions[(ganAddSideIDs.length - 1)];
			loDiv.style.height = "15px";
			loDiv.style.visibility = "visible";
		}
	}
	
	recalculatePrice();
	resetRedirect();
}

function setSpecialty(pnSpecialtyID) {
	var i, j, k, loSpecialtyDiv, loDiv, lbOK, lbNoBaseCheese, lnBaseCheeseItemID, lsSpecialty, lnSpecialStyleID, lsSpecialStyle;
	var lbNeedSpecialStyle = false;
	
	if (gnHalfID != 3) {
		if (gbDeleteMode) {
			// TODO: Remove specialty and reset included sides based on size
		}
		else {
			lbNoBaseCheese = false;
			lnBaseCheeseItemID = 0;
			
			for (i = 0; i < ganUnitSpecialtyIDs.length; i++) {
				if (ganUnitSpecialtyIDs[i] == pnSpecialtyID) {
					lbNoBaseCheese = gabSpecialtyNoBaseCheese[i];
					
					i = ganUnitSpecialtyIDs.length;
				}
			}
			for (i = 0; i < ganUnitItemIDs.length; i++) {
				if (gabUnitItemIsBaseCheeses[i]) {
					lnBaseCheeseItemID = ganUnitItemIDs[i];
					
					i = ganUnitItemIDs.length;
				}
			}
			
			if (ganItemIDs[0] != 0) {
				for (i = 0; i < ganItemIDs.length; i++) {
					lbOK = true;
					
//					for (j = 0; j < ganUnitItemIDs.length; j++) {
//						if (ganItemIDs[i] == ganUnitItemIDs[j]) {
//							if (gabUnitItemIsBaseCheeses[j]) {
//								lbOK = false;
//								
//								j = ganUnitItemIDs.length;
//							}
//						}
//					}
					
					if (lbOK) {
						if (((gnHalfID == 0) && (ganHalfIDs[i] != 3)) || ((gnHalfID != 0) && (gnHalfID != 3) && (ganHalfIDs[i] == 0)) || (ganHalfIDs[i] == gnHalfID)) {
							if ((gnHalfID != 0) && (gnHalfID != 3) && (ganHalfIDs[i] == 0)) {
								loDiv = ie4? eval("document.all.edititem" + i.toString()) : document.getElementById("edititem" + i.toString());
								if (gnHalfID == 1) {
									ganHalfIDs[i] = 2;
									loDiv.innerHTML = "2nd Half Item: " + gasItemShortDescriptions[i];
								}
								else {
									ganHalfIDs[i] = 1;
									loDiv.innerHTML = "1st Half Item: " + gasItemShortDescriptions[i];
								}
								loDiv.style.height = "15px";
								loDiv.style.visibility = "visible";
							}
							else {
								for (j = (i + 1); j < ganItemIDs.length; j++) {
									ganItemIDs[(j - 1)] = ganItemIDs[j];
									ganHalfIDs[(j - 1)] = ganHalfIDs[j];
									gasItemDescriptions[(j - 1)] = gasItemDescriptions[j];
									gasItemShortDescriptions[(j - 1)] = gasItemShortDescriptions[j];
									
									loDiv = ie4? eval("document.all.edititem" + (j - 1).toString()) : document.getElementById("edititem" + (j- 1).toString());
									loDiv2 = ie4? eval("document.all.edititem" + j.toString()) : document.getElementById("edititem" + j.toString());
									
									loDiv.innerHTML = loDiv2.innerHTML;
								}
								loDiv = ie4? eval("document.all.edititem" + (ganItemIDs.length - 1).toString()) : document.getElementById("edititem" + (ganItemIDs.length - 1).toString());
								loDiv.innerHTML = "";
								loDiv.style.height = "0px";
								loDiv.style.visibility = "hidden";
								
								ganItemIDs.length = ganItemIDs.length - 1;
								ganHalfIDs.length = ganHalfIDs.length - 1;
								gasItemDescriptions.length = gasItemDescriptions.length - 1;
								gasItemShortDescriptions.length = gasItemShortDescriptions.length - 1;
								
								i--;
							}
						}
					}
				}
			}
			
			
			for (i = 0; i < ganTopperIDs.length; i++) {
				loDiv = ie4? eval("document.all.edittopper" + i.toString()) : document.getElementById("edittopper" + i.toString());
				loDiv.innerHTML = "";
				loDiv.style.height = "0px";
				loDiv.style.visibility = "hidden";
			}
			ganTopperIDs.length = 1;
			ganTopperIDs[0] = 0;
			gasTopperDescriptions.length = 1;
			gasTopperDescriptions[0] = "";
			gasTopperShortDescriptions.length = 1;
			gasTopperShortDescriptions[0] = "";
			
			gnSpecialtyID = pnSpecialtyID;
			gsSpecialtyDescription = "";
			for (i = 0; i < ganUnitSpecialtyIDs.length; i++) {
				if (ganUnitSpecialtyIDs[i] == pnSpecialtyID) {
					gsSpecialtyDescription = gasUnitSpecialtyShortDescriptions[i];
					setSauce(ganUnitSpecialtySauceID[i]);
					setSauceModifier(0);
					if (!gbManualStyle) {
						if (ganUnitSpecialtyStyleID[i] != 0) {
							if (ganUnitSpecialtyStyleID[i] != ganSizeStyleIDs[0]) {
								lbNeedSpecialStyle = true;
								lsSpecialty = gasUnitSpecialtyDescriptions[i];
								lnSpecialStyleID = ganUnitSpecialtyStyleID[i];
								for (j = 0; j < ganSizeStyleIDs.length; j++) {
									if (ganUnitSpecialtyStyleID[i] == ganSizeStyleIDs[j]) {
										lsSpecialStyle = gasSizeStyleDescriptions[j];
										j = ganSizeStyleIDs.length;
									}
								}
							}
							else {
								setStyle(ganUnitSpecialtyStyleID[i]);
							}
						}
					}
					
					if (!lbNoBaseCheese) {
						addItem(lnBaseCheeseItemID);
					}
					
					for (j = 0; j < ganUnitSpecialtyItemIDs[i].length; j++) {
						if (ganUnitSpecialtyItemIDs[i][j] > 0) {
							for (k = 0; k < ganUnitSpecialtyItemQuantity[i][j]; k++) {
								addItem(ganUnitSpecialtyItemIDs[i][j]);
							}
						}
					}
					
					i = ganUnitSpecialtyIDs.length;
				}
			}
		}
		
		resetIncludedSides();
		
		loSpecialtyDiv = ie4? eval("document.all.editspecialty") : document.getElementById("editspecialty");
		loSpecialtyDiv.innerHTML = gsSpecialtyDescription;
		loSpecialtyDiv.style.height = "15px";
		loSpecialtyDiv.style.visibility = "visible";
	}
	
	recalculatePrice();
	resetRedirect();
	
	if (lbNeedSpecialStyle) {
		gotoSpecialtyStyle(lsSpecialty, lnSpecialStyleID, lsSpecialStyle)
	}
}

function resetIncludedSides() {
	var i, j, k, l, m, n, lnPos, lnCount, lbFound;
	
	gbNeedIncludedSides = true;
	
	for (i = 0; i < ganFreeSideIDs.length; i++) {
		loDiv = ie4? eval("document.all.editfreeside" + i.toString()) : document.getElementById("editfreeside" + i.toString());
		loDiv.innerHTML = "";
		loDiv.style.height = "0px";
		loDiv.style.visibility = "hidden";
	}
	ganFreeSideIDs.length = 1;
	ganFreeSideIDs[0] = 0;
	gasFreeSideDescriptions.length = 1;
	gasFreeSideDescriptions[0] = "";
	gasFreeSideShortDescriptions.length = 1;
	gasFreeSideShortDescriptions[0] = "";
	
	lnPos = 0;
	if (gnSpecialtyID != 0) {
		for (i = 0; i < ganSideGroupSpecialtyIDs.length; i++) {
			if (ganSideGroupSpecialtyIDs[i] == gnSpecialtyID) {
				for (j = 0; j < ganSideGroupSizeIDs[i].length; j++) {
					if (ganSideGroupSizeIDs[i][j] == gnSizeID) {
						if (ganSideGroupSideGroupIDs[i][j][0] != 0) {
							gbHasSpecSides = true;
							
							for (k = 0; k < ganSideGroupSideGroupIDs[i][j].length; k++) {
								for (l = 0; l < ganSideGroupIDs.length; l++) {
									if (ganSideGroupIDs[l] == ganSideGroupSideGroupIDs[i][j][k]) {
										if (ganSideGroupSideIDs[l][0] != 0) {
											for (m = 0; m < gadSideGroupQuantity[i][j][k]; m++) {
												lnCount = 0;
												lbFound = false;
												for (n = 0; n < ganSideGroupSideIDs[l].length; n++) {
													if (gabSideGroupSideIsDefault[l][n]) {
														if (lnCount == lnPos) {
															ganFreeSideIDs.length = (lnPos + 1);
															setInclSide(lnPos, ganSideGroupSideIDs[l][n], gasSideGroupSideDescriptions[l][n], gasSideGroupSideShortDescriptions[l][n]);
															
															lnPos++;
															lbFound = true;
															
															n = ganSideGroupSideIDs[l].length;
														}
														
														lnCount++;
													}
												}
												if (!lbFound) {
													for (n = 0; n < ganSideGroupSideIDs[l].length; n++) {
														if (gabSideGroupSideIsDefault[l][n]) {
															ganFreeSideIDs.length = (lnPos + 1);
															setInclSide(lnPos, ganSideGroupSideIDs[l][n], gasSideGroupSideDescriptions[l][n], gasSideGroupSideShortDescriptions[l][n]);
															
															lnPos++;
															lbFound = true;
															
															n = ganSideGroupSideIDs[l].length;
														}
													}
												}
												if (!lbFound) {
													ganFreeSideIDs.length = (lnPos + 1);
													setInclSide(lnPos, ganSideGroupSideIDs[l][0], gasSideGroupSideDescriptions[l][0], gasSideGroupSideShortDescriptions[l][0]);
													
													lnPos++;
												}
											}
										}
										
										l = ganSideGroupIDs.length;
									}
								}
							}
						}
				
						j = ganSideGroupSizeIDs[i].length;
					}
				}
				
				i = ganSideGroupSpecialtyIDs.length;
			}
		}
	}
	
	if (!gbHasSpecSides) {
		for (j = 0; j < ganUnitGroupSizeIDs.length; j++) {
			if (ganUnitGroupSizeIDs[j] == gnSizeID) {
				if (ganUnitGroupSideGroupIDs[j][0] != 0) {
					for (k = 0; k < ganUnitGroupSideGroupIDs[j].length; k++) {
						for (l = 0; l < ganSideGroupIDs.length; l++) {
							if (ganSideGroupIDs[l] == ganUnitGroupSideGroupIDs[j][k]) {
								if (ganSideGroupSideIDs[l][0] != 0) {
									for (m = 0; m < gadUnitGroupQuantity[j][k]; m++) {
										lnCount = 0;
										lbFound = false;
										for (n = 0; n < ganSideGroupSideIDs[l].length; n++) {
											if (gabSideGroupSideIsDefault[l][n]) {
												if (lnCount == lnPos) {
													ganFreeSideIDs.length = (lnPos + 1);
													setInclSide(lnPos, ganSideGroupSideIDs[l][n], gasSideGroupSideDescriptions[l][n], gasSideGroupSideShortDescriptions[l][n]);
													
													lnPos++;
													lbFound = true;
													
													n = ganSideGroupSideIDs[l].length;
												}
												
												lnCount++;
											}
										}
										if (!lbFound) {
											for (n = 0; n < ganSideGroupSideIDs[l].length; n++) {
												if (gabSideGroupSideIsDefault[l][n]) {
													ganFreeSideIDs.length = (lnPos + 1);
													setInclSide(lnPos, ganSideGroupSideIDs[l][n], gasSideGroupSideDescriptions[l][n], gasSideGroupSideShortDescriptions[l][n]);
													
													lnPos++;
													lbFound = true;
													
													n = ganSideGroupSideIDs[l].length;
												}
											}
										}
										if (!lbFound) {
											ganFreeSideIDs.length = (lnPos + 1);
											setInclSide(lnPos, ganSideGroupSideIDs[l][0], gasSideGroupSideDescriptions[l][0], gasSideGroupSideShortDescriptions[l][0]);
											
											lnPos++;
										}
									}
								}
								
								l = ganSideGroupIDs.length;
							}
						}
					}
				}
		
				j = ganUnitGroupSizeIDs.length;
			}
		}
	}
}

function gotoUnitEditor() {
	var i, j, loDiv;
	
	if (gbNeedStyleDoneIsDone || gbIncludedSidesDoneIsDone || gbNeedTopperDoneIsDone) {
		gbNeedStyleDoneIsDone = false;
		gbIncludedSidesDoneIsDone = false;
		gbNeedTopperDoneIsDone = false;
		saveOrderLine(gbAddAnother);
	}
	else {
		loDiv = ie4? eval("document.all.deleteconfirm") : document.getElementById("deleteconfirm");
		loDiv.style.visibility = "hidden";
		
		loDiv = ie4? eval("document.all.orderlinenotes") : document.getElementById("orderlinenotes");
		loDiv.style.visibility = "hidden";
		
		loDiv = ie4? eval("document.all.quantitydiv") : document.getElementById("quantitydiv");
		loDiv.style.visibility = "hidden";
		
		loDiv = ie4? eval("document.all.pricediv") : document.getElementById("pricediv");
		loDiv.style.visibility = "hidden";
		
		for (i = 0; i < ganUnitGroupSizeIDs.length; i++) {
			loDiv = ie4? eval("document.all.inclsides" + i.toString()) : document.getElementById("inclsides" + i.toString());
			loDiv.style.visibility = "hidden";
		}
		for (i = 0; i < ganSideGroupSpecialtyIDs.length; i++) {
			for (j = 0; j < ganSideGroupSizeIDs[i].length; j++) {
				loDiv = ie4? eval("document.all.specialtysides" + i.toString() + "-" + j.toString()) : document.getElementById("specialtysides" + i.toString() + "-" + j.toString());
				loDiv.style.visibility = "hidden";
			}
		}
		
		loDiv = ie4? eval("document.all.messagediv") : document.getElementById("messagediv");
		loDiv.style.visibility = "hidden";
		
		loDiv = ie4? eval("document.all.specialtystylediv") : document.getElementById("specialtystylediv");
		loDiv.style.visibility = "hidden";
		
		loDiv = ie4? eval("document.all.needstylediv") : document.getElementById("needstylediv");
		loDiv.style.visibility = "hidden";
		
		loDiv = ie4? eval("document.all.needtoppersdiv") : document.getElementById("needtoppersdiv");
		loDiv.style.visibility = "hidden";
		
		loDiv = ie4? eval("document.all.managerdiv") : document.getElementById("managerdiv");
		loDiv.style.visibility = "hidden";
		
		loDiv = ie4? eval("document.all.mporeasondiv") : document.getElementById("mporeasondiv");
		loDiv.style.visibility = "hidden";
		
		loDiv = ie4? eval("document.all.uniteditor") : document.getElementById("uniteditor");
		loDiv.style.visibility = "visible";
		
		resetRedirect();
	}
}

function gotoDeleteConfirm() {
	var i, j, loDiv;
	
	loDiv = ie4? eval("document.all.uniteditor") : document.getElementById("uniteditor");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.orderlinenotes") : document.getElementById("orderlinenotes");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.quantitydiv") : document.getElementById("quantitydiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.pricediv") : document.getElementById("pricediv");
	loDiv.style.visibility = "hidden";
	
	for (i = 0; i < ganUnitGroupSizeIDs.length; i++) {
		loDiv = ie4? eval("document.all.inclsides" + i.toString()) : document.getElementById("inclsides" + i.toString());
		loDiv.style.visibility = "hidden";
	}
	for (i = 0; i < ganSideGroupSpecialtyIDs.length; i++) {
		for (j = 0; j < ganSideGroupSizeIDs[i].length; j++) {
			loDiv = ie4? eval("document.all.specialtysides" + i.toString() + "-" + j.toString()) : document.getElementById("specialtysides" + i.toString() + "-" + j.toString());
			loDiv.style.visibility = "hidden";
		}
	}
	
	loDiv = ie4? eval("document.all.messagediv") : document.getElementById("messagediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.specialtystylediv") : document.getElementById("specialtystylediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.needstylediv") : document.getElementById("needstylediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.needtoppersdiv") : document.getElementById("needtoppersdiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById("managerdiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.mporeasondiv") : document.getElementById("mporeasondiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.deleteconfirm") : document.getElementById("deleteconfirm");
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function gotoOrderLineNotes() {
	var i, j, loNotes, loDiv;
	
	loNotes = ie4? eval("document.all.linenotes") : document.getElementById('linenotes');
	loNotes.value = gsOrderLineNotes;
	
	loDiv = ie4? eval("document.all.uniteditor") : document.getElementById("uniteditor");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.deleteconfirm") : document.getElementById("deleteconfirm");
	loDiv.style.visibility = "hidden";
	
	for (i = 0; i < ganUnitGroupSizeIDs.length; i++) {
		loDiv = ie4? eval("document.all.inclsides" + i.toString()) : document.getElementById("inclsides" + i.toString());
		loDiv.style.visibility = "hidden";
	}
	for (i = 0; i < ganSideGroupSpecialtyIDs.length; i++) {
		for (j = 0; j < ganSideGroupSizeIDs[i].length; j++) {
			loDiv = ie4? eval("document.all.specialtysides" + i.toString() + "-" + j.toString()) : document.getElementById("specialtysides" + i.toString() + "-" + j.toString());
			loDiv.style.visibility = "hidden";
		}
	}
	
	loDiv = ie4? eval("document.all.messagediv") : document.getElementById("messagediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.specialtystylediv") : document.getElementById("specialtystylediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.needstylediv") : document.getElementById("needstylediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.needtoppersdiv") : document.getElementById("needtoppersdiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById("managerdiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.mporeasondiv") : document.getElementById("mporeasondiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.orderlinenotes") : document.getElementById("orderlinenotes");
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function gotoInclSides() {
	var i, j, lbFound, loDiv;
	
	lbFound = false;
	
	if (!gbHasSpecSides) {
		for (i = 0; i < ganUnitGroupSizeIDs.length; i++) {
			if (ganUnitGroupSizeIDs[i] == gnSizeID) {
				lbFound = true;
			}
		}
	}
	else {
		for (i = 0; i < ganSideGroupSpecialtyIDs.length; i++) {
			if (ganSideGroupSpecialtyIDs[i] == gnSpecialtyID) {
				for (j = 0; j < ganSideGroupSizeIDs[i].length; j++) {
					if (ganSideGroupSizeIDs[i][j] == gnSizeID) {
						lbFound = true;
					}
				}
			}
		}
	}
	
	if (lbFound) {
		gbNeedIncludedSides = false;
		
		loDiv = ie4? eval("document.all.uniteditor") : document.getElementById("uniteditor");
		loDiv.style.visibility = "hidden";
		
		loDiv = ie4? eval("document.all.deleteconfirm") : document.getElementById("deleteconfirm");
		loDiv.style.visibility = "hidden";
		
		loDiv = ie4? eval("document.all.orderlinenotes") : document.getElementById("orderlinenotes");
		loDiv.style.visibility = "hidden";
		
		loDiv = ie4? eval("document.all.messagediv") : document.getElementById("messagediv");
		loDiv.style.visibility = "hidden";
		
		loDiv = ie4? eval("document.all.specialtystylediv") : document.getElementById("specialtystylediv");
		loDiv.style.visibility = "hidden";
		
		loDiv = ie4? eval("document.all.needstylediv") : document.getElementById("needstylediv");
		loDiv.style.visibility = "hidden";
		
		loDiv = ie4? eval("document.all.needtoppersdiv") : document.getElementById("needtoppersdiv");
		loDiv.style.visibility = "hidden";
		
		loDiv = ie4? eval("document.all.mporeasondiv") : document.getElementById("mporeasondiv");
		loDiv.style.visibility = "hidden";
		
		loDiv = ie4? eval("document.all.managerdiv") : document.getElementById("managerdiv");
		loDiv.style.visibility = "hidden";
		
		if (!gbHasSpecSides) {
			for (i = 0; i < ganUnitGroupSizeIDs.length; i++) {
				if (ganUnitGroupSizeIDs[i] == gnSizeID) {
					loDiv = ie4? eval("document.all.inclsides" + i.toString()) : document.getElementById("inclsides" + i.toString());
					loDiv.style.visibility = "visible";
				}
				else {
					loDiv = ie4? eval("document.all.inclsides" + i.toString()) : document.getElementById("inclsides" + i.toString());
					loDiv.style.visibility = "hidden";
				}
			}
		}
		else {
			for (i = 0; i < ganSideGroupSpecialtyIDs.length; i++) {
				if (ganSideGroupSpecialtyIDs[i] == gnSpecialtyID) {
					for (j = 0; j < ganSideGroupSizeIDs[i].length; j++) {
						if (ganSideGroupSizeIDs[i][j] == gnSizeID) {
							loDiv = ie4? eval("document.all.specialtysides" + i.toString() + "-" + j.toString()) : document.getElementById("specialtysides" + i.toString() + "-" + j.toString());
							loDiv.style.visibility = "visible";
						}
						else {
							loDiv = ie4? eval("document.all.specialtysides" + i.toString() + "-" + j.toString()) : document.getElementById("specialtysides" + i.toString() + "-" + j.toString());
							loDiv.style.visibility = "hidden";
						}
					}
				}
				else {
					for (j = 0; j < ganSideGroupSizeIDs[i].length; j++) {
						loDiv = ie4? eval("document.all.specialtysides" + i.toString() + "-" + j.toString()) : document.getElementById("specialtysides" + i.toString() + "-" + j.toString());
						loDiv.style.visibility = "hidden";
					}
				}
			}
		}
	}
	
	resetRedirect();
}

function addToNotes(psDigit) {
	var loNotes, lsNotes;
	
	loNotes = ie4? eval("document.all.linenotes") : document.getElementById('linenotes');
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
	
	loNotes = ie4? eval("document.all.linenotes") : document.getElementById('linenotes');
	lsNotes = loNotes.value;
	if (lsNotes.length > 0) {
		lsNotes = lsNotes.substr(0, (lsNotes.length - 1));
		loNotes.value = lsNotes;
	}
	
	resetRedirect();
}

function clearNotes() {
	var loNotes;
	
	loNotes = ie4? eval("document.all.linenotes") : document.getElementById('linenotes');
	loNotes.value = "";
	
	resetRedirect();
}

function properNotes() {
	var loNotes, i, lsText, lbDoUpper;
	
	loNotes = ie4? eval("document.all.linenotes") : document.getElementById('linenotes');
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

function saveNotes() {
	var loNotes, loDiv;
	
	loNotes = ie4? eval("document.all.linenotes") : document.getElementById('linenotes');
	gsOrderLineNotes = loNotes.value;
	
	loDiv = ie4? eval("document.all.editnotes") : document.getElementById("editnotes");
	loDiv.innerHTML = "Notes: " + gsOrderLineNotes;
	loDiv.style.height = "45px";
	loDiv.style.visibility = "visible";
	
	gotoUnitEditor();
	resetRedirect();
}

function setInclSide(pnInclSideIdx, pnSideID, psSideDescription, psSideShortDescription) {
	var i, j, k, loDiv, lsButtonID;
	
	ganFreeSideIDs[pnInclSideIdx] = pnSideID;
	gasFreeSideDescriptions[pnInclSideIdx] = psSideDescription;
	gasFreeSideShortDescriptions[pnInclSideIdx] = psSideShortDescription
	
	loDiv = ie4? eval("document.all.editfreeside" + pnInclSideIdx.toString()) : document.getElementById("editfreeside" + pnInclSideIdx.toString());
	if (pnSideID == 0) {
		loDiv.innerHTML = "";
		loDiv.style.height = "0px";
		loDiv.style.visibility = "hidden";
	}
	else {
		loDiv.innerHTML = "Free Side: " + psSideShortDescription;
		loDiv.style.height = "15px";
		loDiv.style.visibility = "visible";
	}
	
	if (!gbHasSpecSides) {
		for (i = 0; i < ganUnitGroupSizeIDs.length; i++) {
			if (ganUnitGroupSizeIDs[i] == gnSizeID) {
				loDiv = ie4? eval("document.all.inclside" + i.toString() + "-" + pnInclSideIdx.toString()) : document.getElementById("inclside" + i.toString() + "-" + pnInclSideIdx.toString());
				lsButtonID = "inclsidebutton" + i.toString() + "-" + pnInclSideIdx.toString() + "-" + pnSideID.toString();
				for (j = 0; j < loDiv.children.length; j++) {
					if (loDiv.children[j].tagName == "BUTTON") {
						if (loDiv.children[j].id == lsButtonID) {
							loDiv.children[j].style.color = "#FFFFFF";
						}
						else {
							loDiv.children[j].style.color = "#000000";
						}
					}
				}
			}
		}
	}
	else {
		for (i = 0; i < ganSideGroupSpecialtyIDs.length; i++) {
			if (ganSideGroupSpecialtyIDs[i] == gnSpecialtyID) {
				for (j = 0; j < ganSideGroupSizeIDs[i].length; j++) {
					if (ganSideGroupSizeIDs[i][j] == gnSizeID) {
						loDiv = ie4? eval("document.all.specialtyside" + i.toString() + "-" + j.toString() + "-" + pnInclSideIdx.toString()) : document.getElementById("specialtyside" + i.toString() + "-" + j.toString() + "-" + pnInclSideIdx.toString());
						lsButtonID = "specialtysidebutton" + i.toString() + "-" + j.toString() + "-" + pnInclSideIdx.toString() + "-" + pnSideID.toString();
						for (k = 0; k < loDiv.children.length; k++) {
							if (loDiv.children[k].tagName == "BUTTON") {
								if (loDiv.children[k].id == lsButtonID) {
									loDiv.children[k].style.color = "#FFFFFF";
								}
								else {
									loDiv.children[k].style.color = "#000000";
								}
							}
						}
					}
				}
			}
		}
	}
	
	resetRedirect();
}

function gotoManager() {
	var i, j, loDiv, loText;

	loText = ie4? eval("document.all.manager") : document.getElementById('manager');
	loText.value = "";
	
	loDiv = ie4? eval("document.all.uniteditor") : document.getElementById("uniteditor");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.orderlinenotes") : document.getElementById("orderlinenotes");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.deleteconfirm") : document.getElementById("deleteconfirm");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.pricediv") : document.getElementById("pricediv");
	loDiv.style.visibility = "hidden";
	
	for (i = 0; i < ganUnitGroupSizeIDs.length; i++) {
		loDiv = ie4? eval("document.all.inclsides" + i.toString()) : document.getElementById("inclsides" + i.toString());
		loDiv.style.visibility = "hidden";
	}
	for (i = 0; i < ganSideGroupSpecialtyIDs.length; i++) {
		for (j = 0; j < ganSideGroupSizeIDs[i].length; j++) {
			loDiv = ie4? eval("document.all.specialtysides" + i.toString() + "-" + j.toString()) : document.getElementById("specialtysides" + i.toString() + "-" + j.toString());
			loDiv.style.visibility = "hidden";
		}
	}
	
	loDiv = ie4? eval("document.all.messagediv") : document.getElementById("messagediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.specialtystylediv") : document.getElementById("specialtystylediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.needstylediv") : document.getElementById("needstylediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.needtoppersdiv") : document.getElementById("needtoppersdiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.quantitydiv") : document.getElementById("quantitydiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.mporeasondiv") : document.getElementById("mporeasondiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById("managerdiv");
	loDiv.style.visibility = "visible";
	
	loText.focus();
	
	resetRedirect();
}

function addToManager(psDigit) {
	var loManager, lsManager;
	
	loManager = ie4? eval("document.all.manager") : document.getElementById('manager');
	lsManager = loManager.value;
	lsManager += psDigit;
	loManager.value = lsManager;
	
	resetRedirect();
}

function cancelManager() {
	var loText, loDiv;
	
	loText = ie4? eval("document.all.manager") : document.getElementById('manager');
	loText.value = "";
	
	gotoUnitEditor();
}

function checkManagerEnterKey() {
	if (event.keyCode == 13) {
		setManager();
	}
}

function setManager() {
	var loManager, i, lbFound;
	
	lbFound = false;
	
	loManager = ie4? eval("document.all.manager") : document.getElementById('manager');
	if (loManager.value != "") {
		for (i = 0; i < ganEmpIDs.length; i++) {
			if (gasCardIDs[i] == loManager.value) {
				lbFound = true;
				i = ganEmpIDs.length;
			}
		}
		
		if (lbFound) {
			gotoQuantity(true);
		}
		else {
			loManager.value = "";
			loManager.focus();
		}
	}
	
	resetRedirect();
}

function gotoQuantity(pbGetPrice) {
	var i, j, loDiv;
	
	gbGetPrice = pbGetPrice;
	
	loDiv = ie4? eval("document.all.uniteditor") : document.getElementById("uniteditor");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.orderlinenotes") : document.getElementById("orderlinenotes");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.deleteconfirm") : document.getElementById("deleteconfirm");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.pricediv") : document.getElementById("pricediv");
	loDiv.style.visibility = "hidden";
	
	for (i = 0; i < ganUnitGroupSizeIDs.length; i++) {
		loDiv = ie4? eval("document.all.inclsides" + i.toString()) : document.getElementById("inclsides" + i.toString());
		loDiv.style.visibility = "hidden";
	}
	for (i = 0; i < ganSideGroupSpecialtyIDs.length; i++) {
		for (j = 0; j < ganSideGroupSizeIDs[i].length; j++) {
			loDiv = ie4? eval("document.all.specialtysides" + i.toString() + "-" + j.toString()) : document.getElementById("specialtysides" + i.toString() + "-" + j.toString());
			loDiv.style.visibility = "hidden";
		}
	}
	
	loDiv = ie4? eval("document.all.messagediv") : document.getElementById("messagediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.specialtystylediv") : document.getElementById("specialtystylediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.needstylediv") : document.getElementById("needstylediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.needtoppersdiv") : document.getElementById("needtoppersdiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById("managerdiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.mporeasondiv") : document.getElementById("mporeasondiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.quantitydiv") : document.getElementById("quantitydiv");
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function addToQuantity(psDigit) {
	var loText, lsText;
	
	loText = ie4? eval("document.all.quantity") : document.getElementById('quantity');
	lsText = loText.value;
	lsText += psDigit;
	loText.value = lsText;
	
	resetRedirect();
}

function backspaceQuantity() {
	var loText, lsText;
	
	loText = ie4? eval("document.all.quantity") : document.getElementById('quantity');
	lsText = loText.value;
	if (lsText.length > 0) {
		lsText = lsText.substr(0, (lsText.length - 1));
		loText.value = lsText;
	}
	
	resetRedirect();
}

function cancelQuantity() {
	var loText, loDiv;
	
	loText = ie4? eval("document.all.quantity") : document.getElementById('quantity');
	if (gdQuantity == 1) {
		loText.value = "";
	}
	else {
		loText.value = gdQuantity.toString();
	}
	
	gotoUnitEditor();
}

function setQuantity() {
	var loText, ldQuantity;
	
	loText = ie4? eval("document.all.quantity") : document.getElementById('quantity');
	if (loText.value == "") {
		return false;
	}
	else {
		ldQuantity = new Number(loText.value);
		if (ldQuantity == 0) {
			loText.value == "";
			return false;
		}
	}
	
	if (gbGetPrice) {
		gotoPrice();
	}
	else {
		gdQuantity = ldQuantity;
		
		ldOrderPrice = gdOrderTotal + (gdOrderLineCost * gdQuantity);
		ldOrderPrice = Math.round(ldOrderPrice * Math.pow(10,2))/Math.pow(10,2);
		
		s = "This Unit: " + gdQuantity.toString() + " @ " + FormatCurrency(gdOrderLineCost) + "&nbsp;&nbsp;&nbsp; Total: " + FormatCurrency(ldOrderPrice);
		loTotalDiv = ie4? eval("document.all.totaldiv") : document.getElementById("totaldiv");
		loTotalDiv.innerHTML = s;
		
		gotoUnitEditor();
	}
}

function gotoPrice() {
	var loText, loDiv;

	loText = ie4? eval("document.all.price") : document.getElementById('price');
	loText.value = FormatCurrency(gdOrderLineCost).substr(1);
	
	loDiv = ie4? eval("document.all.uniteditor") : document.getElementById("uniteditor");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.orderlinenotes") : document.getElementById("orderlinenotes");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.deleteconfirm") : document.getElementById("deleteconfirm");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.quantitydiv") : document.getElementById("quantitydiv");
	loDiv.style.visibility = "hidden";
	
	for (i = 0; i < ganUnitGroupSizeIDs.length; i++) {
		loDiv = ie4? eval("document.all.inclsides" + i.toString()) : document.getElementById("inclsides" + i.toString());
		loDiv.style.visibility = "hidden";
	}
	for (i = 0; i < ganSideGroupSpecialtyIDs.length; i++) {
		for (j = 0; j < ganSideGroupSizeIDs[i].length; j++) {
			loDiv = ie4? eval("document.all.specialtysides" + i.toString() + "-" + j.toString()) : document.getElementById("specialtysides" + i.toString() + "-" + j.toString());
			loDiv.style.visibility = "hidden";
		}
	}
	
	loDiv = ie4? eval("document.all.messagediv") : document.getElementById("messagediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.specialtystylediv") : document.getElementById("specialtystylediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.needstylediv") : document.getElementById("needstylediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.needtoppersdiv") : document.getElementById("needtoppersdiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById("managerdiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.mporeasondiv") : document.getElementById("mporeasondiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.pricediv") : document.getElementById("pricediv");
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function addToPrice(psDigit) {
	var loText, lsTex, loREt;
	
	loText = ie4? eval("document.all.price") : document.getElementById('price');
	loRE = /\./i;
	lsText = loText.value.replace(loRE, "");
	while (lsText.substr(0, 1) == "0") {
		lsText = lsText.substr(1);
	}
	lsText += psDigit;
	switch (lsText.length) {
		case 0:
			loText.value = lsText;
			break;
		case 1:
			loText.value = "0.0" + lsText;
			break;
		case 2:
			loText.value = "0." + lsText;
			break;
		default:
			loText.value = lsText.substr(0, (lsText.length - 2)) + "." + lsText.substr((lsText.length - 2));
			break;
	}
	
	resetRedirect();
}

function backspacePrice() {
	var loText, lsText, loRE;
	
	loText = ie4? eval("document.all.price") : document.getElementById('price');
	loRE = /\./i;
	lsText = loText.value.replace(loRE, "");
	while (lsText.substr(0, 1) == "0") {
		lsText = lsText.substr(1);
	}
	if (lsText.length > 0) {
		lsText = lsText.substr(0, (lsText.length - 1));
		switch (lsText.length) {
			case 0:
				loText.value = lsText;
				break;
			case 1:
				loText.value = "0.0" + lsText;
				break;
			case 2:
				loText.value = "0." + lsText;
				break;
			default:
				loText.value = lsText.substr(0, (lsText.length - 2)) + "." + lsText.substr((lsText.length - 2));
				break;
		}
	}
	
	resetRedirect();
}

function cancelPrice() {
	var loText, loDiv;
	
	loText = ie4? eval("document.all.quantity") : document.getElementById('quantity');
	if (gdQuantity == 1) {
		loText.value = "";
	}
	else {
		loText.value = gdQuantity.toString();
	}
	loText = ie4? eval("document.all.price") : document.getElementById('price');
	loText.value = gdOrderLineCost.toString();
	
	gotoUnitEditor();
}

function setPrice() {
	var loText, ldPrice;
	
	loText = ie4? eval("document.all.price") : document.getElementById('price');
	if (loText.value == "") {
		return false;
	}
	else {
		ldPrice = new Number(loText.value);
		if (ldPrice == 0) {
			loText.value == "";
			return false;
		}
		ldPrice = Math.round(ldPrice * Math.pow(10,2))/Math.pow(10,2);
	}
	
	gotoMPOReason();
}

function gotoMPOReason() {
	var loNotes, loDiv;
	
	loNotes = ie4? eval("document.all.mporeason") : document.getElementById('mporeason');
	loNotes.value = gsMPOReason;
	
	loDiv = ie4? eval("document.all.uniteditor") : document.getElementById("uniteditor");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.orderlinenotes") : document.getElementById("orderlinenotes");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.deleteconfirm") : document.getElementById("deleteconfirm");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.quantitydiv") : document.getElementById("quantitydiv");
	loDiv.style.visibility = "hidden";
	
	for (i = 0; i < ganUnitGroupSizeIDs.length; i++) {
		loDiv = ie4? eval("document.all.inclsides" + i.toString()) : document.getElementById("inclsides" + i.toString());
		loDiv.style.visibility = "hidden";
	}
	for (i = 0; i < ganSideGroupSpecialtyIDs.length; i++) {
		for (j = 0; j < ganSideGroupSizeIDs[i].length; j++) {
			loDiv = ie4? eval("document.all.specialtysides" + i.toString() + "-" + j.toString()) : document.getElementById("specialtysides" + i.toString() + "-" + j.toString());
			loDiv.style.visibility = "hidden";
		}
	}
	
	loDiv = ie4? eval("document.all.pricediv") : document.getElementById("pricediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.specialtystylediv") : document.getElementById("specialtystylediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.needstylediv") : document.getElementById("needstylediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.needtoppersdiv") : document.getElementById("needtoppersdiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById("managerdiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.messagediv") : document.getElementById("messagediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.mporeasondiv") : document.getElementById("mporeasondiv");
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function cancelMPOReason() {
	var loText, loDiv;
	
	loText = ie4? eval("document.all.quantity") : document.getElementById('quantity');
	if (gdQuantity == 1) {
		loText.value = "";
	}
	else {
		loText.value = gdQuantity.toString();
	}
	loText = ie4? eval("document.all.price") : document.getElementById('price');
	loText.value = gdOrderLineCost.toString();
	loText = ie4? eval("document.all.mporeason") : document.getElementById('mporeason');
	loText.value = gsMPOReason;
	
	gotoUnitEditor();
}

function addToMPOReason(psDigit) {
	var loNotes, lsNotes;
	
	loNotes = ie4? eval("document.all.mporeason") : document.getElementById('mporeason');
	lsNotes = loNotes.value;
	
	if (psDigit.length > 1) {
		if (lsNotes.length > 0) {
			lsNotes = lsNotes + " ";
		}
	}
	
	if (gbMPOReasonInLowerCase) {
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

function backspaceMPOReason() {
	var loNotes, lsNotes;
	
	loNotes = ie4? eval("document.all.mporeason") : document.getElementById('mporeason');
	lsNotes = loNotes.value;
	if (lsNotes.length > 0) {
		lsNotes = lsNotes.substr(0, (lsNotes.length - 1));
		loNotes.value = lsNotes;
	}
	
	resetRedirect();
}

function clearMPOReason() {
	var loNotes;
	
	loNotes = ie4? eval("document.all.mporeason") : document.getElementById('mporeason');
	loNotes.value = "";
	
	resetRedirect();
}

function properMPOReason() {
	var loNotes, i, lsText, lbDoUpper;
	
	loNotes = ie4? eval("document.all.mporeason") : document.getElementById('mporeason');
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

function shiftMPOReason() {
	var loObj;
	
	if (gbMPOReasonInLowerCase) {
		loObj = ie4? eval("document.all.mkey-q") : document.getElementById('mkey-q');
		loObj.innerHTML = "Q";
		loObj = ie4? eval("document.all.mkey-w") : document.getElementById('mkey-w');
		loObj.innerHTML = "W";
		loObj = ie4? eval("document.all.mkey-e") : document.getElementById('mkey-e');
		loObj.innerHTML = "E";
		loObj = ie4? eval("document.all.mkey-r") : document.getElementById('mkey-r');
		loObj.innerHTML = "R";
		loObj = ie4? eval("document.all.mkey-t") : document.getElementById('mkey-t');
		loObj.innerHTML = "T";
		loObj = ie4? eval("document.all.mkey-y") : document.getElementById('mkey-y');
		loObj.innerHTML = "Y";
		loObj = ie4? eval("document.all.mkey-u") : document.getElementById('mkey-u');
		loObj.innerHTML = "U";
		loObj = ie4? eval("document.all.mkey-i") : document.getElementById('mkey-i');
		loObj.innerHTML = "I";
		loObj = ie4? eval("document.all.mkey-o") : document.getElementById('mkey-o');
		loObj.innerHTML = "O";
		loObj = ie4? eval("document.all.mkey-p") : document.getElementById('mkey-p');
		loObj.innerHTML = "P";
		loObj = ie4? eval("document.all.mkey-a") : document.getElementById('mkey-a');
		loObj.innerHTML = "A";
		loObj = ie4? eval("document.all.mkey-s") : document.getElementById('mkey-s');
		loObj.innerHTML = "S";
		loObj = ie4? eval("document.all.mkey-d") : document.getElementById('mkey-d');
		loObj.innerHTML = "D";
		loObj = ie4? eval("document.all.mkey-f") : document.getElementById('mkey-f');
		loObj.innerHTML = "F";
		loObj = ie4? eval("document.all.mkey-g") : document.getElementById('mkey-g');
		loObj.innerHTML = "G";
		loObj = ie4? eval("document.all.mkey-h") : document.getElementById('mkey-h');
		loObj.innerHTML = "H";
		loObj = ie4? eval("document.all.mkey-j") : document.getElementById('mkey-j');
		loObj.innerHTML = "J";
		loObj = ie4? eval("document.all.mkey-k") : document.getElementById('mkey-k');
		loObj.innerHTML = "K";
		loObj = ie4? eval("document.all.mkey-l") : document.getElementById('mkey-l');
		loObj.innerHTML = "L";
		loObj = ie4? eval("document.all.mkey-z") : document.getElementById('mkey-z');
		loObj.innerHTML = "Z";
		loObj = ie4? eval("document.all.mkey-x") : document.getElementById('mkey-x');
		loObj.innerHTML = "X";
		loObj = ie4? eval("document.all.mkey-c") : document.getElementById('mkey-c');
		loObj.innerHTML = "C";
		loObj = ie4? eval("document.all.mkey-v") : document.getElementById('mkey-v');
		loObj.innerHTML = "V";
		loObj = ie4? eval("document.all.mkey-b") : document.getElementById('mkey-b');
		loObj.innerHTML = "B";
		loObj = ie4? eval("document.all.mkey-n") : document.getElementById('mkey-n');
		loObj.innerHTML = "N";
		loObj = ie4? eval("document.all.mkey-m") : document.getElementById('mkey-m');
		loObj.innerHTML = "M";
	}
	else {
		loObj = ie4? eval("document.all.mkey-q") : document.getElementById('mkey-q');
		loObj.innerHTML = "q";
		loObj = ie4? eval("document.all.mkey-w") : document.getElementById('mkey-w');
		loObj.innerHTML = "w";
		loObj = ie4? eval("document.all.mkey-e") : document.getElementById('mkey-e');
		loObj.innerHTML = "e";
		loObj = ie4? eval("document.all.mkey-r") : document.getElementById('mkey-r');
		loObj.innerHTML = "r";
		loObj = ie4? eval("document.all.mkey-t") : document.getElementById('mkey-t');
		loObj.innerHTML = "t";
		loObj = ie4? eval("document.all.mkey-y") : document.getElementById('mkey-y');
		loObj.innerHTML = "y";
		loObj = ie4? eval("document.all.mkey-u") : document.getElementById('mkey-u');
		loObj.innerHTML = "u";
		loObj = ie4? eval("document.all.mkey-i") : document.getElementById('mkey-i');
		loObj.innerHTML = "i";
		loObj = ie4? eval("document.all.mkey-o") : document.getElementById('mkey-o');
		loObj.innerHTML = "o";
		loObj = ie4? eval("document.all.mkey-p") : document.getElementById('mkey-p');
		loObj.innerHTML = "p";
		loObj = ie4? eval("document.all.mkey-a") : document.getElementById('mkey-a');
		loObj.innerHTML = "a";
		loObj = ie4? eval("document.all.mkey-s") : document.getElementById('mkey-s');
		loObj.innerHTML = "s";
		loObj = ie4? eval("document.all.mkey-d") : document.getElementById('mkey-d');
		loObj.innerHTML = "d";
		loObj = ie4? eval("document.all.mkey-f") : document.getElementById('mkey-f');
		loObj.innerHTML = "f";
		loObj = ie4? eval("document.all.mkey-g") : document.getElementById('mkey-g');
		loObj.innerHTML = "g";
		loObj = ie4? eval("document.all.mkey-h") : document.getElementById('mkey-h');
		loObj.innerHTML = "h";
		loObj = ie4? eval("document.all.mkey-j") : document.getElementById('mkey-j');
		loObj.innerHTML = "j";
		loObj = ie4? eval("document.all.mkey-k") : document.getElementById('mkey-k');
		loObj.innerHTML = "k";
		loObj = ie4? eval("document.all.mkey-l") : document.getElementById('mkey-l');
		loObj.innerHTML = "l";
		loObj = ie4? eval("document.all.mkey-z") : document.getElementById('mkey-z');
		loObj.innerHTML = "z";
		loObj = ie4? eval("document.all.mkey-x") : document.getElementById('mkey-x');
		loObj.innerHTML = "x";
		loObj = ie4? eval("document.all.mkey-c") : document.getElementById('mkey-c');
		loObj.innerHTML = "c";
		loObj = ie4? eval("document.all.mkey-v") : document.getElementById('mkey-v');
		loObj.innerHTML = "v";
		loObj = ie4? eval("document.all.mkey-b") : document.getElementById('mkey-b');
		loObj.innerHTML = "b";
		loObj = ie4? eval("document.all.mkey-n") : document.getElementById('mkey-n');
		loObj.innerHTML = "n";
		loObj = ie4? eval("document.all.mkey-m") : document.getElementById('mkey-m');
		loObj.innerHTML = "m";
	}
	
	gbMPOReasonInLowerCase = !gbMPOReasonInLowerCase;
	
	resetRedirect();
}

function saveMPOReason() {
	var loNotes, loText, ldPrice;
	
	loNotes = ie4? eval("document.all.mporeason") : document.getElementById('mporeason');
	gsMPOReason = loNotes.value;
	if (gsMPOReason.length > 0) {
		loText = ie4? eval("document.all.price") : document.getElementById('price');
		if (loText.value == "") {
			return false;
		}
		else {
			ldPrice = new Number(loText.value);
			if (ldPrice == 0) {
				loText.value == "";
				return false;
			}
			ldPrice = Math.round(ldPrice * Math.pow(10,2))/Math.pow(10,2);
		}
		
		loText = ie4? eval("document.all.quantity") : document.getElementById('quantity');
		gdQuantity = new Number(loText.value);
		
		if (gdOrderLineCost != ldPrice) {
			gdOrderLineCost = ldPrice;
			gbHasQuantityPrice = true;
		}
		ldOrderPrice = gdOrderTotal + (gdOrderLineCost * gdQuantity);
		ldOrderPrice = Math.round(ldOrderPrice * Math.pow(10,2))/Math.pow(10,2);
		
		s = "This Unit: " + gdQuantity.toString() + " @ " + FormatCurrency(gdOrderLineCost) + "&nbsp;&nbsp;&nbsp; Total: " + FormatCurrency(ldOrderPrice);
		loTotalDiv = ie4? eval("document.all.totaldiv") : document.getElementById("totaldiv");
		loTotalDiv.innerHTML = s;
		
		gotoUnitEditor();
	}
}

function gotoSpecialMessage(psSpecialMessage) {
	var loText, loDiv;

	loText = ie4? eval("document.all.specialmessage") : document.getElementById('specialmessage');
	loText.innerHTML = psSpecialMessage;
	
	loDiv = ie4? eval("document.all.uniteditor") : document.getElementById("uniteditor");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.orderlinenotes") : document.getElementById("orderlinenotes");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.deleteconfirm") : document.getElementById("deleteconfirm");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.quantitydiv") : document.getElementById("quantitydiv");
	loDiv.style.visibility = "hidden";
	
	for (i = 0; i < ganUnitGroupSizeIDs.length; i++) {
		loDiv = ie4? eval("document.all.inclsides" + i.toString()) : document.getElementById("inclsides" + i.toString());
		loDiv.style.visibility = "hidden";
	}
	for (i = 0; i < ganSideGroupSpecialtyIDs.length; i++) {
		for (j = 0; j < ganSideGroupSizeIDs[i].length; j++) {
			loDiv = ie4? eval("document.all.specialtysides" + i.toString() + "-" + j.toString()) : document.getElementById("specialtysides" + i.toString() + "-" + j.toString());
			loDiv.style.visibility = "hidden";
		}
	}
	
	loDiv = ie4? eval("document.all.pricediv") : document.getElementById("pricediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.specialtystylediv") : document.getElementById("specialtystylediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.needstylediv") : document.getElementById("needstylediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.needtoppersdiv") : document.getElementById("needtoppersdiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById("managerdiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.mporeasondiv") : document.getElementById("mporeasondiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.messagediv") : document.getElementById("messagediv");
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function gotoSpecialtyStyle(psSpecialty, pnSpecialStyleID, psSpecialStyle) {
	var loText, lsButtons, i, j, lbFound, lnCount, loSpan, loDiv;

	loText = ie4? eval("document.all.specialtystylemessage") : document.getElementById('specialtystylemessage');
	loText.innerHTML = "Vito suggests the " + psSpecialty + " pizza with a " + psSpecialStyle + " crust. Is that OK?";
	
	lsButtons = "";
	
	// Make OK button
	for (i = 0; i < ganSizeStyleIDs.length; i++) {
		if (ganSizeStyleIDs[i] == pnSpecialStyleID) {
			lsButtons = "<button onclick=\"gbNeedStyle = false; setStyle(" + ganSizeStyleIDs[i].toString() + ");\">OK</button><br/>&nbsp;<br/>&nbsp;<br/>";
		}
	}
	
	// Make other buttons
	lnCount = 0;
	for (i = 0; i < ganSizeStyleIDs.length; i++) {
		if (ganSizeStyleIDs[i] != pnSpecialStyleID) {
			lbFound = false;
			
			for (j = 0; j < ganSizeStyleSizeIDs[i].length; j++) {
				if (ganSizeStyleSizeIDs[i][j] == gnSizeID) {
					lbFound = true;
					j = ganSizeStyleSizeIDs[i].length;
				}
			}
			
			if (lbFound) {
				if (lnCount) {
					lsButtons = lsButtons + "&nbsp;&nbsp;&nbsp;";
				}
				lsButtons = lsButtons + "<button onclick=\"gbNeedStyle = false; setStyle(" + ganSizeStyleIDs[i].toString() + ");\">" + gasSizeStyleShortDescriptions[i] + "</button>";
				lnCount = lnCount + 1;
			}
		}
	}
	
	loSpan = ie4? eval("document.all.specialtystylebuttons") : document.getElementById('specialtystylebuttons');
	loSpan.innerHTML = lsButtons;
	
	loDiv = ie4? eval("document.all.uniteditor") : document.getElementById("uniteditor");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.orderlinenotes") : document.getElementById("orderlinenotes");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.deleteconfirm") : document.getElementById("deleteconfirm");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.quantitydiv") : document.getElementById("quantitydiv");
	loDiv.style.visibility = "hidden";
	
	for (i = 0; i < ganUnitGroupSizeIDs.length; i++) {
		loDiv = ie4? eval("document.all.inclsides" + i.toString()) : document.getElementById("inclsides" + i.toString());
		loDiv.style.visibility = "hidden";
	}
	for (i = 0; i < ganSideGroupSpecialtyIDs.length; i++) {
		for (j = 0; j < ganSideGroupSizeIDs[i].length; j++) {
			loDiv = ie4? eval("document.all.specialtysides" + i.toString() + "-" + j.toString()) : document.getElementById("specialtysides" + i.toString() + "-" + j.toString());
			loDiv.style.visibility = "hidden";
		}
	}
	
	loDiv = ie4? eval("document.all.pricediv") : document.getElementById("pricediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.messagediv") : document.getElementById("messagediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.needstylediv") : document.getElementById("needstylediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.needtoppersdiv") : document.getElementById("needtoppersdiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById("managerdiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.mporeasondiv") : document.getElementById("mporeasondiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.specialtystylediv") : document.getElementById("specialtystylediv");
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function gotoNeedStyle() {
	var lsButtons, i, j, lbFound, lnCount, loSpan, loDiv;

	lsButtons = "";
	lnCount = 0;
	for (i = 0; i < ganSizeStyleIDs.length; i++) {
		lbFound = false;
		
		for (j = 0; j < ganSizeStyleSizeIDs[i].length; j++) {
			if (ganSizeStyleSizeIDs[i][j] == gnSizeID) {
				lbFound = true;
				j = ganSizeStyleSizeIDs[i].length;
			}
		}
		
		if (lbFound) {
			if (lnCount) {
				lsButtons = lsButtons + "&nbsp;&nbsp;&nbsp;";
			}
			lsButtons = lsButtons + "<button onclick=\"gbNeedStyle = false; setStyle(" + ganSizeStyleIDs[i].toString() + ");\">" + gasSizeStyleShortDescriptions[i] + "</button>";
			lnCount = lnCount + 1;
		}
	}
	
	loSpan = ie4? eval("document.all.needstylebuttons") : document.getElementById('needstylebuttons');
	loSpan.innerHTML = lsButtons;
	
	loDiv = ie4? eval("document.all.uniteditor") : document.getElementById("uniteditor");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.orderlinenotes") : document.getElementById("orderlinenotes");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.deleteconfirm") : document.getElementById("deleteconfirm");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.quantitydiv") : document.getElementById("quantitydiv");
	loDiv.style.visibility = "hidden";
	
	for (i = 0; i < ganUnitGroupSizeIDs.length; i++) {
		loDiv = ie4? eval("document.all.inclsides" + i.toString()) : document.getElementById("inclsides" + i.toString());
		loDiv.style.visibility = "hidden";
	}
	for (i = 0; i < ganSideGroupSpecialtyIDs.length; i++) {
		for (j = 0; j < ganSideGroupSizeIDs[i].length; j++) {
			loDiv = ie4? eval("document.all.specialtysides" + i.toString() + "-" + j.toString()) : document.getElementById("specialtysides" + i.toString() + "-" + j.toString());
			loDiv.style.visibility = "hidden";
		}
	}
	
	loDiv = ie4? eval("document.all.pricediv") : document.getElementById("pricediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.messagediv") : document.getElementById("messagediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.needtoppersdiv") : document.getElementById("needtoppersdiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.specialtystylediv") : document.getElementById("specialtystylediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById("managerdiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.mporeasondiv") : document.getElementById("mporeasondiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.needstylediv") : document.getElementById("needstylediv");
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function gotoNeedToppers() {
	var lsMessage, i, loText, loDiv;
	
	gbNeedToppers = false;
	
	lsMessage = "";
	for (i = 0; i < ganUnitTopperIDs.length; i++) {
		if (i < 2) {
			if (i == 0) {
				lsMessage = lsMessage + gasUnitTopperDescriptions[i];
			}
			else {
				lsMessage = lsMessage + " or " + gasUnitTopperDescriptions[i];
			}
		}
		else {
			i = ganUnitTopperIDs.length;
		}
	}
	lsMessage = lsMessage + " on the edge for free?";
	
	loText = ie4? eval("document.all.needtoppersmessage") : document.getElementById('needtoppersmessage');
	loText.innerHTML = lsMessage;
	
	loDiv = ie4? eval("document.all.uniteditor") : document.getElementById("uniteditor");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.orderlinenotes") : document.getElementById("orderlinenotes");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.deleteconfirm") : document.getElementById("deleteconfirm");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.quantitydiv") : document.getElementById("quantitydiv");
	loDiv.style.visibility = "hidden";
	
	for (i = 0; i < ganUnitGroupSizeIDs.length; i++) {
		loDiv = ie4? eval("document.all.inclsides" + i.toString()) : document.getElementById("inclsides" + i.toString());
		loDiv.style.visibility = "hidden";
	}
	for (i = 0; i < ganSideGroupSpecialtyIDs.length; i++) {
		for (j = 0; j < ganSideGroupSizeIDs[i].length; j++) {
			loDiv = ie4? eval("document.all.specialtysides" + i.toString() + "-" + j.toString()) : document.getElementById("specialtysides" + i.toString() + "-" + j.toString());
			loDiv.style.visibility = "hidden";
		}
	}
	
	loDiv = ie4? eval("document.all.pricediv") : document.getElementById("pricediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.specialtystylediv") : document.getElementById("specialtystylediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.messagediv") : document.getElementById("messagediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.needstylediv") : document.getElementById("needstylediv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.managerdiv") : document.getElementById("managerdiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.mporeasondiv") : document.getElementById("mporeasondiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.needtoppersdiv") : document.getElementById("needtoppersdiv");
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function saveOrderLine(pbAddAnother) {
	var lsLocation, lbFound;
	
	if (!gbDonePressed) {
		gbDonePressed = true;
		
		if (gbNeedStyle && !gbHasQuantityPrice) {
			gbDonePressed = false;
			gbNeedStyleDoneIsDone = true;
			gbAddAnother = pbAddAnother;
			
			resetRedirect();
			gotoNeedStyle();
		}
		else {
			if (gbNeedToppers) {
				gbDonePressed = false;
				gbAddAnother = pbAddAnother;
				
				toggleToppers();
				
				resetRedirect();
				gotoNeedToppers();
			}
			else {
				lbFound = false;
				
				if (!gbHasSpecSides) {
					for (i = 0; i < ganUnitGroupSizeIDs.length; i++) {
						if (ganUnitGroupSizeIDs[i] == gnSizeID) {
							lbFound = true;
						}
					}
				}
				else {
					for (i = 0; i < ganSideGroupSpecialtyIDs.length; i++) {
						if (ganSideGroupSpecialtyIDs[i] == gnSpecialtyID) {
							for (j = 0; j < ganSideGroupSizeIDs[i].length; j++) {
								if (ganSideGroupSizeIDs[i][j] == gnSizeID) {
									lbFound = true;
								}
							}
						}
					}
				}
				
				if (lbFound && gbNeedIncludedSides) {
					gbDonePressed = false;
					gbIncludedSidesDoneIsDone = true;
					gbAddAnother = pbAddAnother;
					
					resetRedirect();
					gotoInclSides();
				}
				else {
					lsLocation = "unitselect.asp?save=yes&l=" + gnOrderLineID.toString();
					lsLocation = lsLocation + "&u=" + gnUnitID.toString();
					lsLocation = lsLocation + "&SpecialtyID=" + gnSpecialtyID;
					lsLocation = lsLocation + "&SizeID=" + gnSizeID;
					lsLocation = lsLocation + "&StyleID=" + gnStyleID;
					lsLocation = lsLocation + "&Half1SauceID=" + gnHalf1SauceID;
					lsLocation = lsLocation + "&Half2SauceID=" + gnHalf2SauceID;
					lsLocation = lsLocation + "&Half1SauceModifierID=" + gnHalf1SauceModifierID;
					lsLocation = lsLocation + "&Half2SauceModifierID=" + gnHalf2SauceModifierID;
					lsLocation = lsLocation + "&OrderLineNotes=" + encodeURIComponent(gsOrderLineNotes);
					
					for (i = 0; i < ganItemIDs.length; i++) {
						if (ganItemIDs[i] > 0) {
							lsLocation = lsLocation + "&ItemID=" + ganItemIDs[i] + "," + ganHalfIDs[i];
						}
					}
					for (i = 0; i < ganTopperIDs.length; i++) {
						if (ganTopperIDs[i] > 0) {
							lsLocation = lsLocation + "&TopperID=" + ganTopperIDs[i] + "," + ganTopperHalfIDs[i];
						}
					}
					for (i = 0; i < ganFreeSideIDs.length; i++) {
						if (ganFreeSideIDs[i] > 0) {
							lsLocation = lsLocation + "&FreeSideID=" + ganFreeSideIDs[i];
						}
					}
					for (i = 0; i < ganAddSideIDs.length; i++) {
						if (ganAddSideIDs[i] > 0) {
							lsLocation = lsLocation + "&AddSideID=" + ganAddSideIDs[i];
						}
					}
					lsLocation = lsLocation + "&i=" + gdQuantity.toString()
					if (gbHasQuantityPrice) {
						lsLocation = lsLocation + "&s=" + gdOrderLineCost.toString() + "&MPOReason=" + encodeURIComponent(gsMPOReason);
					}
					if (pbAddAnother) {
						lsLocation = lsLocation + "&Another=yes"
					}
					
					window.location = lsLocation;
				}
			}
		}
	}
}
//-->
</script>
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
						<div id="uniteditor" style="position: absolute; left: 0px; top: 0px; width: 1010px; height: 723px;">
							<table cellpadding="0" cellspacing="0" width="1010" height="723">
								<tr>
									<td valign="top" width="680">
										<table cellpadding="0" cellspacing="0" width="680" height="335">
											<tr>
												<td align="center" valign="top" width="170" style="border: 1px #000000 solid">
<%
If ganUnitSizeIDs(0) <> 0 Then
%>
													<strong>Size</strong><br/>
													<div style="position: relative; width: 160px;">
<%
	For i = 0 To UBound(ganUnitSizeIDs)
		If i Mod 3 = 0 Then
			If i > 0 And UBound(ganUnitSizeIDs) > 3 Then
%>
															<button style="width: 150px" onclick="toggleDivs('sizediv<%=Int(i/3)-1%>', 'sizediv<%=Int(i/3)%>')">
															(Next)</button>
														</div>
														<div id="sizediv<%=Int(i/3)%>" style="position: absolute; top: 0px; left: 0px; width: 160px; visibility: hidden;">
<%
			Else
				If i = 0 Then
%>
														<div id="sizediv<%=Int(i/3)%>" style="position: absolute; top: 0px; left: 0px; width: 160px;">
<%
				End If
			End If
		End If
%>
															<button style="width: 150px" onclick="setSize(<%=ganUnitSizeIDs(i)%>)"><%=gasUnitSizeShortDescriptions(i)%></button>
<%
	Next
	
	' Add hidden buttons here
	If ((UBound(ganUnitSizeIDs) + 1) Mod 3) > 0 And UBound(ganUnitSizeIDs) <> 3 Then
		For i = ((UBound(ganUnitSizeIDs) + 1) Mod 3) To 2
%>
															<button style="width: 150px; background-color: #C0C0C0;">
															&nbsp;</button>
<%
		Next
	End If

	If UBound(ganUnitSizeIDs) > 3 Then
%>
															<button style="width: 150px" onclick="toggleDivs('sizediv<%=Int(UBound(ganUnitSizeIDs)/3)%>', 'sizediv0')">
															(Next)</button>
<%
	Else
		If UBound(ganUnitSizeIDs) <> 3 Then
%>
															<button style="width: 150px; background-color: #C0C0C0;">
															&nbsp;</button>
<%
		End If
	End If
%>
														</div>
													</div>
<%
End If
%>
												</td>
												<td align="center" valign="top" width="170" style="border: 1px #000000 solid">
<%
If ganSizeStyleIDs(0) <> 0 Then
%>
													<strong>Style</strong><br/>
													<div style="position: relative; width: 160px;">
<%
	For i = 0 To UBound(ganUnitSizeIDs)
%>
														<div id="sizestylediv<%=ganUnitSizeIDs(i)%>" style="position: absolute; top: 0px; left: 0px; width: 160px; visibility: <%If gnSizeID = ganUnitSizeIDs(i) Then Response.Write("visible") Else Response.Write("hidden") End If%>;">
															<div style="position: relative; width: 160px;">
<%
		l = 0
		For j = 0 To UBound(ganSizeStyleIDs)
			For k = 0 To UBound(ganSizeStyleSizeIDs, 2)
				If ganSizeStyleSizeIDs(j, k) = ganUnitSizeIDs(i) Then
					l = l + 1
				End If
			Next
		Next
		l = l - 1
		
		m = 0
		For j = 0 To UBound(ganSizeStyleIDs)
			For k = 0 To UBound(ganSizeStyleSizeIDs, 2)
				If ganSizeStyleSizeIDs(j, k) = ganUnitSizeIDs(i) Then
					If m Mod 3 = 0 Then
						If m > 0 And l > 3 Then
%>
																	<button style="width: 150px" onclick="toggleDivs('stylediv<%=ganUnitSizeIDs(i)%>-<%=Int(m/3)-1%>', 'stylediv<%=ganUnitSizeIDs(i)%>-<%=Int(m/3)%>')">
																	(Next)</button>
																</div>
																<div id="stylediv<%=ganUnitSizeIDs(i)%>-<%=Int(m/3)%>" style="position: absolute; top: 0px; left: 0px; width: 160px; visibility: hidden;">
<%
						Else
							If m = 0 Then
%>
																<div id="stylediv<%=ganUnitSizeIDs(i)%>-<%=Int(m/3)%>" style="position: absolute; top: 0px; left: 0px; width: 160px;">
<%
							End If
						End If
					End If
%>
																	<button style="width: 150px" onclick="gbNeedStyle = false; gbManualStyle = true; setStyle(<%=ganSizeStyleIDs(j)%>)"><%=gasSizeStyleShortDescriptions(j)%></button>
<%
					m = m + 1
				End If
			Next
		Next
		m = m - 1
		
		' Add hidden buttons here
		If m = -1 Then
%>
																<div id="stylediv<%=ganUnitSizeIDs(i)%>-0" style="position: absolute; top: 0px; left: 0px; width: 160px;">
																	<button style="width: 150px; background-color: #C0C0C0;">&nbsp;</button>
																	<button style="width: 150px; background-color: #C0C0C0;">&nbsp;</button>
																	<button style="width: 150px; background-color: #C0C0C0;">&nbsp;</button>
																	<button style="width: 150px; background-color: #C0C0C0;">&nbsp;</button>
<%
		Else
			If ((m + 1) Mod 3) > 0 And l <> 3 Then
				For j = ((m + 1) Mod 3) To 2
%>
																	<button style="width: 150px; background-color: #C0C0C0;">&nbsp;</button>
<%
				Next
			End If
			
			If m > 3 Then
%>
																	<button style="width: 150px" onclick="toggleDivs('stylediv<%=ganUnitSizeIDs(i)%>-<%=Int(m/3)%>', 'stylediv<%=ganUnitSizeIDs(i)%>-0')">(Next)</button>
<%
			Else
				If l <> 3 Then
%>
																	<button style="width: 150px; background-color: #C0C0C0;">&nbsp;</button>
<%
				End If
			End If
		End If
%>
																</div>
															</div>
														</div>
<%
	Next
%>
													</div>
<%
End If
%>
												</td>
												<td align="center" valign="top" width="170" style="border: 1px #000000 solid">
<%
If ganSauceIDs(0) <> 0 Then
%>
													<strong>Sauce</strong><br/>
													<div style="position: relative; width: 160px;">
<%
	For i = 0 To UBound(ganSauceIDs)
		If i Mod 3 = 0 Then
			If i > 0 And UBound(ganSauceIDs) > 3 Then
%>
															<button style="width: 150px" onclick="toggleDivs('saucediv<%=Int(i/3)-1%>', 'saucediv<%=Int(i/3)%>')">
															(Next)</button>
														</div>
														<div id="saucediv<%=Int(i/3)%>" style="position: absolute; top: 0px; left: 0px; width: 160px; visibility: hidden;">
<%
			Else
				If i = 0 Then
%>
														<div id="saucediv<%=Int(i/3)%>" style="position: absolute; top: 0px; left: 0px; width: 160px;">
<%
				End If
			End If
		End If
%>
															<button style="width: 150px" onclick="setSauce(<%=ganSauceIDs(i)%>)"><%=gasSauceShortDescriptions(i)%></button>
<%
	Next
	
	' Add hidden buttons here
	If ((UBound(ganSauceIDs) + 1) Mod 3) > 0 And UBound(ganSauceIDs) <> 3 Then
		For i = ((UBound(ganSauceIDs) + 1) Mod 3) To 2
%>
															<button style="width: 150px; background-color: #C0C0C0;">
															&nbsp;</button>
<%
		Next
	End If

	If UBound(ganSauceIDs) > 3 Then
%>
															<button style="width: 150px" onclick="toggleDivs('saucediv<%=Int(UBound(ganSauceIDs)/3)%>', 'saucediv0')">
															(Next)</button>
<%
	Else
		If UBound(ganSauceIDs) <> 3 Then
%>
															<button style="width: 150px; background-color: #C0C0C0;">
															&nbsp;</button>
<%
		End If
	End If
%>
														</div>
													</div>
<%
End If
%>
												</td>
												<td align="center" valign="top" width="170" style="border: 1px #000000 solid">
<%
If ganSauceIDs(0) <> 0 And ganSauceModifierIDs(0) <> 0 Then
%>
													<strong>Sauce Modifier</strong><br/>
													<div style="position: relative; width: 160px;">
<%
	For i = 0 To UBound(ganSauceModifierIDs)
		If i Mod 3 = 0 Then
			If i > 0 And UBound(ganSauceModifierIDs) > 3 Then
%>
															<button style="width: 150px" onclick="toggleDivs('saucemoddiv<%=Int(i/3)-1%>', 'saucemoddiv<%=Int(i/3)%>')">
															(Next)</button>
														</div>
														<div id="saucemoddiv<%=Int(i/3)%>" style="position: absolute; top: 0px; left: 0px; width: 160px; visibility: hidden;">
<%
			Else
				If i = 0 Then
%>
														<div id="saucemoddiv<%=Int(i/3)%>" style="position: absolute; top: 0px; left: 0px; width: 160px;">
<%
				End If
			End If
		End If
%>
															<button style="width: 150px" onclick="setSauceModifier(<%=ganSauceModifierIDs(i)%>)"><%=gasSauceModifierShortDescriptions(i)%></button>
<%
	Next
	
	' Add hidden buttons here
	If ((UBound(ganSauceModifierIDs) + 1) Mod 3) > 0 And UBound(ganSauceModifierIDs) <> 3 Then
		For i = ((UBound(ganSauceModifierIDs) + 1) Mod 3) To 2
%>
															<button style="width: 150px; background-color: #C0C0C0;">
															&nbsp;</button>
<%
		Next
	End If

	If UBound(ganSauceModifierIDs) > 3 Then
%>
															<button style="width: 150px" onclick="toggleDivs('saucemoddiv<%=Int(UBound(ganSauceModifierIDs)/3)%>', 'saucemoddiv0')">
															(Next)</button>
<%
	Else
		If UBound(ganSauceModifierIDs) <> 3 Then
%>
															<button style="width: 150px; background-color: #C0C0C0;">
															&nbsp;</button>
<%
		End If
	End If
%>
														</div>
													</div>
<%
End If
%>
												</td>
											</tr>
										</table>
										<table cellpadding="0" cellspacing="0" width="680" height="388">
											<tr height="110">
												<td align="center" valign="top" width="680" style="border: 1px #000000 solid">
													<div style="position: relative; width: 670px;">
														<div style="position: absolute; top: 15px; left: 20px;"><button id="deletebutton" style="width: 75px; height: 75px;" onclick="toggleDelete();">
															Delete</button></div>
														<div style="position: absolute; top: 15px; left: 96px;"><button id="wholebutton" style="width: 75px; height: 75px;" onclick="toggleWhole();">
															Whole</button></div>
														<div style="position: absolute; top: 15px; left: 172px;"><button id="itemsbutton" style="width: 75px; height: 75px;" onclick="toggleItems();">
															Items</button></div>
														<div style="position: absolute; top: 15px; left: 248px;"><button id="specialtybutton" style="width: 85px; height: 75px; color: #FFFFFF;" onclick="toggleSpecialty();">
															Specialty</button></div>
														<div style="position: absolute; top: 15px; left: 334px;"><button id="toppersbutton" style="width: 75px; height: 75px;" onclick="toggleToppers();">
															Toppers</button></div>
														<div style="position: absolute; top: 15px; left: 410px;"><button id="sidesbutton" style="width: 75px; height: 75px;" onclick="toggleSides();">
															Addl. Sides</button></div>
														<div style="position: absolute; top: 15px; left: 486px;"><button id="inclsidesbutton" style="width: 75px; height: 75px;" onclick="gotoInclSides();">
															Incl. Sides</button></div>
														<div style="position: absolute; top: 15px; left: 562px;"><button id="notesbutton" style="width: 75px; height: 75px;" onclick="gotoOrderLineNotes();">
															Notes</button></div>
													</div>
												</td>
											</tr>
											<tr height="270">
												<td align="center" valign="top" width="680" style="border: 1px #000000 solid">
													<div style="position: relative; width: 670px;">
<%
If ganUnitItemIDs(0) <> 0 Then
	lnTop = 0
	lnLeft = 0
	For i = 0 To UBound(ganUnitItemIDs)
		If i Mod 20 = 0 Then
			If i > 0 And UBound(ganUnitItemIDs) > 20 Then
%>
															<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 95px; height: 75px;" onclick="toggleDivs('itemdiv<%=Int(i/20)-1%>', 'itemdiv<%=Int(i/20)%>')">
																(Next)</button></div>
														</div>
														<div id="itemdiv<%=Int(i/20)%>" style="position: absolute; top: 0px; left: 0px; width: 670px; visibility: hidden;">
<%
			Else
				If i = 0 Then
%>
														<div id="itemdiv<%=Int(i/20)%>" style="position: absolute; top: 0px; left: 0px; width: 670px; visibility: hidden;">
<%
				End If
			End If
			
			lnTop = 0
			lnLeft = 0
		End If
%>
															<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button id="item-<%=ganUnitItemIDs(i)%>" name="item-<%=ganUnitItemIDs(i)%>" style="width: 95px; height: 75px;" onclick="addItem(<%=ganUnitItemIDs(i)%>)"><%=gasUnitItemShortDescriptions(i)%></button></div>
<%
		If lnLeft = 576 Then
			lnTop = lnTop + 76
			lnLeft = 0
		Else
			lnLeft = lnLeft + 96
		End If
	Next
	
	' Add hidden buttons here
	If ((UBound(ganUnitItemIDs) + 1) Mod 20) > 0 And UBound(ganUnitItemIDs) <> 20 Then
		For i = ((UBound(ganUnitItemIDs) + 1) Mod 20) To 19
%>
															<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 95px; height: 75px; background-color: #C0C0C0;">
																&nbsp;</button></div>
<%
			If lnLeft = 576 Then
				lnTop = lnTop + 76
				lnLeft = 0
			Else
				lnLeft = lnLeft + 96
			End If
		Next
	End If

	If UBound(ganUnitItemIDs) > 20 Then
%>
															<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 95px; height: 75px;" onclick="toggleDivs('itemdiv<%=Int(UBound(ganUnitItemIDs)/20)%>', 'itemdiv0')">
																(Next)</button></div>
<%
	Else
		If UBound(ganUnitItemIDs) <> 20 Then
%>
															<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 95px; height: 75px; background-color: #C0C0C0;">
																&nbsp;</button></div>
<%
		End If
	End If
%>
														</div>
<%
End If

If ganUnitTopperIDs(0) <> 0 Then
	lnTop = 0
	lnLeft = 0
	For i = 0 To UBound(ganUnitTopperIDs)
		If i Mod 20 = 0 Then
			If i > 0 And UBound(ganUnitTopperIDs) > 20 Then
%>
															<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 95px; height: 75px;" onclick="toggleDivs('topperdiv<%=Int(i/20)-1%>', 'topperdiv<%=Int(i/20)%>')">
																(Next)</button></div>
														</div>
														<div id="topperdiv<%=Int(i/20)%>" style="position: absolute; top: 0px; left: 0px; width: 670px; visibility: hidden;">
<%
			Else
				If i = 0 Then
%>
														<div id="topperdiv<%=Int(i/20)%>" style="position: absolute; top: 0px; left: 0px; width: 670px; visibility: hidden;">
<%
				End If
			End If
			
			lnTop = 0
			lnLeft = 0
		End If
%>
															<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 95px; height: 75px;" onclick="addTopper(<%=ganUnitTopperIDs(i)%>)"><%=gasUnitTopperShortDescriptions(i)%></button></div>
<%
		If lnLeft = 576 Then
			lnTop = lnTop + 76
			lnLeft = 0
		Else
			lnLeft = lnLeft + 96
		End If
	Next
	
	' Add hidden buttons here
	If ((UBound(ganUnitTopperIDs) + 1) Mod 20) > 0 And UBound(ganUnitTopperIDs) <> 20 Then
		For i = ((UBound(ganUnitTopperIDs) + 1) Mod 20) To 19
%>
															<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 95px; height: 75px; background-color: #C0C0C0;">
																&nbsp;</button></div>
<%
			If lnLeft = 576 Then
				lnTop = lnTop + 76
				lnLeft = 0
			Else
				lnLeft = lnLeft + 96
			End If
		Next
	End If

	If UBound(ganUnitTopperIDs) > 20 Then
%>
															<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 95px; height: 75px;" onclick="toggleDivs('topperdiv<%=Int(UBound(ganUnitTopperIDs)/20)%>', 'topperdiv0')">
																(Next)</button></div>
<%
	Else
		If UBound(ganUnitTopperIDs) <> 20 Then
%>
															<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 95px; height: 75px; background-color: #C0C0C0;">
																&nbsp;</button></div>
<%
		End If
	End If
%>
														</div>
<%
End If

If ganUnitSideIDs(0) <> 0 Then
	lnTop = 0
	lnLeft = 0
	For i = 0 To UBound(ganUnitSideIDs)
		If i Mod 20 = 0 Then
			If i > 0 And UBound(ganUnitSideIDs) >= 20 Then
%>
															<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 95px; height: 75px;" onclick="toggleDivs('sidediv<%=Int(i/20)-1%>', 'sidediv<%=Int(i/20)%>')">
																(Next)</button></div>
														</div>
														<div id="sidediv<%=Int(i/20)%>" style="position: absolute; top: 0px; left: 0px; width: 670px; visibility: hidden;">
<%
			Else
%>
														<div id="sidediv<%=Int(i/20)%>" style="position: absolute; top: 0px; left: 0px; width: 670px; visibility: hidden;">
<%
			End If
			
			lnTop = 0
			lnLeft = 0
		End If
%>
															<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 95px; height: 75px;" onclick="addSide(<%=ganUnitSideIDs(i)%>)"><%=gasUnitSideShortDescriptions(i)%></button></div>
<%
		If lnLeft = 576 Then
			lnTop = lnTop + 76
			lnLeft = 0
		Else
			lnLeft = lnLeft + 96
		End If
	Next
	
	' Add hidden buttons here
	If ((UBound(ganUnitSideIDs) + 1) Mod 20) > 0 And UBound(ganUnitSideIDs) <> 20 Then
		For i = ((UBound(ganUnitSideIDs) + 1) Mod 20) To 19
%>
															<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 95px; height: 75px; background-color: #C0C0C0;">
																&nbsp;</button></div>
<%
			If lnLeft = 576 Then
				lnTop = lnTop + 76
				lnLeft = 0
			Else
				lnLeft = lnLeft + 96
			End If
		Next
	End If

	If UBound(ganUnitSideIDs) >= 20 Then
%>
															<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 95px; height: 75px;" onclick="toggleDivs('sidediv<%=Int(UBound(ganUnitSideIDs)/20)%>', 'sidediv0')">
																(Next)</button></div>
<%
	Else
		If UBound(ganUnitSideIDs) <> 20 Then
%>
															<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 95px; height: 75px; background-color: #C0C0C0;">
																&nbsp;</button></div>
<%
		End If
	End If
%>
														</div>
<%
End If

If ganUnitSpecialtyIDs(0) <> 0 Then
	lnTop = 0
	lnLeft = 0
	For i = 0 To UBound(ganUnitSpecialtyIDs)
		If i Mod 20 = 0 Then
			If i > 0 And UBound(ganUnitSpecialtyIDs) > 20 Then
%>
															<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 95px; height: 75px;" onclick="toggleDivs('specialtydiv<%=Int(i/20)-1%>', 'specialtydiv<%=Int(i/20)%>')">
																(Next)</button></div>
														</div>
														<div id="specialtydiv<%=Int(i/20)%>" style="position: absolute; top: 0px; left: 0px; width: 670px; visibility: hidden;">
<%
			Else
				If i = 0 Then
%>
														<div id="specialtydiv<%=Int(i/20)%>" style="position: absolute; top: 0px; left: 0px; width: 670px;">
<%
				End If
			End If
			
			lnTop = 0
			lnLeft = 0
		End If
%>
															<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 95px; height: 75px;" onclick="setSpecialty(<%=ganUnitSpecialtyIDs(i)%>)"><%=gasUnitSpecialtyShortDescriptions(i)%></button></div>
<%
		If lnLeft = 576 Then
			lnTop = lnTop + 76
			lnLeft = 0
		Else
			lnLeft = lnLeft + 96
		End If
	Next
	
	' Add hidden buttons here
	If ((UBound(ganUnitSpecialtyIDs) + 1) Mod 20) > 0 And UBound(ganUnitSpecialtyIDs) <> 20 Then
		For i = ((UBound(ganUnitSpecialtyIDs) + 1) Mod 20) To 19
%>
															<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 95px; height: 75px; background-color: #C0C0C0;">
																&nbsp;</button></div>
<%
			If lnLeft = 576 Then
				lnTop = lnTop + 76
				lnLeft = 0
			Else
				lnLeft = lnLeft + 96
			End If
		Next
	End If

	If UBound(ganUnitSpecialtyIDs) > 20 Then
%>
															<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 95px; height: 75px;" onclick="toggleDivs('specialtydiv<%=Int(UBound(ganUnitSpecialtyIDs)/20)%>', 'specialtydiv0')">
																(Next)</button></div>
<%
	Else
		If UBound(ganUnitSpecialtyIDs) <> 20 Then
%>
															<div style="position: absolute; top: <%=lnTop%>px; left: <%=lnLeft%>px;"><button style="width: 95px; height: 75px; background-color: #C0C0C0;">
																&nbsp;</button></div>
<%
		End If
	End If
%>
														</div>
<%
End If
%>
													</div>
												</td>
											</tr>
										</table>
									</td>
									<td align="right" valign="top" width="330">
										<div style="position: relative; width: 320px; height: 470px; text-align: left; background-color: #FFFFFF;">
											<div id="editdiv0" style="position: absolute; top: 0px; padding: 5px;">
<%
If gnSizeID > 0 Then
	lsTmp = gsSizeDescription & " "
Else
	lsTmp = ""
End If
lsTmp = lsTmp & gsUnitDescription
%>
												<div id="editunit" style="height: 15px;"><%=lsTmp%></div>
<%
If gnSpecialtyID > 0 Then
	lsTmp = gsSpecialtyDescription
%>
												<div id="editspecialty" style="height: 15px;"><%=lsTmp%></div>
<%
Else
%>
												<div id="editspecialty" style="height: 0px; visibility: hidden;"></div>
<%
End If

If gnStyleID > 0 Then
	lsTmp = gsStyleDescription
%>
												<div id="editstyle" style="height: 15px;"><%=lsTmp%></div>
<%
Else
%>
												<div id="editstyle" style="height: 0px; visibility: hidden;"></div>
<%
End If

If gnHalf1SauceID > 0 Then
	If (gnHalf1SauceID = gnHalf2SauceID) And (gnHalf1SauceModifierID = gnHalf2SauceModifierID) Then
		lsTmp = "Whole Sauce: " & gsHalf1SauceDescription
	Else
		lsTmp = "1st Half Sauce: " & gsHalf1SauceDescription
	End If
	If gnHalf1SauceModifierID > 0 Then
		lsTmp = lsTmp & " " & gsHalf1SauceModifierDescription
	End If
%>
												<div id="edithalf1sauce" style="height: 15px;"><%=lsTmp%></div>
<%
Else
%>
												<div id="edithalf1sauce" style="height: 0px; visibility: hidden;"></div>
<%
End If

If gnHalf2SauceID > 0 And ((gnHalf1SauceID <> gnHalf2SauceID) Or (gnHalf1SauceModifierID <> gnHalf2SauceModifierID)) Then
	lsTmp = "2nd Half Sauce: " & gsHalf2SauceDescription
	If gnHalf2SauceModifierID > 0 Then
		lsTmp = lsTmp & " " & gsHalf2SauceModifierDescription
	End If
%>
												<div id="edithalf2sauce" style="height: 15px;"><%=lsTmp%></div>
<%
Else
%>
												<div id="edithalf2sauce" style="height: 0px; visibility: hidden;"></div>
<%
End If

If ganItemIDs(0) > 0 Then
	For i = 0 To UBound(ganItemIDs)
		Select Case ganHalfIDs(i)
			Case 0
				lsTmp = "Whole Item: " & gasItemShortDescriptions(i)
			Case 1
				lsTmp = "1st Half Item: " & gasItemShortDescriptions(i)
			Case 2
				lsTmp = "2nd Half Item: " & gasItemShortDescriptions(i)
			Case 3
				lsTmp = "On Side Item: " & gasItemShortDescriptions(i)
		End Select
%>
												<div id="edititem<%=i%>" style="height: 15px;"><%=lsTmp%></div>
<%
	Next
Else
%>
												<div id="edititem0" style="height: 0px; visibility: hidden;"></div>
<%
End If
For i = (UBound(ganItemIDs) + 1) To gnMaxItemsPerUnit
%>
												<div id="edititem<%=i%>" style="height: 0px; visibility: hidden;"></div>
<%
Next

If ganTopperIDs(0) > 0 Then
	For i = 0 To UBound(ganTopperIDs)
		Select Case ganTopperHalfIDs(i)
			Case 0
				lsTmp = "Whole Topper: " & gasTopperShortDescriptions(i)
			Case 1
				lsTmp = "1st Half Topper: " & gasTopperShortDescriptions(i)
			Case 2
				lsTmp = "2nd Half Topper: " & gasTopperShortDescriptions(i)
			Case 3
				lsTmp = "On Side Topper: " & gasTopperShortDescriptions(i)
		End Select
%>
												<div id="edittopper<%=i%>" style="height: 15px;"><%=lsTmp%></div>
<%
	Next
Else
%>
												<div id="edittopper0" style="height: 0px; visibility: hidden;"></div>
<%
End If
For i = (UBound(ganTopperIDs) + 1) To gnMaxToppersPerUnit
%>
												<div id="edittopper<%=i%>" style="height: 0px; visibility: hidden;"></div>
<%
Next

If ganFreeSideIDs(0) > 0 Then
	For i = 0 To UBound(ganFreeSideIDs)
		lsTmp = "Free Side: " & gasFreeSideShortDescriptions(i)
%>
												<div id="editfreeside<%=i%>" style="height: 15px;"><%=lsTmp%></div>
<%
	Next
Else
%>
												<div id="editfreeside0" style="height: 0px; visibility: hidden;"></div>
<%
End If
For i = (UBound(ganFreeSideIDs) + 1) To gnMaxFreeSidesPerUnit
%>
												<div id="editfreeside<%=i%>" style="height: 0px; visibility: hidden;"></div>
<%
Next

If ganAddSideIDs(0) > 0 Then
	For i = 0 To UBound(ganAddSideIDs)
		lsTmp = "Add Side: " & gasAddSideShortDescriptions(i)
%>
												<div id="editaddside<%=i%>" style="height: 15px;"><%=lsTmp%></div>
<%
	Next
Else
%>
												<div id="editaddside0" style="height: 0px; visibility: hidden;"></div>
<%
End If
For i = (UBound(ganAddSideIDs) + 1) To gnMaxAddSidesPerUnit
%>
												<div id="editaddside<%=i%>" style="height: 0px; visibility: hidden;"></div>
<%
Next

If Len(gsOrderLineNotes) > 0 Then
	lsTmp = "Notes: " & gsOrderLineNotes
%>
												<div id="editnotes" style="height: 45px;"><%=lsTmp%></div>
<%
Else
%>
												<div id="editnotes" style="height: 0px; visibility: hidden;"></div>
<%
End If
%>
											</div>
										</div>
										<div id="totaldiv" style="width: 320px; text-align: center; background-color: #FFFFFF;">
<%
If gnQuantity > 1 Then
%>
											This Unit: <%=gnQuantity & " @ " & FormatCurrency(gdOrderLineCost - gdOrderLineDiscount)%>
<%
Else
%>
											This Unit: <%=FormatCurrency(gdOrderLineCost - gdOrderLineDiscount)%>
<%
End If
%>
											&nbsp;&nbsp;&nbsp; Total: <%=FormatCurrency(gdOrderTotal + (gnQuantity * (gdOrderLineCost - gdOrderLineDiscount)))%>
										</div>
										<div style="width: 320px; text-align: left;">
											<button style="width: 157px;" onclick="gotoQuantity(false);">#</button>
<%
If Session("SecurityID") > 1 And ((Left(Request.ServerVariables("remote_addr"), 9) = "10.0.254." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.1." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.2.") OR Session("Swipe")) Then
%>
											<button style="width: 157px;" onclick="gotoQuantity(true);"># @ $</button>
<%
Else
%>
											<button style="width: 157px;" onclick="gotoManager(true);"># @ $</button>
<%
End If

If gnOrderLineID > 0 Then
%>
											<button style="width: 157px" onclick="window.location = 'unitselect.asp?dupe=yes&l=<%=gnOrderLineID%>'">Duplicate</button>
											<button style="width: 157px" onclick="gotoDeleteConfirm();">Delete This Unit</button>
<%
Else
%>
											<button style="width: 157px;" onclick="saveOrderLine(true);">Add Another</button>
											<button style="width: 157px; background-color: #C0C0C0;">&nbsp;</button>
<%
End If
%>
											<button style="width: 157px" onclick="window.location = 'unitselect.asp'">Cancel</button>
											<button style="width: 157px" onclick="saveOrderLine(false);">Done</button>
										</div>
									</td>
								</tr>
							</table>
						</div>
						<div id="deleteconfirm" style="position: absolute; left: 0px; top: 0px; width: 1010px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
							<p align="center"><strong>Are you sure you want to 
							delete this unit from the order?</strong><br/><br/>
							<button onclick="window.location = 'unitselect.asp?delete=yes&l=<%=gnOrderLineID%>'">
							Delete</button>&nbsp;&nbsp;&nbsp;&nbsp;<button onclick="gotoUnitEditor();">
							Cancel</button></p>
						</div>
						<div id="managerdiv" style="position: absolute; left: 0px; top: 0px; width: 1010px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
							<table align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="3"><div align="center">
										<strong>MANAGER OVERRIDE REQUIRED</strong></div></td>
								</tr>
								<tr>
									<td colspan="3"><div align="center">
										<input type="password" id="manager" style="width: 200px" autocomplete="off" onkeyup="checkManagerEnterKey();" /></div></td>
								</tr>
								<tr>
									<td><button onclick="cancelManager()">Cancel</button></td>
									<td>&nbsp;</td>
									<td align="right"><button onclick="setManager()">OK</button></td>
								</tr>
							</table>
						</div>
						<div id="quantitydiv" style="position: absolute; left: 0px; top: 0px; width: 1010px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
							<table align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="3"><div align="center">
										<strong>ENTER QUANTITY OF THIS UNIT</strong></div></td>
								</tr>
								<tr>
									<td colspan="3"><div align="center">
										<input type="text" id="quantity" style="width: 200px" autocomplete="off" onkeydown="disableEnterKey();" /></div></td>
								</tr>
								<tr>
									<td><button onclick="addToQuantity('1')">1</button></td>
									<td><button onclick="addToQuantity('2')">2</button></td>
									<td><button onclick="addToQuantity('3')">3</button></td>
								</tr>
								<tr>
									<td><button onclick="addToQuantity('4')">4</button></td>
									<td><button onclick="addToQuantity('5')">5</button></td>
									<td><button onclick="addToQuantity('6')">6</button></td>
								</tr>
								<tr>
									<td><button onclick="addToQuantity('7')">7</button></td>
									<td><button onclick="addToQuantity('8')">8</button></td>
									<td><button onclick="addToQuantity('9')">9</button></td>
								</tr>
								<tr>
									<td><button onclick="cancelQuantity()">Cancel</button></td>
									<td><button onclick="addToQuantity('0')">0</button></td>
									<td><button onclick="backspaceQuantity()">Bksp</button></td>
								</tr>
								<tr>
									<td colspan="3"><button style="width: 235px;" onclick="setQuantity()">OK</button></td>
								</tr>
							</table>
						</div>
						<div id="pricediv" style="position: absolute; left: 0px; top: 0px; width: 1010px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
							<table align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="3"><div align="center">
										<strong>ENTER PRICE OF THIS UNIT</strong></div></td>
								</tr>
								<tr>
									<td colspan="3"><div align="center">
										<input type="text" id="price" style="width: 200px" autocomplete="off" onkeydown="disableEnterKey();" /></div></td>
								</tr>
								<tr>
									<td><button onclick="addToPrice('1')">1</button></td>
									<td><button onclick="addToPrice('2')">2</button></td>
									<td><button onclick="addToPrice('3')">3</button></td>
								</tr>
								<tr>
									<td><button onclick="addToPrice('4')">4</button></td>
									<td><button onclick="addToPrice('5')">5</button></td>
									<td><button onclick="addToPrice('6')">6</button></td>
								</tr>
								<tr>
									<td><button onclick="addToPrice('7')">7</button></td>
									<td><button onclick="addToPrice('8')">8</button></td>
									<td><button onclick="addToPrice('9')">9</button></td>
								</tr>
								<tr>
									<td><button onclick="cancelPrice()">Cancel</button></td>
									<td><button onclick="addToPrice('0')">0</button></td>
									<td><button onclick="backspacePrice()">Bksp</button></td>
								</tr>
								<tr>
									<td colspan="3"><button style="width: 235px;" onclick="setPrice()">OK</button></td>
								</tr>
							</table>
						</div>
						<div id="orderlinenotes" style="position: absolute; left: 0px; top: 0px; width: 1010px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
							<table align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="9"><div align="center">
										<strong>ENTER ADDITIONAL NOTES FOR THIS UNIT</strong></div></td>
								</tr>
								<tr>
									<td colspan="9"><div align="center">
										<textarea id="linenotes" style="width: 930px; height: 60px;"></textarea></div></td>
								</tr>
								<tr>
									<td width="11%"><button style="width: 100px;" onclick="addToNotes('Parmesean on Top')">
									Parmesean on Top</button></td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%"><button style="width: 100px;" onclick="gotoUnitEditor()">Cancel</button></td>
									<td width="11%"><button style="width: 100px;" onclick="clearNotes()">Clear</button></td>
									<td width="11%"><button style="width: 100px;" onclick="saveNotes()">Done</button></td>
								</tr>
								<tr>
									<td width="11%"><button style="width: 100px;" onclick="addToNotes('Cook Lightly')">
									Cook Lightly</button></td>
									<td width="11%"><button style="width: 100px;" onclick="addToNotes('Cook Well Done')">
									Cook Well Done</button></td>
									<td width="11%"><button style="width: 100px;" onclick="addToNotes('Cut in 6')">
									Cut in 6</button></td>
									<td width="11%"><button style="width: 100px;" onclick="addToNotes('Cut in 8')">
									Cut in 8</button></td>
									<td width="11%"><button style="width: 100px;" onclick="addToNotes('Cut in Squares')">
									Cut in Squares</button></td>
									<td width="11%"><button style="width: 100px;" onclick="addToNotes('Don't Cut')">
									Don't Cut</button></td>
									<td width="11%"><button style="width: 100px;" onclick="addToNotes('Items to Edge')">
									Items to Edge</button></td>
									<td width="11%"><button style="width: 100px;" onclick="addToNotes('Put All Items on Top')">
									Put All Items on Top</button></td>
									<td width="11%"><button style="width: 100px;" onclick="addToNotes('Light Cheese')">
									Light Cheese</button></td>
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
<%
For i = 0 To UBound(ganUnitGroupSizeIDs)
%>
						<div id="inclsides<%=i%>" style="position: absolute; left: 0px; top: 0px; width: 1010px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
							<p align="center"><strong>Included Sides</strong></p>
							<div style="position: relative;">
<%
	If ganUnitGroupSideGroupIDs(i, 0) <> 0 Then
		gnPos = 0
		For j = 0 To UBound(ganUnitGroupSideGroupIDs, 2)
			For k = 0 To UBound(ganSideGroupIDs)
				If ganSideGroupIDs(k) = ganUnitGroupSideGroupIDs(i, j) Then
					If ganSideGroupSideIDs(k, 0) <> 0 Then
						For l = 0 To (gadUnitGroupQuantity(i, j) - 1)
%>
								<div id="inclside<%=i%>-<%=gnPos%>" style="float: left; border: 1px #000000 solid; margin: 0px 5px 5px 0px; padding: 2px 2px 2px 2px;">
									<strong>Side #<%=(gnPos + 1)%></strong><br/>
<%
							For m = 0 To UBound(ganSideGroupSideIDs, 2)
								If ganSideGroupSideIDs(k, m) <> 0 Then
									If ganUnitGroupSizeIDs(i) = gnSizeID Then
										If ganFreeSideIDs(gnPos) = ganSideGroupSideIDs(k, m) Then
											lsTmp = "FFFFFF"
										Else
											lsTmp = "000000"
										End If
									Else
										lsTmp = "000000"
									End If
%>
									<button id="inclsidebutton<%=i%>-<%=gnPos%>-<%=ganSideGroupSideIDs(k, m)%>" style="width: 75px; height: 75px; color: #<%=lsTmp%>;" onclick="setInclSide(<%=gnPos%>, <%=ganSideGroupSideIDs(k, m)%>, '<%=gasSideGroupSideDescriptions(k, m)%>', '<%=gasSideGroupSideShortDescriptions(k, m)%>');"><%=gasSideGroupSideShortDescriptions(k, m)%></button>
<%
								End If
							Next
							
							gnPos = gnPos + 1
%>
								</div>
<%
						Next
					End If
					
					Exit For
				End If
			Next
		Next
	End If
%>
							</div>
							<div style="clear: both;">
								<p align="center"><button onclick="gotoUnitEditor();">
								Done</button></p>
							</div>
						</div>
<%
Next

For i = 0 To UBound(ganSideGroupSpecialtyIDs)
	For j = 0 To UBound(ganSideGroupSizeIDs, 2)
%>
						<div id="specialtysides<%=i%>-<%=j%>" style="position: absolute; left: 0px; top: 0px; width: 1010px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
							<p align="center"><strong>Included Sides</strong></p>
							<div style="position: relative;">
<%
		If ganSideGroupSideGroupIDs(i, j, 0) <> 0 Then
			gnPos = 0
			For k = 0 To UBound(ganSideGroupSideGroupIDs, 3)
				For l = 0 To UBound(ganSideGroupIDs)
					If ganSideGroupIDs(l) = ganSideGroupSideGroupIDs(i, j, k) Then
						If ganSideGroupSideIDs(l, 0) <> 0 Then
							For m = 0 To (gadSideGroupQuantity(i, j, k) - 1)
%>
								<div id="specialtyside<%=i%>-<%=j%>-<%=gnPos%>" style="float: left; border: 1px #000000 solid; margin: 0px 5px 5px 0px; padding: 2px 2px 2px 2px;">
									<strong>Side #<%=(gnPos + 1)%></strong><br/>
<%
								For n = 0 To UBound(ganSideGroupSideIDs, 2)
									If ganSideGroupSideIDs(l, n) <> 0 Then
										lsTmp = "000000"
%>
									<button id="specialtysidebutton<%=i%>-<%=j%>-<%=gnPos%>-<%=ganSideGroupSideIDs(l, n)%>" style="width: 75px; height: 75px; color: #<%=lsTmp%>;" onclick="setInclSide(<%=gnPos%>, <%=ganSideGroupSideIDs(l, n)%>, '<%=gasSideGroupSideDescriptions(l, n)%>', '<%=gasSideGroupSideShortDescriptions(l, n)%>');"><%=gasSideGroupSideShortDescriptions(l, n)%></button>
<%
									End If
								Next
								
								gnPos = gnPos + 1
%>
								</div>
<%
							Next
						End If
						
						Exit For
					End If
				Next
			Next
		End If
%>
							</div>
							<div style="clear: both;">
								<p align="center"><button onclick="gotoUnitEditor();">
								Done</button></p>
							</div>
						</div>
<%
	Next
Next
%>
						<div id="messagediv" style="position: absolute; left: 0px; top: 0px; width: 1010px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
							<p align="center">&nbsp;</p>
							<p align="center"><strong>SPECIAL MESSAGE</strong><br/><strong>Please inform the customer of the following:</strong></p>
							<p align="center">&nbsp;</p>
							<p align="center"><strong><span id="specialmessage" name="specialmessage">The special message goes here.</span></strong></p>
							<p align="center">&nbsp;</p>
							<p align="center"><button onclick="gotoUnitEditor();">OK</button></p>
						</div>
						<div id="specialtystylediv" style="position: absolute; left: 0px; top: 0px; width: 1010px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
							<p align="center">&nbsp;</p>
							<p align="center">&nbsp;</p>
							<p align="center"><strong>SPECIAL STYLE</strong><br/><strong>Please ask the customer the following:</strong></p>
							<p align="center">&nbsp;</p>
							<p align="center">&nbsp;</p>
							<p align="center"><strong><span id="specialtystylemessage" name="specialtystylemessage">The specialty style message goes here.</span></strong></p>
							<p align="center">&nbsp;</p>
							<p align="center"><span id="specialtystylebuttons" name="specialtystylebuttons"></span></p>
						</div>
						<div id="needstylediv" style="position: absolute; left: 0px; top: 0px; width: 1010px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
							<p align="center">&nbsp;</p>
							<p align="center"><strong>SELECT STYLE</strong><br/><strong>Please ask the customer which style they prefer:</strong></p>
							<p align="center"><span id="needstylebuttons" name="needstylebuttons"></span></p>
						</div>
						<div id="needtoppersdiv" style="position: absolute; left: 0px; top: 0px; width: 1010px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
							<p align="center">&nbsp;</p>
							<p align="center">&nbsp;</p>
							<p align="center"><strong>SELECT CRUST TOPPERS</strong><br/><strong>Please ask the customer the following:</strong></p>
							<p align="center">&nbsp;</p>
							<p align="center">&nbsp;</p>
							<p align="center"><strong><span id="needtoppersmessage" name="needtoppersmessage">The need toppers message goes here.</span></strong></p>
							<p align="center">&nbsp;</p>
							<p align="center">&nbsp;</p>
							<p align="center">
<%
For i = 0 To UBound(ganUnitTopperIDs)
	If i > 0 Then
%>
&nbsp;&nbsp;
<%
	End If
%>
								<button onclick="addTopper(<%=ganUnitTopperIDs(i)%>); gotoUnitEditor();"><%=gasUnitTopperShortDescriptions(i)%></button>
<%
Next
%>
							</p>
							<p align="center"><button onclick="gbNeedTopperDoneIsDone = true; gotoUnitEditor();">No Toppers</button></p>
						</div>
						<div id="mporeasondiv" style="position: absolute; left: 0px; top: 0px; width: 1010px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
							<table align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="9"><div align="center">
										<strong>WHY IS THE PRICE BEING OVERRIDDEN</strong></div></td>
								</tr>
								<tr>
									<td colspan="9"><div align="center">
										<textarea id="mporeason" style="width: 930px; height: 60px;"></textarea></div></td>
								</tr>
								<tr>
<%
If Len(gasMPOReasons(0)) > 0 Then
	For i = 0 To UBound(gasMPOReasons)
%>
									<td width="11%"><button style="width: 100px;" onclick="addToMPOReason('<%=gasMPOReasons(i)%>')"><%=gasMPOReasons(i)%></button></td>
<%
		If i = 5 Then
			Exit For
		End If
	Next
	
	If UBound(gasMPOReasons) < 5 Then
		For i = 4 To UBound(gasMPOReasons) Step -1
%>
									<td width="11%">&nbsp;</td>
<%
		Next
	End If
Else
%>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
<%
End If
%>
									<td width="11%"><button style="width: 100px;" onclick="cancelMPOReason();">Cancel</button></td>
									<td width="11%"><button style="width: 100px;" onclick="clearMPOReason();">Clear</button></td>
									<td width="11%"><button style="width: 100px;" onclick="saveMPOReason();">Done</button></td>
								</tr>
								<tr>
<%
If UBound(gasMPOReasons) > 5 Then
	For i = 6 To UBound(gasMPOReasons)
%>
									<td width="11%"><button style="width: 100px;" onclick="addToMPOReason('<%=gasMPOReasons(i)%>')"><%=gasMPOReasons(i)%></button></td>
<%
		If i = 14 Then
			Exit For
		End If
	Next
	
	If UBound(gasMPOReasons) < 14 Then
		For i = 13 To UBound(gasMPOReasons) Step -1
%>
									<td width="11%">&nbsp;</td>
<%
		Next
	End If
Else
%>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
<%
End If
%>
								</tr>
							</table>
							<table align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td><div align="center">
										<button onclick="addToMPOReason('+')">+</button><button onclick="addToMPOReason('!')">!</button><button onclick="addToMPOReason('@')">@</button><button onclick="addToMPOReason('#')">#</button><button onclick="addToMPOReason('$')">$</button><button onclick="addToMPOReason('%')">%</button><button onclick="addToMPOReason('^')">^</button><button onclick="addToMPOReason('&')">&amp;</button><button onclick="addToMPOReason('*')">*</button><button onclick="addToMPOReason('(')">(</button><button onclick="addToMPOReason(')')">)</button><button onclick="addToMPOReason(':')">:</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="addToMPOReason('=')">=</button><button onclick="addToMPOReason('1')">1</button><button onclick="addToMPOReason('2')">2</button><button onclick="addToMPOReason('3')">3</button><button onclick="addToMPOReason('4')">4</button><button onclick="addToMPOReason('5')">5</button><button onclick="addToMPOReason('6')">6</button><button onclick="addToMPOReason('7')">7</button><button onclick="addToMPOReason('8')">8</button><button onclick="addToMPOReason('9')">9</button><button onclick="addToMPOReason('0')">0</button><button onclick="addToMPOReason('?')">?</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="addToMPOReason('\'')">'</button><button name="mkey-q" id="mkey-q" onclick="addToMPOReason('Q')">Q</button><button name="mkey-w" id="mkey-w" onclick="addToMPOReason('W')">W</button><button name="mkey-e" id="mkey-e" onclick="addToMPOReason('E')">E</button><button name="mkey-r" id="mkey-r" onclick="addToMPOReason('R')">R</button><button name="mkey-t" id="mkey-t" onclick="addToMPOReason('T')">T</button><button name="mkey-y" id="mkey-y" onclick="addToMPOReason('Y')">Y</button><button name="mkey-u" id="mkey-u" onclick="addToMPOReason('U')">U</button><button name="mkey-i" id="mkey-i" onclick="addToMPOReason('I')">I</button><button name="mkey-o" id="mkey-o" onclick="addToMPOReason('O')">O</button><button name="mkey-p" id="mkey-p" onclick="addToMPOReason('P')">P</button><button onclick="addToMPOReason('/')">/</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="addToMPOReason('.')">.</button><button name="mkey-a" id="mkey-a" onclick="addToMPOReason('A')">A</button><button name="mkey-s" id="mkey-s" onclick="addToMPOReason('S')">S</button><button name="mkey-d" id="mkey-d" onclick="addToMPOReason('D')">D</button><button name="mkey-f" id="mkey-f" onclick="addToMPOReason('F')">F</button><button name="mkey-g" id="mkey-g" onclick="addToMPOReason('G')">G</button><button name="mkey-h" id="mkey-h" onclick="addToMPOReason('H')">H</button><button name="mkey-j" id="mkey-j" onclick="addToMPOReason('J')">J</button><button name="mkey-k" id="mkey-k" onclick="addToMPOReason('K')">K</button><button name="mkey-l" id="mkey-l" onclick="addToMPOReason('L')">L</button><button onclick="addToMPOReason(',')">,</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="shiftMPOReason()">Shift</button><button onclick="addToMPOReason('<')">&lt;</button><button name="mkey-z" id="mkey-z" onclick="addToMPOReason('Z')">Z</button><button name="mkey-x" id="mkey-x" onclick="addToMPOReason('X')">X</button><button name="mkey-c" id="mkey-c" onclick="addToMPOReason('C')">C</button><button name="mkey-v" id="mkey-v" onclick="addToMPOReason('V')">V</button><button name="mkey-b" id="mkey-b" onclick="addToMPOReason('B')">B</button><button name="mkey-n" id="mkey-n" onclick="addToMPOReason('N')">N</button><button name="mkey-m" id="mkey-m" onclick="addToMPOReason('M')">M</button><button onclick="addToMPOReason('>')">&gt;</button><button onclick="properMPOReason()">Proper</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button style="width: 150px;" onclick="addToMPOReason('1/2')">1/2</button>&nbsp;<button style="width: 600px;" onclick="addToMPOReason(' ')">Space</button>&nbsp;<button style="width: 150px;" onclick="backspaceMPOReason()">Bksp</button></div></td>
								</tr>
							</table>
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
If Request("l").Count <> 0 Then
%>
<script type="text/javascript">
<!--
	if (ganFreeSideIDs[0] != 0) {
		for (i = 0; i < ganFreeSideIDs.length; i++) {
			setInclSide(i, ganFreeSideIDs[i], gasFreeSideDescriptions[i], gasFreeSideShortDescriptions[i]);
		}
	}
//-->
</script>
<%
End If

If gbNeedPrinterAlert Then
%>
<script type="text/javascript">
<!--
alert("<%=gsLocalErrorMsg%>\nCHECK PRINTER!");
//-->
</script>
<%
End If

If ganUnitSpecialtyIDs(0) = 0 Then
%>
<script type="text/javascript">
<!--
toggleItems();
//-->
</script>
<%
End If
%>
</body>

</html>
<!-- #Include Virtual="include2/db-disconnect.asp" -->
