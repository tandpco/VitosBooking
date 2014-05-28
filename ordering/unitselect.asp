<%
Option Explicit
Response.buffer = TRUE

If Session("SecurityID") = "" Then
	Response.Redirect("/default.asp")
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
Dim gnOrderID, gnSessionID, gsIPAddress, gnEmpID, gsRefID, gdtTransactionDate, gdtSubmitDate, gdtReleaseDate, gdtExpectedDate, gnStoreID, gnCustomerID, gsCustomerName, gsCustomerPhone, gnAddressID, gnOrderTypeID, gbIsPaid, gnPaymentTypeID, gsPaymentReference, gnAccountID, gdDeliveryCharge, gdDriverMoney, gdTax, gdTax2, gdTip, gnOrderStatusID, gsOrderNotes
Dim gbQuickMode
Dim gsOrderTypeDescription
Dim gsEMail, gsFirstName, gsLastName, gdtBirthdate, gnPrimaryAddressID, gsHomePhone, gsCellPhone, gsWorkPhone, gsFAXPhone, gbIsEMailList, gbIsTextList
Dim gsAddress1, gsAddress2, gsCity, gsState, gsPostalCode, gsAddressNotes
Dim gsAddressDescription, gsCustomerNotes
Dim ganUnitIDs(), gasDescriptions(), gasShortDescriptions()
Dim ganOrderLineIDs(), gasOrderLineDescriptions(), ganQuantity(), gadCost(), gadDiscount()
Dim gdOrderTotal, gdOrderDiscountTotal
Dim gnOrderLineID
Dim i, j, gnTmp, gsTmp1, gsTmp2, gsTmp3, gsTmp4
Dim ganOrderIDs(), gsLocalErrorMsg
Dim gnOpenMon, gnCloseMon, gnOpenTue, gnCloseTue, gnOpenWed, gnCloseWed, gnOpenThu, gnCloseThu, gnOpenFri, gnCloseFri, gnOpenSat, gnCloseSat, gnOpenSun, gnCloseSun
Dim gsHoldDate, gsHoldTime, gbAM, gsPrintTime
Dim gdOriginalPrice, gdManagerPrice, gsMPOReason, gdNewPrice, gsCouponIDs
Dim gnAddQuantity, gdAddPrice
Dim gbConfirmDelivery
Dim gbNeedPrinterAlert
Dim gsFirst, gsLast, gsHome, gsCell, gsWork, gsFAX
Dim gsReturnURL
Dim gasVoidReasons()
Dim gasEditReasons()

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

' Highlight the currently selected tab.
Dim currentTab
currentTab = "order"

gdManagerPrice = 0.00
gbConfirmDelivery = FALSE

If Request("o").Count > 0 Then
	If IsNumeric(Request("o")) Then
		If Request("dupe") = "yes" Then
			If Not DuplicateOrder(CLng(Request("o")), gnOrderID, gsCouponIDs) Then
				Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
			End If
			
			If Not RecalculateOrderPrice(Session("StoreID"), gnOrderID) Then
				Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
			End If
			
			If Len(gsCouponIDs) > 0 Then
				Session("CouponIDs") = gsCouponIDs
				RecalculateOrderDiscounts Session("StoreID"), gnOrderID, Session("CouponIDs")
			End If
			
			If Not RecalculateOrderTax(Session("StoreID"), gnOrderID) Then
				Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
			End If
			
			Session("OrderEdited") = TRUE
			Session("NewOrder") = TRUE
		Else
			gnOrderID = CLng(Request("o"))
			
			If Request("zero") = "yes" Then
				If ZeroOutOrder(gnOrderID, Request("MPOReason")) Then
					Session("OrderEdited") = TRUE
				Else
					Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
				End If
			End If
		End If
		
		Session("OrderID") = gnOrderID
		If GetOrderDetails(gnOrderID, gnSessionID, gsIPAddress, gnEmpID, gsRefID, gdtTransactionDate, gdtSubmitDate, gdtReleaseDate, gdtExpectedDate, gnStoreID, gnCustomerID, gsCustomerName, gsCustomerPhone, gnAddressID, gnOrderTypeID, gbIsPaid, gnPaymentTypeID, gsPaymentReference, gnAccountID, gdDeliveryCharge, gdDriverMoney, gdTax, gdTax2, gdTip, gnOrderStatusID, gsOrderNotes) Then
'			If gnStoreID <> Session("StoreID") Then
'				Response.Redirect("otherstore.asp?t=" & gnOrderTypeID & "&s=" & gnStoreID & "&c=" & gnCustomerID & "&a=" & gnAddressID)
'			End If
			
			If gnOrderStatusID <> 2 And DateValue(gdtTransactionDate) <> DateValue(Session("TransactionDate")) Then
				Response.Redirect("/error.asp?err=" & Server.URLEncode("Order is not From Today"))
			End If
			
' 2013-08-20 TAM: Allow edit order even if complete and paid
'			If gnOrderStatusID >= 10 Then
'				Response.Redirect("/error.asp?err=" & Server.URLEncode("Order is Complete"))
'			End If
'
'			If gbIsPaid Then
'				Response.Redirect("/error.asp?err=" & Server.URLEncode("Order Has Been Paid"))
'			End If
			
			If gnEmpID = 1 And gnOrderStatusID = 1 Then
				Session("OrderEdited") = TRUE
				Session("NewOrder") = TRUE
			End If
			
' Don't replace SessionID or IPAddress
'			Session("SessionID") = gnSessionID
'			Session("IPAddress") = gsIPAddress
			Session("OrderEmpID") = gnEmpID
			Session("SessionID") = Session.SessionID
			Session("IPAddress") = Request.ServerVariables("REMOTE_ADDR")
			gnSessionID = Session("SessionID")
			gsIPAddress = Session("IPAddress")
			Session("RefID") = gsRefID
			Session("SubmitDate") = gdtSubmitDate
			Session("ReleaseDate") = gdtReleaseDate
			Session("ExpectedDate") = gdtExpectedDate
			Session("CustomerID") = gnCustomerID
			Session("CustomerName") = gsCustomerName
			Session("CustomerPhone") = gsCustomerPhone
			Session("AddressID") = gnAddressID
			Session("OrderTypeID") = gnOrderTypeID
			Session("IsPaid") = gbIsPaid
			Session("PaymentTypeID") = gnPaymentTypeID
			Session("PaymentReference") = gsPaymentReference
			Session("AccountID") = gnAccountID
			Session("DeliveryCharge") = gdDeliveryCharge
			Session("DriveMoney") = gdDriverMoney
			Session("Tax") = gdTax
			Session("Tax2") = gdTax2
			Session("Tip") = gdTip
			Session("OrderStatusID") = gnOrderStatusID
			Session("OrderNotes") = gsOrderNotes
			
			gsOrderTypeDescription = GetOrderTypeDescription(gnOrderTypeID)
			Session("OrderTypeDescription") = gsOrderTypeDescription
			
			If GetCustomerDetails(gnCustomerID, gsEMail, gsFirstName, gsLastName, gdtBirthdate, gnPrimaryAddressID, gsHomePhone, gsCellPhone, gsWorkPhone, gsFAXPhone, gbIsEMailList, gbIsTextList) Then
				Session("EMail") = gsEMail
				Session("FirstName") = gsFirstName
				Session("LastName") = gsLastName
				Session("Birthdate") = gdtBirthdate
				Session("PrimaryAddressID") = gnPrimaryAddressID
				Session("HomePhone") = gsHomePhone
				Session("CellPhone") = gsCellPhone
				Session("WorkPhone") = gsWorkPhone
				Session("FAXPhone") = gsFAXPhone
				Session("IsEmailList") = gbIsEMailList
				Session("IsTextList") = gbIsTextList
			Else
				Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
			End If
			
			If GetAddressDetails(gnAddressID, gnTmp, gsAddress1, gsAddress2, gsCity, gsState, gsPostalCode, gsAddressNotes) Then
				Session("Address1") = gsAddress1
				Session("Address2") = gsAddress2
				Session("City") = gsCity
				Session("State") = gsState
				Session("PostalCode") = gsPostalCode
				Session("AddressNotes") = gsAddressNotes
			Else
				Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
			End If
			
			If gnCustomerID = 1 Then
				' Incomplete anonymous online orders won't have an address association
				gsAddressDescription = ""
				gsCustomerNotes = ""
			Else
				If gnAddressID = 1 Then
					gsAddressDescription = ""
					gsCustomerNotes = ""
				Else
					If Not GetCustomerAddressDetails(gnCustomerID, gnAddressID, gsAddressDescription, gsCustomerNotes) Then
						Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
					End If
				End If
			End If
			Session("AddressDescription") = gsAddressDescription
			Session("CustomerNotes") = gsCustomerNotes
			
			If Request("OrderNotes").Count > 0 Then
				gsOrderNotes = Request("OrderNotes")
				If SetOrderNotes(gnOrderID, gsOrderNotes) Then
					Session("OrderNotes") = gsOrderNotes
					Session("OrderEdited") = TRUE
				Else
					Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
				End If
			End If
		Else
			If Len(gsDBErrorMessage) > 0 Then
				Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
			Else
				Response.Redirect("/error.asp?err=" & Server.URLEncode("Invalid Order Specified"))
			End If
		End If
	Else
		Response.Redirect("neworder.asp")
	End If
Else
	gnOrderID = Session("OrderID")
	gnSessionID = Session("SessionID")
	gsIPAddress = Session("IPAddress")
	gnEmpID = Session("OrderEmpID")
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
	
	If Request("q") = "yes" Then
		gbQuickMode = TRUE
		Session("QuickMode") = gbQuickMode
		
		Session("CustomerID") = 1
		Session("EMail") = ""
		Session("FirstName") = ""
		Session("LastName") = ""
		Session("Birthdate") = DateValue("1/1/1900")
		Session("PrimaryAddressID") = 1
		Session("HomePhone") = ""
		Session("CellPhone") = ""
		Session("WorkPhone") = ""
		Session("FAXPhone") = ""
		Session("IsEmailList") = TRUE
		Session("IsTextList") = TRUE
		Session("CustomerName") = ""
		Session("CustomerPhone") = ""
		Session("AddressID") = 1
		Session("Address1") = ""
		Session("Address2") = ""
		Session("City") = ""
		Session("State") = ""
		Session("PostalCode") = ""
		Session("AddressNotes") = ""
		Session("AddressDescription") = ""
		Session("CustomerNotes") = ""
		
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
		gnAddressID = CLng(Session("AddressID"))
		gsAddress1 = Session("Address1")
		gsAddress2 = Session("Address2")
		gsCity = Session("City")
		gsState = Session("State")
		gsPostalCode = Session("PostalCode")
		gsAddressNotes = Session("AddressNotes")
		gsAddressDescription = Session("AddressDescription")
		gsCustomerNotes = Session("CustomerNotes")
		
		If gnOrderID <> 0 Then
			If Not SetOrderCustomer(gnOrderID, gnCustomerID, gsCustomerName, gsCustomerPhone) Then
				Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
			End If
			If Not SetOrderCustomerName(gnOrderID, gsCustomerName) Then
				Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
			End If
			If Not SetOrderAddress(gnOrderID, gnAddressID) Then
				Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
			End If
			
			Session("OrderEdited") = TRUE
		End If
	Else
		gbQuickMode = Session("QuickMode")
	End If
	
	If Request("t").Count > 0 Then
		If IsNumeric(Request("t")) Then
			gnOrderTypeID = CLng(Request("t"))
			Session("OrderTypeID") = gnOrderTypeID
			gsOrderTypeDescription = GetOrderTypeDescription(gnOrderTypeID)
			Session("OrderTypeDescription") = gsOrderTypeDescription
			
			If Request("n").Count > 0 And Request("c").Count = 0 And Request("a").Count = 0 Then
				gdDeliveryCharge = 0.00
				gdDriverMoney = 0.00
				Session("DeliveryCharge") = gdDeliveryCharge
				Session("DriverMoney") = gdDriverMoney
			End If
		Else
			Response.Redirect("neworder.asp")
		End If
	Else
		gnOrderTypeID = CLng(Session("OrderTypeID"))
		gsOrderTypeDescription = Session("OrderTypeDescription")
	End If
	
	If Request("c").Count > 0 Then
		If IsNumeric(Request("c")) Then
			gnCustomerID = CLng(Request("c"))
			Session("CustomerID") = gnCustomerID
			If GetCustomerDetails(gnCustomerID, gsEMail, gsFirstName, gsLastName, gdtBirthdate, gnPrimaryAddressID, gsHomePhone, gsCellPhone, gsWorkPhone, gsFAXPhone, gbIsEMailList, gbIsTextList) Then
				Session("EMail") = gsEMail
				Session("FirstName") = gsFirstName
				Session("LastName") = gsLastName
				Session("Birthdate") = gdtBirthdate
				Session("PrimaryAddressID") = gnPrimaryAddressID
				Session("HomePhone") = gsHomePhone
				Session("CellPhone") = gsCellPhone
				Session("WorkPhone") = gsWorkPhone
				Session("FAXPhone") = gsFAXPhone
				Session("IsEmailList") = gbIsEMailList
				Session("IsTextList") = gbIsTextList
				If Len(gsFirstName) > 0 Then
					If Len(gsLastName) > 0 Then
						gsCustomerName = gsFirstName & " " & gsLastName
					Else
						gsCustomerName = gsFirstName
					End If
				Else
					If Len(gsLastName) > 0 Then
						gsCustomerName = gsLastName
					Else
						gsCustomerName = ""
					End If
				End If
				Session("CustomerName") = gsCustomerName
				
				If gnOrderID <> 0 Then
					If Not SetOrderCustomer(gnOrderID, gnCustomerID, gsCustomerName, gsCustomerPhone) Then
						Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
					End If
					
					Session("OrderEdited") = TRUE
				End If
			Else
				Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
			End If
		Else
			Response.Redirect("neworder.asp")
		End If
	Else
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
	End If
	
	If Request("n").Count > 0 Then
		gsCustomerName = Trim(Request("n"))
		Session("CustomerName") = gsCustomerName
		
		If gnOrderID <> 0 Then
			If Not SetOrderCustomerName(gnOrderID, gsCustomerName) Then
				Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
			End If
			
			Session("OrderEdited") = TRUE
		End If
		
'Response.Write("HERE " & gnCustomerID & " " & Request("c").Count & " " & Request("newcustomer").Count & " " & Len(Session("CustomerPhone")) & " " & Len(Session("CustomerName")) & "<br>")
'Response.End
		If gnCustomerID = 1 And Request("c").Count = 0 And Request("newcustomer").Count = 0 And Len(gsCustomerPhone) > 0 And Len(gsCustomerName) > 0 Then
			gsFirst = ""
			gsLast = ""
			If InStrRev(gsCustomerName, " ") = 0 Then
				gsFirst = gsCustomerName
				gsLast = ""
			Else
				gsFirst = Left(gsCustomerName, (InStrRev(gsCustomerName, " ") - 1))
				gsLast = Mid(gsCustomerName, (InStrRev(gsCustomerName, " ") + 1))
			End If
			
			gnCustomerID = AddCustomerPhoneName(gsFirst, gsLast, gsCustomerPhone)
'Response.Write("HERE " & gnCustomerID & "<br>")
'Response.End
			Session("CustomerID") = gnCustomerID
			If gnCustomerID = 0 Then
				Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
			End If
		End If
	Else
		gsCustomerName = Session("CustomerName")
	End If
	
	If Request("a").Count > 0 Then
		If IsNumeric(Request("a")) Then
			gnAddressID = CLng(Request("a"))
			Session("AddressID") = gnAddressID
			
			If GetAddressDetails(gnAddressID, gnTmp, gsAddress1, gsAddress2, gsCity, gsState, gsPostalCode, gsAddressNotes) Then
				Session("Address1") = gsAddress1
				Session("Address2") = gsAddress2
				Session("City") = gsCity
				Session("State") = gsState
				Session("PostalCode") = gsPostalCode
				Session("AddressNotes") = gsAddressNotes
			Else
				Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
			End If
			
			If gnCustomerID <> 1 And Request("assigncustomeraddress").Count = 0 Then
				If gnAddressID = 1 Then
					gsAddressDescription = ""
					gsCustomerNotes = ""
				Else
					If Not GetCustomerAddressDetails(gnCustomerID, gnAddressID, gsAddressDescription, gsCustomerNotes) Then
						Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
					End If
				End If
				Session("AddressDescription") = gsAddressDescription
				Session("CustomerNotes") = gsCustomerNotes
			Else
				gsAddressDescription = Session("AddressDescription")
				gsCustomerNotes = Session("CustomerNotes")
			End If
			
			If gnOrderTypeID = 1 Then
				gsTmp1 = gsAddress1
				gsTmp2 = gsAddress2
				gsTmp3 = gsCity
				gsTmp4 = gsState
				gnStoreID = GetStoreByAddress(gsPostalCode, gsTmp1, gsTmp2, gsTmp3, gsTmp4, gdDeliveryCharge, gdDriverMoney)
'				If gnStoreID <> Session("StoreID") Then
'					Response.Redirect("otherstore.asp?t=" & gnOrderTypeID & "&s=" & gnStoreID & "&c=" & gnCustomerID & "&a=" & gnAddressID)
'				End If
				
				If gdDeliveryCharge = 0.00 Then
					gdDeliveryCharge = GetDefaultDeliveryCharge(gnStoreID)
				End If
				
				If gdDriverMoney = 0.00 Then
					gdDriverMoney = GetDefaultDriverMoney(gnStoreID)
				End If
			Else
				gdDeliveryCharge = 0.00
				gdDriverMoney = 0.00
			End If
			Session("DeliveryCharge") = gdDeliveryCharge
			Session("DriverMoney") = gdDriverMoney
			
			If gnOrderID <> 0 Then
				If Not SetOrderAddress(gnOrderID, gnAddressID) Then
					Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
				End If
				
				Session("OrderEdited") = TRUE
			End If
		Else
			Response.Redirect("neworder.asp")
		End If
	Else
		gnAddressID = CLng(Session("AddressID"))
		gsAddress1 = Session("Address1")
		gsAddress2 = Session("Address2")
		gsCity = Session("City")
		gsState = Session("State")
		gsPostalCode = Session("PostalCode")
		gsAddressNotes = Session("AddressNotes")
		gsAddressDescription = Session("AddressDescription")
		gsCustomerNotes = Session("CustomerNotes")
	End If
	
	If Request("t").Count > 0 Then
		If gnOrderID <> 0 Then
			If Not SetOrderType(gnOrderID, Session("StoreID"), gnOrderTypeID, gdDeliveryCharge, gdDriverMoney) Then
				Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
			End If
			
			Session("OrderEdited") = TRUE
		End If
	End If
	
	If Request("em").Count > 0 Then
		gsEmail = LCase(Request("em"))
		Session("EMail") = gsEmail
	Else
		gsEmail = Session("EMail")
	End If
End If

If Request("newcustomer") = "yes" Then
	If Request("n").Count = 0 Then
		Response.Redirect("neworder.asp")
	Else
		If Request("h").Count = 0 Then
			Response.Redirect("neworder.asp")
		Else
			If Not IsNumeric(Request("h")) Then
				Response.Redirect("neworder.asp")
			Else
				gsFirst = ""
				gsLast = ""
				If InStrRev(gsCustomerName, " ") = 0 Then
					gsFirst = gsCustomerName
					gsLast = ""
				Else
					gsFirst = Left(gsCustomerName, (InStrRev(gsCustomerName, " ") - 1))
					gsLast = Mid(gsCustomerName, (InStrRev(gsCustomerName, " ") + 1))
				End If
				
				gsHome = ""
				gsCell = ""
				gsWork = ""
				gsFAX = ""
				
				Select Case CLng(Request("h"))
					Case 1
						gsCell = gsCustomerPhone
					Case 2
						gsWork = gsCustomerPhone
					Case 3
						gsFAX = gsCustomerPhone
					Case Else
						gsHome = gsCustomerPhone
				End Select
				
				gnCustomerID = AddCustomer(gsEmail, "", gsFirst, gsLast, DateValue("1/1/1900"), gnAddressID, gsHome, gsCell, gsWork, gsFAX, FALSE, FALSE)
				If gnCustomerID = 0 Then
					Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
				Else
					If Not AddCustomerAddress(gnCustomerID, gnAddressID, "Primary Address") Then
						Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
					Else
						Session("CustomerID") = gnCustomerID
						Session("FirstName") = gsFirst
						Session("LastName") = gsLast
						Session("PrimaryAddressID") = gnAddressID
						Session("HomePhone") = ""
						Session("CellPhone") = ""
						Session("WorkPhone") = ""
						Session("FAXPhone") = ""
						Select Case CLng(Request("h"))
							Case 1
								Session("CellPhone") = gsCustomerPhone
							Case 2
								Session("WorkPhone") = gsCustomerPhone
							Case 3
								Session("FAXPhone") = gsCustomerPhone
							Case Else
								Session("HomePhone") = gsCustomerPhone
						End Select
						
						gsFirstName = Session("FirstName")
						gsLastName = Session("LastName")
						gnPrimaryAddressID = CLng(Session("PrimaryAddressID"))
						gsHomePhone = Session("HomePhone")
						gsCellPhone = Session("CellPhone")
						gsWorkPhone = Session("WorkPhone")
						gsFAXPhone = Session("FAXPhone")
						
						If gnOrderID > 0 Then
							If Not SetOrderCustomer(gnOrderID, gnCustomerID, gsCustomerName, gsCustomerPhone) Then
								Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
							End If
						End If
					End If
				End If
			End If
		End If
	End If
Else
	If Request("assigncustomerphone") = "yes" Then
		If Request("h").Count = 0 Then
			Response.Redirect("neworder.asp")
		Else
			If Not IsNumeric(Request("h")) Then
				Response.Redirect("neworder.asp")
			Else
				If Not AssignCustomerPhone(gnCustomerID, gsCustomerPhone, CLng(Request("h"))) Then
					Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
				End If
			End If
		End If
	Else
		If Request("assigncustomeraddress") = "yes" Then
			If AddCustomerAddress(gnCustomerID, gnAddressID, "Alternate Address") Then
				Session("AddressDescription") = "Alternate Address"
				Session("CustomerNotes") = ""
				gsAddressDescription = Session("AddressDescription")
				gsCustomerNotes = Session("CustomerNotes")
			Else
				Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
			End If
		Else
			If Request("save") = "yes" Then
				If Request("l").Count = 0 Then
					Response.Redirect("neworder.asp")
				Else
					If Not IsNumeric(Request("l")) Then
						Response.Redirect("neworder.asp")
					End If
				End If
				
				If Request("u").Count = 0 Then
					Response.Redirect("neworder.asp")
				Else
					If Not IsNumeric(Request("u")) Then
						Response.Redirect("neworder.asp")
					End If
				End If
				
				If Request("SpecialtyID").Count = 0 Then
					Response.Redirect("neworder.asp")
				Else
					If Not IsNumeric(Request("SpecialtyID")) Then
						Response.Redirect("neworder.asp")
					End If
				End If
				
				If Request("SizeID").Count = 0 Then
					Response.Redirect("neworder.asp")
				Else
					If Not IsNumeric(Request("SizeID")) Then
						Response.Redirect("neworder.asp")
					End If
				End If
				
				If Request("StyleID").Count = 0 Then
					Response.Redirect("neworder.asp")
				Else
					If Not IsNumeric(Request("StyleID")) Then
						Response.Redirect("neworder.asp")
					End If
				End If
				
				If Request("Half1SauceID").Count = 0 Then
					Response.Redirect("neworder.asp")
				Else
					If Not IsNumeric(Request("Half1SauceID")) Then
						Response.Redirect("neworder.asp")
					End If
				End If
				
				If Request("Half2SauceID").Count = 0 Then
					Response.Redirect("neworder.asp")
				Else
					If Not IsNumeric(Request("Half2SauceID")) Then
						Response.Redirect("neworder.asp")
					End If
				End If
				
				If Request("Half1SauceModifierID").Count = 0 Then
					Response.Redirect("neworder.asp")
				Else
					If Not IsNumeric(Request("Half1SauceModifierID")) Then
						Response.Redirect("neworder.asp")
					End If
				End If
				
				If Request("Half2SauceModifierID").Count = 0 Then
					Response.Redirect("neworder.asp")
				Else
					If Not IsNumeric(Request("Half2SauceModifierID")) Then
						Response.Redirect("neworder.asp")
					End If
				End If
				
				If Request("MPOReason").Count <> 0 Then
					gsMPOReason = Trim(Request("MPOReason"))
				End If
				
				If gnOrderID = 0 Then
					gnOrderID = CreateOrder(gnSessionID, gsIPAddress, Session("EmpID"), gsRefID, Session("TransactionDate"), Session("StoreID"), gnCustomerID, gsCustomerName, gsCustomerPhone, gnAddressID, gnOrderTypeID, gdDeliveryCharge, gdDriverMoney, gsOrderNotes)
					If gnOrderID = 0 Then
						Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
					End If
					Session("OrderID") = gnOrderID
					Session("NewOrder") = TRUE
				Else
					If CLng(Request("l")) <> 0 Then
						If Not GetManagerPrice(Request("l"), gdOriginalPrice, gdManagerPrice, gsMPOReason) Then
							Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
						End If
						DeleteOrderLine(Request("l"))
					End If
				End If
				
				If Request("i").Count <> 0 Then
					If IsNumeric(Request("i")) Then
						gnAddQuantity = CLng(Request("i"))
					Else
						gnAddQuantity = 1
					End If
				Else
					gnAddQuantity = 1
				End If
				
				If Request("s").Count <> 0 Then
					If IsNumeric(Request("s")) Then
						gdAddPrice = CDbl(Request("s"))
					Else
						gdAddPrice = 0.00
					End If
				Else
					gdAddPrice = 0.00
				End If
				
				gnOrderLineID = CreateOrderLine(gnOrderID, Request("u"), Request("SpecialtyID"), Request("SizeID"), Request("StyleID"), Request("Half1SauceID"), Request("Half2SauceID"), Request("Half1SauceModifierID"), Request("Half2SauceModifierID"), Request("OrderLineNotes"), gnAddQuantity)
				If gnOrderLineID = 0 Then
					Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
				End If
				
				For i = 1 To Request("ItemID").Count
					If InStr(Request("ItemID")(i), ",") > 0 Then
						If IsNumeric(Left(Request("ItemID")(i), (InStr(Request("ItemID")(i), ",") - 1))) Then
							If IsNumeric(Mid(Request("ItemID")(i), (InStr(Request("ItemID")(i), ",") + 1))) Then
								gnTmp = CreateOrderLineItem(gnOrderLineID, Left(Request("ItemID")(i), (InStr(Request("ItemID")(i), ",") - 1)), Mid(Request("ItemID")(i), (InStr(Request("ItemID")(i), ",") + 1)))
								If gnTmp = 0 Then
									Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
								End If
							End If
						End If
					End If
				Next
				
				For i = 1 To Request("TopperID").Count
					If InStr(Request("TopperID")(i), ",") > 0 Then
						If IsNumeric(Left(Request("TopperID")(i), (InStr(Request("TopperID")(i), ",") - 1))) Then
							If IsNumeric(Mid(Request("TopperID")(i), (InStr(Request("TopperID")(i), ",") + 1))) Then
								gnTmp = CreateOrderLineTopper(gnOrderLineID, Left(Request("TopperID")(i), (InStr(Request("TopperID")(i), ",") - 1)), Mid(Request("TopperID")(i), (InStr(Request("TopperID")(i), ",") + 1)))
								If gnTmp = 0 Then
									Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
								End If
							End If
						End If
					End If
				Next
				
				For i = 1 To Request("FreeSideID").Count
					If IsNumeric(Request("FreeSideID")(i)) Then
						gnTmp = CreateOrderLineSide(gnOrderLineID, Request("FreeSideID")(i), 1)
						If gnTmp = 0 Then
							Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
						End If
					End If
				Next
				
				For i = 1 To Request("AddSideID").Count
					If IsNumeric(Request("AddSideID")(i)) Then
						gnTmp = CreateOrderLineSide(gnOrderLineID, Request("AddSideID")(i), 0)
						If gnTmp = 0 Then
							Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
						End If
					End If
				Next
				
				If Not RecalculateOrderLinePrice(Session("StoreID"), gnOrderLineID, gdNewPrice) Then
					Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
				End If
				
				If gdAddPrice = 0 Then
					If gdManagerPrice <> 0 And gdOriginalPrice = gdNewPrice Then
						If Not SetManagerPriceOverride(gnOrderLineID, gdManagerPrice, gsMPOReason) Then
							Response.Redirect("/error.asp?err=" & Server.URLEncode(gsLocalErrorMsg))
						End If
					End If
				Else
					If Not SetManagerPriceOverride(gnOrderLineID, gdAddPrice, gsMPOReason) Then
						Response.Redirect("/error.asp?err=" & Server.URLEncode(gsLocalErrorMsg))
					End If
				End If
				
				If Len(Session("CouponIDs")) > 0 Then
					RecalculateOrderDiscounts Session("StoreID"), gnOrderID, Session("CouponIDs")
				End If
				
				If Not RecalculateOrderTax(Session("StoreID"), gnOrderID) Then
					Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
				End If
				
' Do not change EmpID now that we are tracking edits
'				If gnEmpID <> 1 Then
'					If SetOrderEmployee(gnOrderID, Session("EmpID")) Then
'						Session("OrderEmpID") = CLng(Session("EmpID"))
'					Else
'						Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
'					End If
'				End If
				
				If GetOrderDetails(gnOrderID, gnSessionID, gsIPAddress, gnEmpID, gsRefID, gdtTransactionDate, gdtSubmitDate, gdtReleaseDate, gdtExpectedDate, gnStoreID, gnCustomerID, gsCustomerName, gsCustomerPhone, gnAddressID, gnOrderTypeID, gbIsPaid, gnPaymentTypeID, gsPaymentReference, gnAccountID, gdDeliveryCharge, gdDriverMoney, gdTax, gdTax2, gdTip, gnOrderStatusID, gsOrderNotes) Then
					Session("Tax") = gdTax
					Session("Tax2") = gdTax2
					
					' Don't replace SessionID or IPAddress
					gnSessionID = Session("SessionID")
					gsIPAddress = Session("IPAddress")
				Else
					If Len(gsDBErrorMessage) > 0 Then
						Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
					Else
						Response.Redirect("/error.asp?err=" & Server.URLEncode("Invalid Order Specified"))
					End If
				End If
				
				Session("OrderEdited") = TRUE
				
				If Request("Another").Count <> 0 Then
					If Request("Another") = "yes" Then
						Response.Redirect("unitedit.asp?u=" & Request("u") & "&s=" & Request("SizeID"))
					End If
				End If
			Else
				If Request("delete") = "yes" Then
					If Request("l").Count = 0 Then
						Response.Redirect("neworder.asp")
					Else
						If IsNumeric(Request("l")) Then
							DeleteOrderLine(Request("l"))
						Else
							Response.Redirect("neworder.asp")
						End If
					End If
					
					If Len(Session("CouponIDs")) > 0 Then
						RecalculateOrderDiscounts Session("StoreID"), gnOrderID, Session("CouponIDs")
					End If
					
					If Not RecalculateOrderTax(Session("StoreID"), gnOrderID) Then
						Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
					End If
					
' Do not change EmpID now that we are tracking edits
'					If gnEmpID <> 1 Then
'						If SetOrderEmployee(gnOrderID, Session("EmpID")) Then
'							Session("OrderEmpID") = CLng(Session("EmpID"))
'						Else
'							Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
'						End If
'					End If
					
					If GetOrderDetails(gnOrderID, gnSessionID, gsIPAddress, gnEmpID, gsRefID, gdtTransactionDate, gdtSubmitDate, gdtReleaseDate, gdtExpectedDate, gnStoreID, gnCustomerID, gsCustomerName, gsCustomerPhone, gnAddressID, gnOrderTypeID, gbIsPaid, gnPaymentTypeID, gsPaymentReference, gnAccountID, gdDeliveryCharge, gdDriverMoney, gdTax, gdTax2, gdTip, gnOrderStatusID, gsOrderNotes) Then
						' Don't replace SessionID or IPAddress
						gnSessionID = Session("SessionID")
						gsIPAddress = Session("IPAddress")
						
						Session("Tax") = gdTax
						Session("Tax2") = gdTax2
					Else
						If Len(gsDBErrorMessage) > 0 Then
							Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
						Else
							Response.Redirect("/error.asp?err=" & Server.URLEncode("Invalid Order Specified"))
						End If
					End If
					
					Session("OrderEdited") = TRUE
				Else
					If Request("dupe") = "yes" And Request("l").Count <> 0 Then
						If Not IsNumeric(Request("l")) Then
							Response.Redirect("neworder.asp")
						End If
						
						If Not DuplicateOrderLine(gnOrderID, Request("l"), gnOrderLineID) Then
							Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
						End If
						
						If Not RecalculateOrderLinePrice(Session("StoreID"), gnOrderLineID, gdNewPrice) Then
							Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
						End If
						
						If Len(Session("CouponIDs")) > 0 Then
							RecalculateOrderDiscounts Session("StoreID"), gnOrderID, Session("CouponIDs")
						End If
						
						If Not RecalculateOrderTax(Session("StoreID"), gnOrderID) Then
							Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
						End If
						
' Do not change EmpID now that we are tracking edits
'						If gnEmpID <> 1 Then
'							If SetOrderEmployee(gnOrderID, Session("EmpID")) Then
'								Session("OrderEmpID") = CLng(Session("EmpID"))
'							Else
'								Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
'							End If
'						End If
						
						If GetOrderDetails(gnOrderID, gnSessionID, gsIPAddress, gnEmpID, gsRefID, gdtTransactionDate, gdtSubmitDate, gdtReleaseDate, gdtExpectedDate, gnStoreID, gnCustomerID, gsCustomerName, gsCustomerPhone, gnAddressID, gnOrderTypeID, gbIsPaid, gnPaymentTypeID, gsPaymentReference, gnAccountID, gdDeliveryCharge, gdDriverMoney, gdTax, gdTax2, gdTip, gnOrderStatusID, gsOrderNotes) Then
							' Don't replace SessionID or IPAddress
							gnSessionID = Session("SessionID")
							gsIPAddress = Session("IPAddress")
							
							Session("Tax") = gdTax
							Session("Tax2") = gdTax2
						Else
							If Len(gsDBErrorMessage) > 0 Then
								Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
							Else
								Response.Redirect("/error.asp?err=" & Server.URLEncode("Invalid Order Specified"))
							End If
						End If
						
						Session("OrderEdited") = TRUE
					End If
				End If	
			End If
		End If
	End If
End If

If Not GetUnits(Session("StoreID"), ganUnitIDs, gasDescriptions, gasShortDescriptions) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If Not GetOrderLines(gnOrderID, ganOrderLineIDs, gasOrderLineDescriptions, ganQuantity, gadCost, gadDiscount) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If ganOrderLineIDs(0) = 0 Then
	Session("OrderLineCount") = 0
Else
	Session("OrderLineCount") = (UBound(ganOrderLineIDs) + 1)
End If

gdOrderTotal = 0.00
gdOrderDiscountTotal = 0.00
If ganOrderLineIDs(0) <> 0 Then
	For i = 0 To UBound(ganOrderLineIDs)
		gdOrderTotal = gdOrderTotal + (ganQuantity(i) * gadCost(i))
		gdOrderDiscountTotal = gdOrderDiscountTotal + (ganQuantity(i) * gadDiscount(i))
	Next
End If
gdOrderTotal = gdOrderTotal + gdDeliveryCharge + gdTax + gdTax2 + gdTip
Session("OrderTotal") = (gdOrderTotal - gdOrderDiscountTotal)

If Not GetStoreHours(Session("StoreID"), gnOpenMon, gnCloseMon, gnOpenTue, gnCloseTue, gnOpenWed, gnCloseWed, gnOpenThu, gnCloseThu, gnOpenFri, gnCloseFri, gnOpenSat, gnCloseSat, gnOpenSun, gnCloseSun) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If gnOrderStatusID = 2 Then
	If Month(gdtExpectedDate) < 10 Then
		gsHoldDate = "0" & Month(gdtExpectedDate)
	Else
		gsHoldDate = Month(gdtExpectedDate)
	End If
	If Day(gdtExpectedDate) < 10 Then
		gsHoldDate = gsHoldDate & "0" & Day(gdtExpectedDate)
	Else
		gsHoldDate = gsHoldDate & Day(gdtExpectedDate)
	End If
	gsHoldDate = gsHoldDate & Right(Year(gdtExpectedDate), 2)
	If Hour(gdtExpectedDate) > 21 Then
		gsHoldTime = Hour(gdtExpectedDate) - 12
		gbAM = FALSE
	Else
		If Hour(gdtExpectedDate) > 12 Then
			gsHoldTime = "0" & (Hour(gdtExpectedDate) - 12)
			gbAM = FALSE
		Else
			If Hour(gdtExpectedDate) < 10 Then
				gsHoldTime = "0" & Hour(gdtExpectedDate)
			Else
				gsHoldTime = Hour(gdtExpectedDate)
			End If
			gbAM = TRUE
		End If
	End If
	If Minute(gdtExpectedDate) < 10 Then
		gsHoldTime = gsHoldTime & "0" & Minute(gdtExpectedDate)
	Else
		gsHoldTime = gsHoldTime & Minute(gdtExpectedDate)
	End If
	
	gsPrintTime = DateDiff("n", gdtReleaseDate, gdtExpectedDate)
Else
	gbAM = TRUE
	gsHoldDate = Replace(CStr(now), "/", "")
	If Month(now) < 10 Then
		If Day(now) < 10 Then
			gsHoldDate = "0" & Left(gsHoldDate, 1) & "0" & Mid(gsHoldDate, 2, 1) & Mid(gsHoldDate, 5, 2)
		Else
			gsHoldDate = "0" & Left(gsHoldDate, 3) & Mid(gsHoldDate, 6, 2)
		End If
	Else
		If Day(now) < 10 Then
			gsHoldDate = Left(gsHoldDate, 2) & "0" & Mid(gsHoldDate, 3, 1) & Mid(gsHoldDate, 6, 2)
		Else
			gsHoldDate = Left(gsHoldDate, 4) & Mid(gsHoldDate, 7, 2)
		End If
	End If
	gsHoldTime = ""
	If gnOrderTypeID = 1 Then
		gsPrintTime = "30"
	Else
		gsPrintTime = "15"
	End If
End If

If gnOrderTypeID = 1 Then
	gbConfirmDelivery = CheckExtraDelivery(gnStoreID, gnCustomerID, gnAddressID, gdDeliveryCharge)
End If

If Not GetVoidReasons(gasVoidReasons) Then
	Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
End If

If Not GetEditReasons(gasEditReasons) Then
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
var gsOrderNotes = "<%=gsOrderNotes%>";
var gnOrderID = <%=gnOrderID%>;
var gbAM = <%=LCase(gbAM)%>;
var gbNotesInLowerCase = false;
var gbVoidReasonInLowerCase = false;
var gbEditReasonInLowerCase = false;
var gsEditReason = "";
var gbEditGotoPayment = false;

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

function cancelOrder() {
	var loDiv;
	
<%
If Session("OrderLineCount") = 0 Then
%>
	window.location = "neworder.asp?cancel=yes";
<%
Else
%>
	loDiv = ie4? eval("document.all.unitselector") : document.getElementById("unitselector");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.ordernotes") : document.getElementById("ordernotes");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.holdorder") : document.getElementById("holdorder");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmholdorder") : document.getElementById("confirmholdorder");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById("confirmreprint");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.calendar") : document.getElementById("calendar");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmdelivery") : document.getElementById("confirmdelivery");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.voidreasondiv") : document.getElementById("voidreasondiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.editreasondiv") : document.getElementById("editreasondiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmcancel") : document.getElementById("confirmcancel");
	loDiv.style.visibility = "visible";
	
	resetRedirect();
<%
End If
%>
}

function gotoUnitSelector() {
	var loDiv;
	
	loDiv = ie4? eval("document.all.ordernotes") : document.getElementById("ordernotes");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.holdorder") : document.getElementById("holdorder");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmholdorder") : document.getElementById("confirmholdorder");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmcancel") : document.getElementById("confirmcancel");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById("confirmreprint");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.calendar") : document.getElementById("calendar");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmdelivery") : document.getElementById("confirmdelivery");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.voidreasondiv") : document.getElementById("voidreasondiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.editreasondiv") : document.getElementById("editreasondiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.unitselector") : document.getElementById("unitselector");
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function gotoOrderNotes() {
	var loNotes, loDiv;
	
	loNotes = ie4? eval("document.all.notes") : document.getElementById('notes');
	loNotes.value = gsOrderNotes;
	
	loDiv = ie4? eval("document.all.unitselector") : document.getElementById("unitselector");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.holdorder") : document.getElementById("holdorder");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmholdorder") : document.getElementById("confirmholdorder");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmcancel") : document.getElementById("confirmcancel");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById("confirmreprint");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.calendar") : document.getElementById("calendar");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmdelivery") : document.getElementById("confirmdelivery");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.voidreasondiv") : document.getElementById("voidreasondiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.editreasondiv") : document.getElementById("editreasondiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.ordernotes") : document.getElementById("ordernotes");
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function addToNotes(psDigit) {
	var loNotes, lsNotes;
	
	loNotes = ie4? eval("document.all.notes") : document.getElementById('notes');
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
	
	loNotes = ie4? eval("document.all.notes") : document.getElementById('notes');
	lsNotes = loNotes.value;
	if (lsNotes.length > 0) {
		lsNotes = lsNotes.substr(0, (lsNotes.length - 1));
		loNotes.value = lsNotes;
	}
	
	resetRedirect();
}

function clearNotes() {
	var loNotes;
	
	loNotes = ie4? eval("document.all.notes") : document.getElementById('notes');
	loNotes.value = "";
	
	resetRedirect();
}

function properNotes() {
	var loNotes, i, lsText, lbDoUpper;
	
	loNotes = ie4? eval("document.all.notes") : document.getElementById('notes');
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
	var loNotes;
	
	loNotes = ie4? eval("document.all.notes") : document.getElementById('notes');
	gsOrderNotes = loNotes.value;
	console.log("unitselect.asp?o=" + gnOrderID.toString() + "&OrderNotes=" + encodeURIComponent(gsOrderNotes))
	window.location = "unitselect.asp?o=" + gnOrderID.toString() + "&OrderNotes=" + encodeURIComponent(gsOrderNotes);
}

function gotoHoldOrder() {
	var loDiv;
	
	loDiv = ie4? eval("document.all.unitselector") : document.getElementById("unitselector");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.ordernotes") : document.getElementById("ordernotes");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmholdorder") : document.getElementById("confirmholdorder");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmcancel") : document.getElementById("confirmcancel");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById("confirmreprint");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.calendar") : document.getElementById("calendar");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmdelivery") : document.getElementById("confirmdelivery");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.voidreasondiv") : document.getElementById("voidreasondiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.editreasondiv") : document.getElementById("editreasondiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.holdorder") : document.getElementById("holdorder");
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function cancelHoldOrder() {
	var loField, loAMPMBtn, loDiv;
	
	loField = ie4? eval("document.all.date") : document.getElementById('date');
	loField.value = "<%=gsHoldDate%>";
	
	loField = ie4? eval("document.all.time") : document.getElementById('time');
	loField.value = "<%=gsHoldTime%>";
	
	gbAM = <%=LCase(gbAM)%>;
	loAMPMBtn = ie4? eval("document.all.ampm") : document.getElementById('ampm');
	if (<%=LCase(gbAM)%>) {
		loAMPMBtn.innerHTML = "AM";
	}
	else {
		loAMPMBtn.innerHTML = "PM";
	}
	
	loDiv = ie4? eval("document.all.ordernotes") : document.getElementById("ordernotes");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.holdorder") : document.getElementById("holdorder");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmholdorder") : document.getElementById("confirmholdorder");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmcancel") : document.getElementById("confirmcancel");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById("confirmreprint");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.calendar") : document.getElementById("calendar");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmdelivery") : document.getElementById("confirmdelivery");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.voidreasondiv") : document.getElementById("voidreasondiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.editreasondiv") : document.getElementById("editreasondiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.unitselector") : document.getElementById("unitselector");
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function addToDate(psDigit) {
	var loDate, lsDate;
	
	loDate = ie4? eval("document.all.date") : document.getElementById('date');
	lsDate = loDate.value;
	
	lsDate += psDigit;
	
	loDate.value = lsDate;
	
	resetRedirect();
}

function backspaceDate() {
	var loText, lsText;
	
	loText = ie4? eval("document.all.date") : document.getElementById('date');
	lsText = loText.value;
	if (lsText.length > 0) {
		lsText = lsText.substr(0, (lsText.length - 1));
		loText.value = lsText;
	}
	
	resetRedirect();
}

function previousDate() {
	var loText, lsText, ldtDate, ldtDate2;
	
	loText = ie4? eval("document.all.date") : document.getElementById('date');
	lsText = loText.value;
	if (lsText.length == 6) {
		ldtDate = new Date(parseInt("20" + lsText.substr(4, 2), 10), (parseInt(lsText.substr(0, 2), 10) - 1), parseInt(lsText.substr(2, 2), 10));
		ldtDate2 = new Date();
		ldtDate2.setDate(ldtDate.getDate() - 1);
		if (ldtDate2.getMonth() >= 9) {
			lsText = (ldtDate2.getMonth() + 1).toString();
		}
		else {
			lsText = "0" + (ldtDate2.getMonth() + 1).toString();
		}
		if (ldtDate2.getDate() > 9) {
			lsText += ldtDate2.getDate().toString();
		}
		else {
			lsText += "0" + ldtDate2.getDate().toString();
		}
		lsText += ldtDate2.getFullYear().toString().substr(2);
		loText.value = lsText;
	}
	
	resetRedirect();
}

function nextDate() {
	var loText, lsText, lsMonth, lsDay, ldtDate, ldtDate2;
	
	loText = ie4? eval("document.all.date") : document.getElementById('date');
	lsText = loText.value;
	if (lsText.length == 6) {
		if (lsText.substr(0, 1) == "0") {
			lsMonth = lsText.substr(1, 1);
		}
		else {
			lsMonth = lsText.substr(0, 2);
		}
		if (lsText.substr(2, 1) == "0") {
			lsDay = lsText.substr(3, 1);
		}
		else {
			lsDay = lsText.substr(2, 2);
		}
		ldtDate = new Date(parseInt("20" + lsText.substr(4, 2), 10), (parseInt(lsMonth, 10) - 1), parseInt(lsDay, 10));
		ldtDate2 = new Date();
		ldtDate2.setDate(ldtDate.getDate() + 1);
		if (ldtDate2.getMonth() >= 9) {
			lsText = (ldtDate2.getMonth() + 1).toString();
		}
		else {
			lsText = "0" + (ldtDate2.getMonth() + 1).toString();
		}
		if (ldtDate2.getDate() > 9) {
			lsText += ldtDate2.getDate().toString();
		}
		else {
			lsText += "0" + ldtDate2.getDate().toString();
		}
		lsText += ldtDate2.getFullYear().toString().substr(2);
		loText.value = lsText;
	}
	
	resetRedirect();
}

function addToTime(psDigit) {
	var loTime, lsTime;
	
	loTime = ie4? eval("document.all.time") : document.getElementById('time');
	lsTime = loTime.value;
	
	lsTime += psDigit;
	
	loTime.value = lsTime;
	
	resetRedirect();
}

function backspaceTime() {
	var loText, lsText;
	
	loText = ie4? eval("document.all.time") : document.getElementById('time');
	lsText = loText.value;
	if (lsText.length > 0) {
		lsText = lsText.substr(0, (lsText.length - 1));
		loText.value = lsText;
	}
	
	resetRedirect();
}

function addToMinutes(psDigit) {
	var loTime, lsTime;
	
	loTime = ie4? eval("document.all.minutes") : document.getElementById('minutes');
	lsTime = loTime.value;
	
	lsTime += psDigit;
	
	loTime.value = lsTime;
	
	resetRedirect();
}

function backspaceMinutes() {
	var loText, lsText;
	
	loText = ie4? eval("document.all.minutes") : document.getElementById('minutes');
	lsText = loText.value;
	if (lsText.length > 0) {
		lsText = lsText.substr(0, (lsText.length - 1));
		loText.value = lsText;
	}
	
	resetRedirect();
}

function toggleAMPM() {
	var loAMPMBtn;
	
	gbAM = !gbAM;
	
	loAMPMBtn = ie4? eval("document.all.ampm") : document.getElementById('ampm');
	if (gbAM) {
		loAMPMBtn.innerHTML = "AM";
	}
	else {
		loAMPMBtn.innerHTML = "PM";
	}
}

function gotoConfirmHoldOrder() {
	var loDiv;
	
	loDiv = ie4? eval("document.all.unitselector") : document.getElementById("unitselector");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.ordernotes") : document.getElementById("ordernotes");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.holdorder") : document.getElementById("holdorder");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmcancel") : document.getElementById("confirmcancel");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById("confirmreprint");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.calendar") : document.getElementById("calendar");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmdelivery") : document.getElementById("confirmdelivery");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.voidreasondiv") : document.getElementById("voidreasondiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.editreasondiv") : document.getElementById("editreasondiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmholdorder") : document.getElementById("confirmholdorder");
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function saveHoldOrder(pbConfirm) {
	var loDate, loTime, ldtDate, ldtNow;
	var lnOpenMon, lnCloseMon, lnOpenTue, lnCloseTue, lnOpenWed, lnCloseWed, lnOpenThu, lnCloseThu, lnOpenFri, lnCloseFri, lnOpenSat, lnCloseSat, lnOpenSun, lnCloseSun;
	var lnTime, lbTimeOK;
	var lsLocation, loNotes, loMinutes;
	
	loMinutes = ie4? eval("document.all.minutes") : document.getElementById('minutes');
	if (loMinutes.value.length == 0) {
		return false;
	}
	if (parseInt(loMinutes.value, 10) == 0) {
		return false;
	}
	loDate = ie4? eval("document.all.date") : document.getElementById('date');
	if (loDate.value.length != 6) {
		return false;
	}
	if ((parseInt(loDate.value.substr(0, 2), 10) == 0) || (parseInt(loDate.value.substr(0, 2), 10) > 12)) {
		return false;
	}
	switch (parseInt(loDate.value.substr(0, 2), 10)) {
		case 1:
		case 3:
		case 5:
		case 7:
		case 8:
		case 10:
		case 12:
			if ((parseInt(loDate.value.substr(2, 2), 10) == 0) || (parseInt(loDate.value.substr(2, 2), 10) > 31)) {
				return false;
			}
			break;
		case 2:
			if (((parseInt(loDate.value.substr(0, 2), 10) % 4) == 0) && (((parseInt(loDate.value.substr(0, 2), 10) % 100) != 0) || ((parseInt(loDate.value.substr(0, 2), 10) % 400) == 0))) {
				if ((parseInt(loDate.value.substr(2, 2), 10) == 0) || (parseInt(loDate.value.substr(2, 2), 10) > 29)) {
					return false;
				}
			}
			else {
				if ((parseInt(loDate.value.substr(2, 2), 10) == 0) || (parseInt(loDate.value.substr(2, 2), 10) > 28)) {
					return false;
				}
			}
			break;
		case 4:
		case 6:
		case 9:
		case 11:
			if ((parseInt(loDate.value.substr(2, 2), 10) == 0) || (parseInt(loDate.value.substr(2, 2), 10) > 30)) {
				return false;
			}
			break;
	}
	
	loTime = ie4? eval("document.all.time") : document.getElementById('time');
	if (loTime.value.length != 4) {
		return false;
	}
	if ((parseInt(loTime.value.substr(0, 2), 10) == 0) || (parseInt(loTime.value.substr(0, 2), 10) > 12)) {
		return false;
	}
	if (parseInt(loTime.value.substr(2, 2), 10) > 59) {
		return false;
	}
	
	if (gbAM) {
		if (parseInt(loTime.value.substr(0, 2), 10) == 12) {
			ldtDate = new Date(parseInt("20" + loDate.value.substr(4, 2), 10), (parseInt(loDate.value.substr(0, 2), 10) - 1), parseInt(loDate.value.substr(2, 2), 10), 0, parseInt(loTime.value.substr(2, 2), 10));
		}
		else {
			ldtDate = new Date(parseInt("20" + loDate.value.substr(4, 2), 10), (parseInt(loDate.value.substr(0, 2), 10) - 1), parseInt(loDate.value.substr(2, 2), 10), parseInt(loTime.value.substr(0, 2), 10), parseInt(loTime.value.substr(2, 2), 10));
		}
	}
	else {
		ldtDate = new Date(parseInt("20" + loDate.value.substr(4, 2), 10), (parseInt(loDate.value.substr(0, 2), 10) - 1), parseInt(loDate.value.substr(2, 2), 10), (parseInt(loTime.value.substr(0, 2), 10) + 12), parseInt(loTime.value.substr(2, 2), 10));
	}
	ldtNow = new Date();
	if (ldtDate.valueOf() <= ldtNow.valueOf()) {
		return false;
	}
	
	if (pbConfirm) {
		lbTimeOK = true;
	}
	else {
		lnOpenMon = <%=gnOpenMon%>;
		lnCloseMon = <%=gnCloseMon%>;
		lnOpenTue = <%=gnOpenTue%>;
		lnCloseTue = <%=gnCloseTue%>;
		lnOpenWed = <%=gnOpenWed%>;
		lnCloseWed = <%=gnCloseWed%>;
		lnOpenThu = <%=gnOpenThu%>;
		lnCloseThu = <%=gnCloseThu%>;
		lnOpenFri = <%=gnOpenFri%>;
		lnCloseFri = <%=gnCloseFri%>;
		lnOpenSat = <%=gnOpenSat%>;
		lnCloseSat = <%=gnCloseSat%>;
		lnOpenSun = <%=gnOpenSun%>;
		lnCloseSun = <%=gnCloseSun%>;
		
		lnTime = ldtDate.getHours() * 100 + ldtDate.getMinutes();
		lbTimeOK = false;
		switch (ldtDate.getDay()) {
			case 0:
				if (lnCloseSun < lnOpenSun) {
					if (((lnTime >= lnOpenSun) && (lnTime <= 2400)) || ((lnTime >= 0) && (lnTime <= lnCloseSun))) {
						lbTimeOK = true;
					}
				}
				else {
					if ((lnTime >= lnOpenSun) && (lnTime <= lnCloseSun)) {
						lbTimeOK = true;
					}
				}
				break;
			case 1:
				if (lnCloseMon < lnOpenMon) {
					if (((lnTime >= lnOpenMon) && (lnTime <= 2400)) || ((lnTime >= 0) && (lnTime <= lnCloseMon))) {
						lbTimeOK = true;
					}
				}
				else {
					if ((lnTime >= lnOpenMon) && (lnTime <= lnCloseMon)) {
						lbTimeOK = true;
					}
				}
				break;
			case 2:
				if (lnCloseTue < lnOpenTue) {
					if (((lnTime >= lnOpenTue) && (lnTime <= 2400)) || ((lnTime >= 0) && (lnTime <= lnCloseTue))) {
						lbTimeOK = true;
					}
				}
				else {
					if ((lnTime >= lnOpenTue) && (lnTime <= lnCloseTue)) {
						lbTimeOK = true;
					}
				}
				break;
			case 3:
				if (lnCloseWed < lnOpenWed) {
					if (((lnTime >= lnOpenWed) && (lnTime <= 2400)) || ((lnTime >= 0) && (lnTime <= lnCloseWed))) {
						lbTimeOK = true;
					}
				}
				else {
					if ((lnTime >= lnOpenWed) && (lnTime <= lnCloseWed)) {
						lbTimeOK = true;
					}
				}
				break;
			case 4:
				if (lnCloseThu < lnOpenThu) {
					if (((lnTime >= lnOpenThu) && (lnTime <= 2400)) || ((lnTime >= 0) && (lnTime <= lnCloseThu))) {
						lbTimeOK = true;
					}
				}
				else {
					if ((lnTime >= lnOpenThu) && (lnTime <= lnCloseThu)) {
						lbTimeOK = true;
					}
				}
				break;
			case 5:
				if (lnCloseFri < lnOpenFri) {
					if (((lnTime >= lnOpenFri) && (lnTime <= 2400)) || ((lnTime >= 0) && (lnTime <= lnCloseFri))) {
						lbTimeOK = true;
					}
				}
				else {
					if ((lnTime >= lnOpenFri) && (lnTime <= lnCloseFri)) {
						lbTimeOK = true;
					}
				}
				break;
			case 6:
				if (lnCloseSat < lnOpenSat) {
					if (((lnTime >= lnOpenSat) && (lnTime <= 2400)) || ((lnTime >= 0) && (lnTime <= lnCloseSat))) {
						lbTimeOK = true;
					}
				}
				else {
					if ((lnTime >= lnOpenSat) && (lnTime <= lnCloseSat)) {
						lbTimeOK = true;
					}
				}
				break;
		}
	}
	
	if (lbTimeOK) {
		lsLocation = "neworder.asp?";
		
		loNotes = ie4? eval("document.all.notes") : document.getElementById('notes');
		if (loNotes.value.length > 0) {
			lsLocation = lsLocation + "OrderNotes=" + loNotes.value + "&";
		}
		
		lsLocation = lsLocation + "d=" + loMinutes.value;
		
		lsLocation = lsLocation + "&e=" + loDate.value + "&g=" + loTime.value;
		if (gbAM) {
			lsLocation = lsLocation + "1";
		}
		else {
			lsLocation = lsLocation + "2";
		}
		
		window.location = lsLocation;
	}
	else {
		gotoConfirmHoldOrder();
	}
}

function gotoConfirmReprint() {
	var loDiv;
	
	loDiv = ie4? eval("document.all.unitselector") : document.getElementById("unitselector");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.ordernotes") : document.getElementById("ordernotes");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmholdorder") : document.getElementById("confirmholdorder");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmcancel") : document.getElementById("confirmcancel");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.holdorder") : document.getElementById("holdorder");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.calendar") : document.getElementById("calendar");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmdelivery") : document.getElementById("confirmdelivery");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.voidreasondiv") : document.getElementById("voidreasondiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.editreasondiv") : document.getElementById("editreasondiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById("confirmreprint");
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function gotoVoidReason() {
	var loNotes, loDiv;
	
	loNotes = ie4? eval("document.all.voidreason") : document.getElementById('voidreason');
	loNotes.value = "";
	
	loDiv = ie4? eval("document.all.unitselector") : document.getElementById("unitselector");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.ordernotes") : document.getElementById("ordernotes");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmholdorder") : document.getElementById("confirmholdorder");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmcancel") : document.getElementById("confirmcancel");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.holdorder") : document.getElementById("holdorder");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.calendar") : document.getElementById("calendar");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmdelivery") : document.getElementById("confirmdelivery");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById("confirmreprint");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.editreasondiv") : document.getElementById("editreasondiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.voidreasondiv") : document.getElementById("voidreasondiv");
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function addToVoidReason(psDigit) {
	var loNotes, lsNotes;
	
	loNotes = ie4? eval("document.all.voidreason") : document.getElementById('voidreason');
	lsNotes = loNotes.value;
	
	if (psDigit.length > 1) {
		if (lsNotes.length > 0) {
			lsNotes = lsNotes + " ";
		}
	}
	
	if (gbVoidReasonInLowerCase) {
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

function backspaceVoidReason() {
	var loNotes, lsNotes;
	
	loNotes = ie4? eval("document.all.voidreason") : document.getElementById('voidreason');
	lsNotes = loNotes.value;
	if (lsNotes.length > 0) {
		lsNotes = lsNotes.substr(0, (lsNotes.length - 1));
		loNotes.value = lsNotes;
	}
	
	resetRedirect();
}

function clearVoidReason() {
	var loNotes;
	
	loNotes = ie4? eval("document.all.voidreason") : document.getElementById('voidreason');
	loNotes.value = "";
	
	resetRedirect();
}

function properVoidReason() {
	var loNotes, i, lsText, lbDoUpper;
	
	loNotes = ie4? eval("document.all.voidreason") : document.getElementById('voidreason');
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

function shiftVoidReason() {
	var loObj;
	
	if (gbVoidReasonInLowerCase) {
		loObj = ie4? eval("document.all.vkey-q") : document.getElementById('vkey-q');
		loObj.innerHTML = "Q";
		loObj = ie4? eval("document.all.vkey-w") : document.getElementById('vkey-w');
		loObj.innerHTML = "W";
		loObj = ie4? eval("document.all.vkey-e") : document.getElementById('vkey-e');
		loObj.innerHTML = "E";
		loObj = ie4? eval("document.all.vkey-r") : document.getElementById('vkey-r');
		loObj.innerHTML = "R";
		loObj = ie4? eval("document.all.vkey-t") : document.getElementById('vkey-t');
		loObj.innerHTML = "T";
		loObj = ie4? eval("document.all.vkey-y") : document.getElementById('vkey-y');
		loObj.innerHTML = "Y";
		loObj = ie4? eval("document.all.vkey-u") : document.getElementById('vkey-u');
		loObj.innerHTML = "U";
		loObj = ie4? eval("document.all.vkey-i") : document.getElementById('vkey-i');
		loObj.innerHTML = "I";
		loObj = ie4? eval("document.all.vkey-o") : document.getElementById('vkey-o');
		loObj.innerHTML = "O";
		loObj = ie4? eval("document.all.vkey-p") : document.getElementById('vkey-p');
		loObj.innerHTML = "P";
		loObj = ie4? eval("document.all.vkey-a") : document.getElementById('vkey-a');
		loObj.innerHTML = "A";
		loObj = ie4? eval("document.all.vkey-s") : document.getElementById('vkey-s');
		loObj.innerHTML = "S";
		loObj = ie4? eval("document.all.vkey-d") : document.getElementById('vkey-d');
		loObj.innerHTML = "D";
		loObj = ie4? eval("document.all.vkey-f") : document.getElementById('vkey-f');
		loObj.innerHTML = "F";
		loObj = ie4? eval("document.all.vkey-g") : document.getElementById('vkey-g');
		loObj.innerHTML = "G";
		loObj = ie4? eval("document.all.vkey-h") : document.getElementById('vkey-h');
		loObj.innerHTML = "H";
		loObj = ie4? eval("document.all.vkey-j") : document.getElementById('vkey-j');
		loObj.innerHTML = "J";
		loObj = ie4? eval("document.all.vkey-k") : document.getElementById('vkey-k');
		loObj.innerHTML = "K";
		loObj = ie4? eval("document.all.vkey-l") : document.getElementById('vkey-l');
		loObj.innerHTML = "L";
		loObj = ie4? eval("document.all.vkey-z") : document.getElementById('vkey-z');
		loObj.innerHTML = "Z";
		loObj = ie4? eval("document.all.vkey-x") : document.getElementById('vkey-x');
		loObj.innerHTML = "X";
		loObj = ie4? eval("document.all.vkey-c") : document.getElementById('vkey-c');
		loObj.innerHTML = "C";
		loObj = ie4? eval("document.all.vkey-v") : document.getElementById('vkey-v');
		loObj.innerHTML = "V";
		loObj = ie4? eval("document.all.vkey-b") : document.getElementById('vkey-b');
		loObj.innerHTML = "B";
		loObj = ie4? eval("document.all.vkey-n") : document.getElementById('vkey-n');
		loObj.innerHTML = "N";
		loObj = ie4? eval("document.all.vkey-m") : document.getElementById('vkey-m');
		loObj.innerHTML = "M";
	}
	else {
		loObj = ie4? eval("document.all.vkey-q") : document.getElementById('vkey-q');
		loObj.innerHTML = "q";
		loObj = ie4? eval("document.all.vkey-w") : document.getElementById('vkey-w');
		loObj.innerHTML = "w";
		loObj = ie4? eval("document.all.vkey-e") : document.getElementById('vkey-e');
		loObj.innerHTML = "e";
		loObj = ie4? eval("document.all.vkey-r") : document.getElementById('vkey-r');
		loObj.innerHTML = "r";
		loObj = ie4? eval("document.all.vkey-t") : document.getElementById('vkey-t');
		loObj.innerHTML = "t";
		loObj = ie4? eval("document.all.vkey-y") : document.getElementById('vkey-y');
		loObj.innerHTML = "y";
		loObj = ie4? eval("document.all.vkey-u") : document.getElementById('vkey-u');
		loObj.innerHTML = "u";
		loObj = ie4? eval("document.all.vkey-i") : document.getElementById('vkey-i');
		loObj.innerHTML = "i";
		loObj = ie4? eval("document.all.vkey-o") : document.getElementById('vkey-o');
		loObj.innerHTML = "o";
		loObj = ie4? eval("document.all.vkey-p") : document.getElementById('vkey-p');
		loObj.innerHTML = "p";
		loObj = ie4? eval("document.all.vkey-a") : document.getElementById('vkey-a');
		loObj.innerHTML = "a";
		loObj = ie4? eval("document.all.vkey-s") : document.getElementById('vkey-s');
		loObj.innerHTML = "s";
		loObj = ie4? eval("document.all.vkey-d") : document.getElementById('vkey-d');
		loObj.innerHTML = "d";
		loObj = ie4? eval("document.all.vkey-f") : document.getElementById('vkey-f');
		loObj.innerHTML = "f";
		loObj = ie4? eval("document.all.vkey-g") : document.getElementById('vkey-g');
		loObj.innerHTML = "g";
		loObj = ie4? eval("document.all.vkey-h") : document.getElementById('vkey-h');
		loObj.innerHTML = "h";
		loObj = ie4? eval("document.all.vkey-j") : document.getElementById('vkey-j');
		loObj.innerHTML = "j";
		loObj = ie4? eval("document.all.vkey-k") : document.getElementById('vkey-k');
		loObj.innerHTML = "k";
		loObj = ie4? eval("document.all.vkey-l") : document.getElementById('vkey-l');
		loObj.innerHTML = "l";
		loObj = ie4? eval("document.all.vkey-z") : document.getElementById('vkey-z');
		loObj.innerHTML = "z";
		loObj = ie4? eval("document.all.vkey-x") : document.getElementById('vkey-x');
		loObj.innerHTML = "x";
		loObj = ie4? eval("document.all.vkey-c") : document.getElementById('vkey-c');
		loObj.innerHTML = "c";
		loObj = ie4? eval("document.all.vkey-v") : document.getElementById('vkey-v');
		loObj.innerHTML = "v";
		loObj = ie4? eval("document.all.vkey-b") : document.getElementById('vkey-b');
		loObj.innerHTML = "b";
		loObj = ie4? eval("document.all.vkey-n") : document.getElementById('vkey-n');
		loObj.innerHTML = "n";
		loObj = ie4? eval("document.all.vkey-m") : document.getElementById('vkey-m');
		loObj.innerHTML = "m";
	}
	
	gbVoidReasonInLowerCase = !gbVoidReasonInLowerCase;
	
	resetRedirect();
}

function saveVoidReason() {
	var loNotes;
	
	loNotes = ie4? eval("document.all.voidreason") : document.getElementById('voidreason');
	gsOrderNotes = loNotes.value;
	if (gsOrderNotes.length > 0) {
		window.location = "neworder.asp?cancel=yes&VoidReason=" + encodeURIComponent(gsOrderNotes);
	}
}

function gotoEditReason(pbGotoPayment) {
	var loNotes, loDiv;
	
	gbEditGotoPayment = pbGotoPayment;
	
	loNotes = ie4? eval("document.all.editreason") : document.getElementById('editreason');
	loNotes.value = "";
	
	loDiv = ie4? eval("document.all.unitselector") : document.getElementById("unitselector");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.ordernotes") : document.getElementById("ordernotes");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmholdorder") : document.getElementById("confirmholdorder");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmcancel") : document.getElementById("confirmcancel");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.holdorder") : document.getElementById("holdorder");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.calendar") : document.getElementById("calendar");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmdelivery") : document.getElementById("confirmdelivery");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.confirmreprint") : document.getElementById("confirmreprint");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.voidreasondiv") : document.getElementById("voidreasondiv");
	loDiv.style.visibility = "hidden";
	
	loDiv = ie4? eval("document.all.editreasondiv") : document.getElementById("editreasondiv");
	loDiv.style.visibility = "visible";
	
	resetRedirect();
}

function addToEditReason(psDigit) {
	var loNotes, lsNotes;
	
	loNotes = ie4? eval("document.all.editreason") : document.getElementById('editreason');
	lsNotes = loNotes.value;
	
	if (psDigit.length > 1) {
		if (lsNotes.length > 0) {
			lsNotes = lsNotes + " ";
		}
	}
	
	if (gbEditReasonInLowerCase) {
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

function backspaceEditReason() {
	var loNotes, lsNotes;
	
	loNotes = ie4? eval("document.all.editreason") : document.getElementById('editreason');
	lsNotes = loNotes.value;
	if (lsNotes.length > 0) {
		lsNotes = lsNotes.substr(0, (lsNotes.length - 1));
		loNotes.value = lsNotes;
	}
	
	resetRedirect();
}

function clearEditReason() {
	var loNotes;
	
	loNotes = ie4? eval("document.all.editreason") : document.getElementById('editreason');
	loNotes.value = "";
	
	resetRedirect();
}

function properEditReason() {
	var loNotes, i, lsText, lbDoUpper;
	
	loNotes = ie4? eval("document.all.editreason") : document.getElementById('editreason');
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

function shiftEditReason() {
	var loObj;
	
	if (gbEditReasonInLowerCase) {
		loObj = ie4? eval("document.all.ekey-q") : document.getElementById('ekey-q');
		loObj.innerHTML = "Q";
		loObj = ie4? eval("document.all.ekey-w") : document.getElementById('ekey-w');
		loObj.innerHTML = "W";
		loObj = ie4? eval("document.all.ekey-e") : document.getElementById('ekey-e');
		loObj.innerHTML = "E";
		loObj = ie4? eval("document.all.ekey-r") : document.getElementById('ekey-r');
		loObj.innerHTML = "R";
		loObj = ie4? eval("document.all.ekey-t") : document.getElementById('ekey-t');
		loObj.innerHTML = "T";
		loObj = ie4? eval("document.all.ekey-y") : document.getElementById('ekey-y');
		loObj.innerHTML = "Y";
		loObj = ie4? eval("document.all.ekey-u") : document.getElementById('ekey-u');
		loObj.innerHTML = "U";
		loObj = ie4? eval("document.all.ekey-i") : document.getElementById('ekey-i');
		loObj.innerHTML = "I";
		loObj = ie4? eval("document.all.ekey-o") : document.getElementById('ekey-o');
		loObj.innerHTML = "O";
		loObj = ie4? eval("document.all.ekey-p") : document.getElementById('ekey-p');
		loObj.innerHTML = "P";
		loObj = ie4? eval("document.all.ekey-a") : document.getElementById('ekey-a');
		loObj.innerHTML = "A";
		loObj = ie4? eval("document.all.ekey-s") : document.getElementById('ekey-s');
		loObj.innerHTML = "S";
		loObj = ie4? eval("document.all.ekey-d") : document.getElementById('ekey-d');
		loObj.innerHTML = "D";
		loObj = ie4? eval("document.all.ekey-f") : document.getElementById('ekey-f');
		loObj.innerHTML = "F";
		loObj = ie4? eval("document.all.ekey-g") : document.getElementById('ekey-g');
		loObj.innerHTML = "G";
		loObj = ie4? eval("document.all.ekey-h") : document.getElementById('ekey-h');
		loObj.innerHTML = "H";
		loObj = ie4? eval("document.all.ekey-j") : document.getElementById('ekey-j');
		loObj.innerHTML = "J";
		loObj = ie4? eval("document.all.ekey-k") : document.getElementById('ekey-k');
		loObj.innerHTML = "K";
		loObj = ie4? eval("document.all.ekey-l") : document.getElementById('ekey-l');
		loObj.innerHTML = "L";
		loObj = ie4? eval("document.all.ekey-z") : document.getElementById('ekey-z');
		loObj.innerHTML = "Z";
		loObj = ie4? eval("document.all.ekey-x") : document.getElementById('ekey-x');
		loObj.innerHTML = "X";
		loObj = ie4? eval("document.all.ekey-c") : document.getElementById('ekey-c');
		loObj.innerHTML = "C";
		loObj = ie4? eval("document.all.ekey-v") : document.getElementById('ekey-v');
		loObj.innerHTML = "V";
		loObj = ie4? eval("document.all.ekey-b") : document.getElementById('ekey-b');
		loObj.innerHTML = "B";
		loObj = ie4? eval("document.all.ekey-n") : document.getElementById('ekey-n');
		loObj.innerHTML = "N";
		loObj = ie4? eval("document.all.ekey-m") : document.getElementById('ekey-m');
		loObj.innerHTML = "M";
	}
	else {
		loObj = ie4? eval("document.all.ekey-q") : document.getElementById('ekey-q');
		loObj.innerHTML = "q";
		loObj = ie4? eval("document.all.ekey-w") : document.getElementById('ekey-w');
		loObj.innerHTML = "w";
		loObj = ie4? eval("document.all.ekey-e") : document.getElementById('ekey-e');
		loObj.innerHTML = "e";
		loObj = ie4? eval("document.all.ekey-r") : document.getElementById('ekey-r');
		loObj.innerHTML = "r";
		loObj = ie4? eval("document.all.ekey-t") : document.getElementById('ekey-t');
		loObj.innerHTML = "t";
		loObj = ie4? eval("document.all.ekey-y") : document.getElementById('ekey-y');
		loObj.innerHTML = "y";
		loObj = ie4? eval("document.all.ekey-u") : document.getElementById('ekey-u');
		loObj.innerHTML = "u";
		loObj = ie4? eval("document.all.ekey-i") : document.getElementById('ekey-i');
		loObj.innerHTML = "i";
		loObj = ie4? eval("document.all.ekey-o") : document.getElementById('ekey-o');
		loObj.innerHTML = "o";
		loObj = ie4? eval("document.all.ekey-p") : document.getElementById('ekey-p');
		loObj.innerHTML = "p";
		loObj = ie4? eval("document.all.ekey-a") : document.getElementById('ekey-a');
		loObj.innerHTML = "a";
		loObj = ie4? eval("document.all.ekey-s") : document.getElementById('ekey-s');
		loObj.innerHTML = "s";
		loObj = ie4? eval("document.all.ekey-d") : document.getElementById('ekey-d');
		loObj.innerHTML = "d";
		loObj = ie4? eval("document.all.ekey-f") : document.getElementById('ekey-f');
		loObj.innerHTML = "f";
		loObj = ie4? eval("document.all.ekey-g") : document.getElementById('ekey-g');
		loObj.innerHTML = "g";
		loObj = ie4? eval("document.all.ekey-h") : document.getElementById('ekey-h');
		loObj.innerHTML = "h";
		loObj = ie4? eval("document.all.ekey-j") : document.getElementById('ekey-j');
		loObj.innerHTML = "j";
		loObj = ie4? eval("document.all.ekey-k") : document.getElementById('ekey-k');
		loObj.innerHTML = "k";
		loObj = ie4? eval("document.all.ekey-l") : document.getElementById('ekey-l');
		loObj.innerHTML = "l";
		loObj = ie4? eval("document.all.ekey-z") : document.getElementById('ekey-z');
		loObj.innerHTML = "z";
		loObj = ie4? eval("document.all.ekey-x") : document.getElementById('ekey-x');
		loObj.innerHTML = "x";
		loObj = ie4? eval("document.all.ekey-c") : document.getElementById('ekey-c');
		loObj.innerHTML = "c";
		loObj = ie4? eval("document.all.ekey-v") : document.getElementById('ekey-v');
		loObj.innerHTML = "v";
		loObj = ie4? eval("document.all.ekey-b") : document.getElementById('ekey-b');
		loObj.innerHTML = "b";
		loObj = ie4? eval("document.all.ekey-n") : document.getElementById('ekey-n');
		loObj.innerHTML = "n";
		loObj = ie4? eval("document.all.ekey-m") : document.getElementById('ekey-m');
		loObj.innerHTML = "m";
	}
	
	gbEditReasonInLowerCase = !gbEditReasonInLowerCase;
	
	resetRedirect();
}

function saveEditReason() {
	var loNotes;
	
	loNotes = ie4? eval("document.all.editreason") : document.getElementById('editreason');
	gsEditReason = loNotes.value;
	if (gsEditReason.length > 0) {
		if (gbEditGotoPayment) {
			window.location = "payment.asp?o=<%=gnOrderID%>&EditReason=" + encodeURIComponent(gsEditReason);
		}
		else {
			gotoConfirmReprint();
		}
	}
}

function verifyClick() {
    alert("Hello this is an Alert");
}

function back2Delivery() {
    var lsLocation = "neworder.asp";
    alert("Back 2 Delivery");
    window.location = lsLocation;
}

function back2Phone() {
    var lsLocation = "neworder.asp";
    alert("Back 2 Phone");
    window.location = lsLocation;
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

			<!-- #Include Virtual="ordering/top-header.asp" -->
			<tr height="733">
				<td valign="top" width="1010">
					<div id="content" style="position: relative; width: 1010px; height: 723px; overflow: auto;">
						<div id="unitselector" style="position: absolute; left: 0px; top: 0px; width: 1010px; height: 723px; visibility: <%If gbConfirmDelivery Then Response.Write("hidden") Else Response.Write("visible")%>;">
							<table cellpadding="0" cellspacing="0" width="1010" height="723">
								<tr>
									<td valign="top" width="340">
<%
' 2013-08-20 TAM: Allow edit order even if complete and paid
'If gnOrderStatusID < 7 Then
If gnOrderStatusID <= 10 Then
	If ganUnitIDs(0) <> 0 Then
		For i = 0 To UBound(ganUnitIDs)
			If i Mod 7 = 0 Then
				If i > 0 And UBound(ganUnitIDs) > 7 Then
%>
											<button style="width: 340px" onclick="toggleDivs('unitdiv<%=Int(i/7)-1%>', 'unitdiv<%=Int(i/7)%>')">(Next)</button><br/>
<%
					If Session("NewOrder") Or gnOrderID = 0 Then
%>
											<button style="width: 167px" onclick="cancelOrder()">Void Order</button>
<%
					Else
%>
											<button style="width: 167px" onclick="cancelOrder()">Void Order</button>
<%
					End If
					
					If gnOrderID <> 0 Then
%>
											<button style="width: 167px" onclick="gotoOrderNotes();">Order Notes</button>
<%
					Else
%>
											<button style="width: 167px; background-color: #C0C0C0;">&nbsp;</button>
<%
					End If
%>
										</div>
										<div id="unitdiv<%=Int(i/7)%>" style="position: absolute; top: 0px; left: 0px; width: 340px; visibility: hidden;">
<%
				Else
					If i = 0 Then
%>
										<div id="unitdiv<%=Int(i/7)%>" style="position: absolute; top: 0px; left: 0px; width: 340px;">
<%
					End If
				End If
			End If
%>
											<button style="width: 340px" onclick="window.location = 'unitedit.asp?u=<%=ganUnitIDs(i)%>'"><%=gasShortDescriptions(i)%></button><br/>
<%
		Next
		
		' Add hidden buttons here
		If ((UBound(ganUnitIDs) + 1) Mod 7) > 0 And UBound(ganUnitIDs) <> 7 Then
			For i = ((UBound(ganUnitIDs) + 1) Mod 7) To 6
%>
											<button style="width: 340px; background-color: #C0C0C0;">&nbsp;</button>
<%
			Next
		End If
	
		If UBound(ganUnitIDs) > 7 Then
			If UBound(ganUnitIDs) <> 7 Then
%>
											<button style="width: 340px" onclick="toggleDivs('unitdiv<%=Int(UBound(ganUnitIDs)/7)%>', 'unitdiv0')">(Next)</button><br/>
<%
			End If
		End If
	End If
Else
%>
										<div id="unitdiv0" style="position: absolute; top: 0px; left: 0px; width: 340px;">
											<button style="width: 340px; background-color: #C0C0C0">&nbsp;</button>
											<button style="width: 340px; background-color: #C0C0C0">&nbsp;</button>
											<button style="width: 340px; background-color: #C0C0C0">&nbsp;</button>
											<button style="width: 340px; background-color: #C0C0C0">&nbsp;</button>
											<button style="width: 340px; background-color: #C0C0C0">&nbsp;</button>
											<button style="width: 340px; background-color: #C0C0C0">&nbsp;</button>
											<button style="width: 340px; background-color: #C0C0C0">&nbsp;</button>
											<button style="width: 340px; background-color: #C0C0C0">&nbsp;</button>
<%
End If

If Session("NewOrder") Or gnOrderID = 0 Then
%>
											<button style="width: 167px" onclick="cancelOrder()">Void Order</button>
<%
Else
%>
											<button style="width: 167px" onclick="cancelOrder()">Void Order</button>
<%
End If
					
' 2013-08-20 TAM: Allow edit order even if complete and paid
'If gnOrderID <> 0 And gnOrderStatusID < 7 Then
If gnOrderID <> 0 And gnOrderStatusID <= 10 Then
%>
											<button style="width: 167px" onclick="gotoOrderNotes();">Order Notes</button>
<%
Else
%>
											<button style="width: 167px; background-color: #C0C0C0">&nbsp;</button>
<%
End If
%>
										</div>
									</td>
									<td valign="top" width="340">
<%
' 2013-08-20 TAM: Allow edit order even if complete and paid
'If gnOrderStatusID < 7 Then
If gnOrderStatusID <= 10 Then
	If gnOrderTypeID = 1 Then
%>
										<button style="width: 340px;" onclick="window.location = 'neworder.asp?o=<%=gnOrderID%>'"><%=gsOrderTypeDescription & " " & FormatCurrency(gdDeliveryCharge)%></button>
<%
	Else
%>
										<button style="width: 340px;" onclick="window.location = 'neworder.asp?o=<%=gnOrderID%>'"><%=gsOrderTypeDescription%></button>
<%
	End If
Else
%>
										<button style="width: 340px; background-color: #C0C0C0">&nbsp;</button>
<%
End If
%>
										<div style="height: 375px; padding: 10px;">
										<%=gsCustomerName%><br/>
<%
If gnAddressID = 1 Then
%>
										&nbsp;<br/>
<%
	If Len(gsCustomerPhone) > 0 Then
%>
										Phone: <%="(" & Left(gsCustomerPhone, 3) & ") " & Mid(gsCustomerPhone, 4, 3) & "-" & Mid(gsCustomerPhone, 7)%><br/>
										&nbsp;<br/>
<%
	End If
Else
	If Len(gsAddress2) = 0 Then
		If Len(gsAddress1) > 0 Then
%>
										<%=gsAddress1%><br/>
<%
		End If
	Else
%>
										<%=gsAddress1%> #<%=gsAddress2%><br/>
<%
	End If
	
	If Len(gsCity) > 0 Or Len(gsState) > 0 Or Len(gsPostalCode) > 0 Then
%>
										<%=gsCity%>, <%=gsState%>&nbsp;<%=gsPostalCode%><br/>
<%
	End If
%>
										&nbsp;<br/>
<%
	If Len(gsAddressNotes) > 0 Then
%>
										Address Notes: <%=gsAddressNotes%><br/>
										&nbsp;<br/>
<%
	End If
	
	If Len(gsCustomerPhone) > 0 Then
%>
										Phone: <%="(" & Left(gsCustomerPhone, 3) & ") " & Mid(gsCustomerPhone, 4, 3) & "-" & Mid(gsCustomerPhone, 7)%><br/>
										&nbsp;<br/>
<%
	End If
	
	If Len(gsCustomerNotes) > 0 Then
%>
										Customer Notes: <%=gsCustomerNotes%><br/>
										&nbsp;<br/>
<%
	End If
End If

If Len(gsOrderNotes) > 0 Then
%>
										Order Notes: <%=gsOrderNotes%><br/>
										&nbsp;<br/>
<%
End If
%>
										</div>
<%
If gnOrderStatusID < 7 And (gnCustomerID > 1 Or Len(gsCustomerName) > 0) And ganOrderLineIDs(0) > 0 Then
%>
										<button style="width: 340px;" onclick="gotoHoldOrder();">Hold Order</button>
<%
Else
%>
										<button style="width: 340px; background-color: #C0C0C0;">&nbsp;</button>
<%
End If

If ganOrderLineIDs(0) = 0 Then
%>
										<button style="width: 167px; background-color: #C0C0C0;">&nbsp;</button>
<%
Else
%>
										<button style="width: 167px;" onclick="window.location = 'coupons.asp?o=<%=gnOrderID%>'">Coupons</button>
<%
End If

If gnOrderID <> 0 And ((gnOrderTypeID = 1 And gdOrderTotal <> 0) Or (gnOrderTypeID <> 1 And (gdOrderTotal - gdOrderDiscountTotal) <> 0)) Then
	If Session("NewOrder") Then
%>
										<button style="width: 167px;" onclick="window.location = 'payment.asp?o=<%=gnOrderID%>'">Payment</button>
<%
	Else
%>
										<button style="width: 167px;" onclick="gotoEditReason(true);">Payment</button>
<%
	End If
Else
%>
										<button style="width: 167px; background-color: #C0C0C0;">&nbsp;</button>
<%
End If

' 2013-08-20 TAM: Allow edit order even if complete and paid
'If gnOrderStatusID < 7 Then
If gnOrderStatusID <= 10 Then
	If gnCustomerID > 1 And gnAddressID > 1 Then
		If gnOrderID > 0 Then
			gsReturnURL = Server.URLEncode("/ordering/unitselect.asp?o=" & gnOrderID)
		Else
			gsReturnURL = Server.URLEncode("/ordering/unitselect.asp?" & Request.QueryString)
		End If
%>
										<button style="width: 167px;" onclick="window.location = '/custmaint/addressnotes.asp?CustomerID=<%=gnCustomerID%>&AddressID=<%=gnAddressID%>&ReturnURL=<%=gsReturnURL%>'">Edit Address Notes</button>
<%
	Else
%>
										<button style="width: 167px; background-color: #C0C0C0;">&nbsp;</button>
<%
	End If
Else
%>
										<button style="width: 167px; background-color: #C0C0C0">&nbsp;</button>
<%
End If

If Session("OrderEdited") Then
	If Session("NewOrder") Then
		If gnOrderID <> 0 And gdOrderTotal <> 0 And gnOrderTypeID <> 1 Then
			If (gdOrderTotal - gdOrderDiscountTotal) <> 0 Then
%>
										<button style="width: 167px;" onclick="window.location = 'neworder.asp?o=<%=gnOrderID%>&v=<%=gnPaymentTypeID%>'">Send To Kitchen</button>
<%
			Else
%>
										<button style="width: 167px;" onclick="window.location = 'neworder.asp?o=<%=gnOrderID%>&v=<%=gnPaymentTypeID%>&s=0'">Send To Kitchen</button>
<%
			End If
		Else
			If gnOrderID <> 0 And gdOrderTotal <> 0 And gnOrderTypeID = 1 And gdOrderTotal = gdOrderDiscountTotal Then
%>
										<button style="width: 167px;" onclick="window.location = 'neworder.asp?o=<%=gnOrderID%>&v=<%=gnPaymentTypeID%>&s=0'">Send To Kitchen</button>
<%
			Else
%>
										<button style="width: 167px; background-color: #C0C0C0;">&nbsp;</button>
<%
			End If
		End If
	Else
		If gnOrderID <> 0 And gdOrderTotal <> 0 And gnOrderTypeID <> 1 Then
%>
										<button style="width: 167px;" onclick="gotoEditReason(false)">Send To Kitchen</button>
<%
		Else
			If gnOrderID <> 0 And gdOrderTotal <> 0 And gnOrderTypeID = 1 And ((gdOrderTotal = gdOrderDiscountTotal) Or (gnPaymentTypeID > 2)) Then
%>
										<button style="width: 167px;" onclick="gotoEditReason(false)">Send To Kitchen</button>
<%
			Else
%>
										<button style="width: 167px; background-color: #C0C0C0;">&nbsp;</button>
<%
			End If
		End If
	End If
Else
%>
										<button style="width: 167px;" onclick="window.location = 'neworder.asp'">Cancel</button>
<%
End If
%>
									</td>
									<td align="right" valign="top" width="330">
										<div style="position: relative; width: 320px; height: 691px; text-align: left; background-color: #FFFFFF;">
<%
If ganOrderLineIDs(0) <> 0 Then
	For i = 0 To UBound(ganOrderLineIDs)
		If i Mod 4 = 0 Then
			If i > 0 Then
%>
												<button style="width: 320px; color: #FFFFFF; background-color: #FF0000;" onclick="toggleDivs('itemdiv<%=Int(i/4)-1%>', 'itemdiv<%=Int(i/4)%>')">Page <%=Int(i/4)%> of <%=Int(UBound(ganOrderLineIDs)/4)+1%><br/>(Next)</button>
											</div>
											<div id="itemdiv<%=Int(i/4)%>" style="position: absolute; top: 0px; visibility: hidden;">
<%
			Else
%>
											<div id="itemdiv<%=Int(i/4)%>" style="position: absolute; top: 0px;">
<%
			End If
		End If
%>
												<div style="height: 142px; padding: 5px; overflow: auto;">
<%
' 2013-08-20 TAM: Allow edit order even if complete and paid
'		If gnOrderStatusID < 7 Then
		If gnOrderStatusID <= 10 Then
%>
													<div onclick="window.location = 'unitedit.asp?l=<%=ganOrderLineIDs(i)%>'"><%=gasOrderLineDescriptions(i)%></div>
<%
		Else
%>
													<div><%=gasOrderLineDescriptions(i)%></div>
<%
		End If
%>
												</div>
<%
	Next
	
	' Add hidden divs here
	If ((UBound(ganOrderLineIDs) + 1) Mod 4) > 0 Then
		For i = ((UBound(ganOrderLineIDs) + 1) Mod 4) To 3
%>
												<div style="height: 152px;">&nbsp;</div>
<%
		Next
	End If
	
	If UBound(ganOrderLineIDs) > 3 Then
%>
												<button style="width: 320px; color: #FFFFFF; background-color: #FF0000;" onclick="toggleDivs('itemdiv<%=Int(UBound(ganOrderLineIDs)/4)%>', 'itemdiv0')">Page <%=Int(UBound(ganOrderLineIDs)/4)+1%> of <%=Int(UBound(ganOrderLineIDs)/4)+1%><br/>(Next)</button>
<%
	End If
End If
%>
											</div>
										</div>
										<div style="width: 320px; text-align: center; background-color: #FFFFFF;">Tax: <%=FormatCurrency(gdTax + gdTax2)%>&nbsp; Delivery: <%=FormatCurrency(gdDeliveryCharge)%>&nbsp; Total: <%=FormatCurrency(gdOrderTotal - gdOrderDiscountTotal)%></div>
									</td>
								</tr>
							</table>
						</div>
						<div id="ordernotes" style="position: absolute; left: 0px; top: 0px; width: 1010px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
							<table align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="9"><div align="center">
										<strong>ENTER ADDITIONAL NOTES FOR THIS ORDER</strong></div></td>
								</tr>
								<tr>
									<td colspan="9"><div align="center">
										<textarea id="notes" style="width: 930px; height: 60px;"></textarea></div></td>
								</tr>
								<tr>
									<td width="11%"><button style="width: 100px;" onclick="addToNotes('Bring Forks')">
									Bring Forks</button></td>
									<td width="11%"><button style="width: 100px;" onclick="addToNotes('Bring Napkins')">
									Bring Napkins</button></td>
									<td width="11%"><button style="width: 100px;" onclick="addToNotes('Bring Plates')">
									Bring Plates</button></td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%"><button style="width: 100px;" onclick="gotoUnitSelector()">Cancel</button></td>
									<td width="11%"><button style="width: 100px;" onclick="clearNotes()">Clear</button></td>
									<td width="11%"><button style="width: 100px;" onclick="saveNotes()">Done</button></td>
								</tr>
								<tr>
									<td width="11%"><button style="width: 100px;" onclick="addToNotes('School Fundraiser')">
									School<br/>Fundraiser</button></td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
									<td width="11%">&nbsp;</td>
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
						<div id="holdorder" style="position: absolute; top: 0px; left: 0px; width: 1010px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
							<table align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="11"><div align="center">
										<strong>ENTER DATE AND TIME CUSTOMER EXPECTS ORDER</strong></div></td>
								</tr>
								<tr>
									<td colspan="11">&nbsp;</td>
								</tr>
								<tr>
									<td colspan="3"><div align="center">
										<strong>ENTER DATE AS MMDDYY<br/>(Ex. 081511)</strong></div></td>
									<td width="25">&nbsp;</td>
									<td colspan="3"><div align="center">
										<strong>ENTER TIME AS HHMM<br/>(Ex. 0453)</strong></div></td>
									<td width="25">&nbsp;</td>
									<td colspan="3"><div align="center">
										<strong>ENTER NUMBER OF MINUTES<br/>AHEAD TO PRINT</strong></div></td>
								</tr>
								<tr>
									<td colspan="3"><div align="center">
										<input type="text" id="date" style="width: 200px" autocomplete="off" value="<%=gsHoldDate%>" onkeydown="disableEnterKey();" /></div></td>
									<td>&nbsp;</td>
									<td colspan="3"><div align="center">
										<input type="text" id="time" style="width: 200px" autocomplete="off" value="<%=gsHoldTime%>" onkeydown="disableEnterKey();" /></div></td>
									<td>&nbsp;</td>
									<td colspan="3"><div align="center">
										<input type="text" id="minutes" style="width: 200px" autocomplete="off" value="<%=gsPrintTime%>" onkeydown="disableEnterKey();" /></div></td>
								</tr>
								<tr>
									<td><button onclick="addToDate('1')">1</button></td>
									<td><button onclick="addToDate('2')">2</button></td>
									<td><button onclick="addToDate('3')">3</button></td>
									<td width="25">&nbsp;</td>
									<td><button onclick="addToTime('1')">1</button></td>
									<td><button onclick="addToTime('2')">2</button></td>
									<td><button onclick="addToTime('3')">3</button></td>
									<td width="25">&nbsp;</td>
									<td><button onclick="addToMinutes('1')">1</button></td>
									<td><button onclick="addToMinutes('2')">2</button></td>
									<td><button onclick="addToMinutes('3')">3</button></td>
								</tr>
								<tr>
									<td><button onclick="addToDate('4')">4</button></td>
									<td><button onclick="addToDate('5')">5</button></td>
									<td><button onclick="addToDate('6')">6</button></td>
									<td width="25">&nbsp;</td>
									<td><button onclick="addToTime('4')">4</button></td>
									<td><button onclick="addToTime('5')">5</button></td>
									<td><button onclick="addToTime('6')">6</button></td>
									<td width="25">&nbsp;</td>
									<td><button onclick="addToMinutes('4')">4</button></td>
									<td><button onclick="addToMinutes('5')">5</button></td>
									<td><button onclick="addToMinutes('6')">6</button></td>
								</tr>
								<tr>
									<td><button onclick="addToDate('7')">7</button></td>
									<td><button onclick="addToDate('8')">8</button></td>
									<td><button onclick="addToDate('9')">9</button></td>
									<td width="25">&nbsp;</td>
									<td><button onclick="addToTime('7')">7</button></td>
									<td><button onclick="addToTime('8')">8</button></td>
									<td><button onclick="addToTime('9')">9</button></td>
									<td width="25">&nbsp;</td>
									<td><button onclick="addToMinutes('7')">7</button></td>
									<td><button onclick="addToMinutes('8')">8</button></td>
									<td><button onclick="addToMinutes('9')">9</button></td>
								</tr>
								<tr>
									<td>&nbsp;</td>
									<td><button onclick="addToDate('0')">0</button></td>
									<td><button onclick="backspaceDate()">Bksp</button></td>
									<td width="25">&nbsp;</td>
									<td><button id="ampm" onclick="toggleAMPM();"><%If gbAM Then %>AM<% Else %>PM<% End If%></button></td>
									<td><button onclick="addToTime('0')">0</button></td>
									<td><button onclick="backspaceTime()">Bksp</button></td>
									<td width="25">&nbsp;</td>
									<td>&nbsp;</td>
									<td><button onclick="addToMinutes('0')">0</button></td>
									<td><button onclick="backspaceMinutes()">Bksp</button></td>
								</tr>
								<tr>
									<td><button onclick="previousDate()">&lt;&lt;</button></td>
									<td><button onclick="return false;" disabled="disabled"></button></td>
									<td><button onclick="nextDate()">&gt;&gt;</button></td>
									<td width="25">&nbsp;</td>
									<td>&nbsp;</td>
									<td>&nbsp;</td>
									<td>&nbsp;</td>
									<td width="25">&nbsp;</td>
									<td>&nbsp;</td>
									<td>&nbsp;</td>
									<td>&nbsp;</td>
								</tr>
								<tr>
									<td colspan="11">&nbsp;</td>
								</tr>
								<tr>
									<td colspan="11" align="center"><button onclick="cancelHoldOrder();">Cancel</button>&nbsp;&nbsp;<button onclick="saveHoldOrder(false);">Done</button></td>
								</tr>
							</table>
						</div>
						<div id="confirmholdorder" style="position: absolute; top: 0px; left: 0px; width: 1010px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
							<p align="center"><strong>The time specified is outside of store hours. Are you sure this is correct?</strong><br/><br/>
							<button onclick="saveHoldOrder(true);">Confirm</button>&nbsp;&nbsp;&nbsp;&nbsp;<button onclick="cancelHoldOrder();">Cancel</button></p>
						</div>
						<div id="confirmcancel" style="position: absolute; top: 0px; left: 0px; width: 1010px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
							<p align="center"><strong>Are you sure you want to cancel this order?</strong><br/><br/>
							<button onclick="gotoVoidReason();">Confirm</button>&nbsp;&nbsp;&nbsp;&nbsp;<button onclick="gotoUnitSelector();">Cancel</button></p>
						</div>
						<div id="confirmreprint" style="position: absolute; top: 0px; left: 0px; width: 1010px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
							<p align="center"><strong>Do you want to reprint this order?</strong><br/><br/>
<%
If gnOrderID <> 0 And gdOrderTotal <> 0 And gnOrderTypeID <> 1 Then
	If (gdOrderTotal - gdOrderDiscountTotal) <> 0 Then
		If gnPaymentTypeID = 3 Then
%>
							<button onclick="window.location = 'neworder.asp?o=<%=gnOrderID%>&v=<%=gnPaymentTypeID%>&r=<%=gsPaymentReference%>&EditReason=' + encodeURIComponent(gsEditReason)">Reprint</button>&nbsp;&nbsp;&nbsp;&nbsp;<button onclick="window.location = 'neworder.asp?o=<%=gnOrderID%>&v=<%=gnPaymentTypeID%>&r=<%=gsPaymentReference%>&q=yes&EditReason=' + encodeURIComponent(gsEditReason)">No Reprint</button>&nbsp;&nbsp;&nbsp;&nbsp;<button onclick="gotoUnitSelector();">Cancel</button>
<%
		Else
			If gnPaymentTypeID = 4 Then
%>
							<button onclick="window.location = 'neworder.asp?o=<%=gnOrderID%>&v=<%=gnPaymentTypeID%>&r=<%=gnAccountID%>&EditReason=' + encodeURIComponent(gsEditReason)">Reprint</button>&nbsp;&nbsp;&nbsp;&nbsp;<button onclick="window.location = 'neworder.asp?o=<%=gnOrderID%>&v=<%=gnPaymentTypeID%>&r=<%=gnAccountID%>&q=yes&EditReason=' + encodeURIComponent(gsEditReason)">No Reprint</button>&nbsp;&nbsp;&nbsp;&nbsp;<button onclick="gotoUnitSelector();">Cancel</button>
<%
			Else
%>
							<button onclick="window.location = 'neworder.asp?o=<%=gnOrderID%>&v=<%=gnPaymentTypeID%>&EditReason=' + encodeURIComponent(gsEditReason)">Reprint</button>&nbsp;&nbsp;&nbsp;&nbsp;<button onclick="window.location = 'neworder.asp?o=<%=gnOrderID%>&v=<%=gnPaymentTypeID%>&q=yes&EditReason=' + encodeURIComponent(gsEditReason)">No Reprint</button>&nbsp;&nbsp;&nbsp;&nbsp;<button onclick="gotoUnitSelector();">Cancel</button>
<%
			End If
		End If
	Else
%>
							<button onclick="window.location = 'neworder.asp?o=<%=gnOrderID%>&v=<%=gnPaymentTypeID%>&s=0&EditReason=' + encodeURIComponent(gsEditReason)">Reprint</button>&nbsp;&nbsp;&nbsp;&nbsp;<button onclick="window.location = 'neworder.asp?o=<%=gnOrderID%>&v=<%=gnPaymentTypeID%>&s=0&q=yes&EditReason=' + encodeURIComponent(gsEditReason)">No Reprint</button>&nbsp;&nbsp;&nbsp;&nbsp;<button onclick="gotoUnitSelector();">Cancel</button>
<%
	End If
Else
	If gnOrderID <> 0 And gdOrderTotal <> 0 And gnOrderTypeID = 1 And ((gdOrderTotal = gdOrderDiscountTotal) Or (gnPaymentTypeID > 2)) Then
%>
							<button onclick="window.location = 'neworder.asp?o=<%=gnOrderID%>&v=<%=gnPaymentTypeID%>&s=0&EditReason=' + encodeURIComponent(gsEditReason)">Reprint</button>&nbsp;&nbsp;&nbsp;&nbsp;<button onclick="window.location = 'neworder.asp?o=<%=gnOrderID%>&v=<%=gnPaymentTypeID%>&s=0&q=yes&EditReason=' + encodeURIComponent(gsEditReason)">No Reprint</button>&nbsp;&nbsp;&nbsp;&nbsp;<button onclick="gotoUnitSelector();">Cancel</button>
<%
	End If
End If
%>
							</p>
						</div>
						<div id="calendar" style="position: absolute; top: 0px; left: 0px; width: 1010px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
						</div>
						<div id="confirmdelivery" style="position: absolute; top: 0px; left: 0px; width: 1010px; height: 723px; visibility: <%If gbConfirmDelivery Then Response.Write("visible") Else Response.Write("hidden")%>; background-color: #fbf3c5;">
							<p align="center"><strong>Tell the customer &quot;The delivery charge to your home is <%=FormatCurrency(gdDeliveryCharge)%> due to the distance to your home and the price of gas&quot;.</strong><br/><br/>
							<button onclick="gotoUnitSelector();">Continue</button></p>
						</div>
						<div id="voidreasondiv" style="position: absolute; left: 0px; top: 0px; width: 1010px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
							<table align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="9"><div align="center">
										<strong>WHY IS THIS ORDER BEING VOIDED</strong></div></td>
								</tr>
								<tr>
									<td colspan="9"><div align="center">
										<textarea id="voidreason" style="width: 930px; height: 60px;"></textarea></div></td>
								</tr>
								<tr>
<%
If Len(gasVoidReasons(0)) > 0 Then
	For i = 0 To UBound(gasVoidReasons)
%>
									<td width="11%"><button style="width: 100px;" onclick="addToVoidReason('<%=gasVoidReasons(i)%>')"><%=gasVoidReasons(i)%></button></td>
<%
		If i = 5 Then
			Exit For
		End If
	Next
	
	If UBound(gasVoidReasons) < 5 Then
		For i = 4 To UBound(gasVoidReasons) Step -1
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
									<td width="11%"><button style="width: 100px;" onclick="gotoUnitSelector();">Cancel</button></td>
									<td width="11%"><button style="width: 100px;" onclick="clearVoidReason();">Clear</button></td>
									<td width="11%"><button style="width: 100px;" onclick="saveVoidReason();">Done</button></td>
								</tr>
								<tr>
<%
If UBound(gasVoidReasons) > 5 Then
	For i = 6 To UBound(gasVoidReasons)
%>
									<td width="11%"><button style="width: 100px;" onclick="addToVoidReason('<%=gasVoidReasons(i)%>')"><%=gasVoidReasons(i)%></button></td>
<%
		If i = 14 Then
			Exit For
		End If
	Next
	
	If UBound(gasVoidReasons) < 14 Then
		For i = 13 To UBound(gasVoidReasons) Step -1
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
										<button onclick="addToVoidReason('+')">+</button><button onclick="addToVoidReason('!')">!</button><button onclick="addToVoidReason('@')">@</button><button onclick="addToVoidReason('#')">#</button><button onclick="addToVoidReason('$')">$</button><button onclick="addToVoidReason('%')">%</button><button onclick="addToVoidReason('^')">^</button><button onclick="addToVoidReason('&')">&amp;</button><button onclick="addToVoidReason('*')">*</button><button onclick="addToVoidReason('(')">(</button><button onclick="addToVoidReason(')')">)</button><button onclick="addToVoidReason(':')">:</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="addToVoidReason('=')">=</button><button onclick="addToVoidReason('1')">1</button><button onclick="addToVoidReason('2')">2</button><button onclick="addToVoidReason('3')">3</button><button onclick="addToVoidReason('4')">4</button><button onclick="addToVoidReason('5')">5</button><button onclick="addToVoidReason('6')">6</button><button onclick="addToVoidReason('7')">7</button><button onclick="addToVoidReason('8')">8</button><button onclick="addToVoidReason('9')">9</button><button onclick="addToVoidReason('0')">0</button><button onclick="addToVoidReason('?')">?</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="addToVoidReason('\'')">'</button><button name="vkey-q" id="vkey-q" onclick="addToVoidReason('Q')">Q</button><button name="vkey-w" id="vkey-w" onclick="addToVoidReason('W')">W</button><button name="vkey-e" id="vkey-e" onclick="addToVoidReason('E')">E</button><button name="vkey-r" id="vkey-r" onclick="addToVoidReason('R')">R</button><button name="vkey-t" id="vkey-t" onclick="addToVoidReason('T')">T</button><button name="vkey-y" id="vkey-y" onclick="addToVoidReason('Y')">Y</button><button name="vkey-u" id="vkey-u" onclick="addToVoidReason('U')">U</button><button name="vkey-i" id="vkey-i" onclick="addToVoidReason('I')">I</button><button name="vkey-o" id="vkey-o" onclick="addToVoidReason('O')">O</button><button name="vkey-p" id="vkey-p" onclick="addToVoidReason('P')">P</button><button onclick="addToVoidReason('/')">/</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="addToVoidReason('.')">.</button><button name="vkey-a" id="vkey-a" onclick="addToVoidReason('A')">A</button><button name="vkey-s" id="vkey-s" onclick="addToVoidReason('S')">S</button><button name="vkey-d" id="vkey-d" onclick="addToVoidReason('D')">D</button><button name="vkey-f" id="vkey-f" onclick="addToVoidReason('F')">F</button><button name="vkey-g" id="vkey-g" onclick="addToVoidReason('G')">G</button><button name="vkey-h" id="vkey-h" onclick="addToVoidReason('H')">H</button><button name="vkey-j" id="vkey-j" onclick="addToVoidReason('J')">J</button><button name="vkey-k" id="vkey-k" onclick="addToVoidReason('K')">K</button><button name="vkey-l" id="vkey-l" onclick="addToVoidReason('L')">L</button><button onclick="addToVoidReason(',')">,</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="shiftVoidReason()">Shift</button><button onclick="addToVoidReason('<')">&lt;</button><button name="vkey-z" id="vkey-z" onclick="addToVoidReason('Z')">Z</button><button name="vkey-x" id="vkey-x" onclick="addToVoidReason('X')">X</button><button name="vkey-c" id="vkey-c" onclick="addToVoidReason('C')">C</button><button name="vkey-v" id="vkey-v" onclick="addToVoidReason('V')">V</button><button name="vkey-b" id="vkey-b" onclick="addToVoidReason('B')">B</button><button name="vkey-n" id="vkey-n" onclick="addToVoidReason('N')">N</button><button name="vkey-m" id="vkey-m" onclick="addToVoidReason('M')">M</button><button onclick="addToVoidReason('>')">&gt;</button><button onclick="properVoidReason()">Proper</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button style="width: 150px;" onclick="addToVoidReason('1/2')">1/2</button>&nbsp;<button style="width: 600px;" onclick="addToVoidReason(' ')">Space</button>&nbsp;<button style="width: 150px;" onclick="backspaceVoidReason()">Bksp</button></div></td>
								</tr>
							</table>
						</div>
						<div id="editreasondiv" style="position: absolute; left: 0px; top: 0px; width: 1010px; height: 723px; visibility: hidden; background-color: #fbf3c5;">
							<table align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="9"><div align="center">
										<strong>WHY IS THIS ORDER BEING EDITED</strong></div></td>
								</tr>
								<tr>
									<td colspan="9"><div align="center">
										<textarea id="editreason" style="width: 930px; height: 60px;"></textarea></div></td>
								</tr>
								<tr>
<%
If Len(gasEditReasons(0)) > 0 Then
	For i = 0 To UBound(gasEditReasons)
%>
									<td width="11%"><button style="width: 100px;" onclick="addToEditReason('<%=gasEditReasons(i)%>')"><%=gasEditReasons(i)%></button></td>
<%
		If i = 5 Then
			Exit For
		End If
	Next
	
	If UBound(gasEditReasons) < 5 Then
		For i = 4 To UBound(gasEditReasons) Step -1
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
									<td width="11%"><button style="width: 100px;" onclick="gotoUnitSelector();">Cancel</button></td>
									<td width="11%"><button style="width: 100px;" onclick="clearEditReason();">Clear</button></td>
									<td width="11%"><button style="width: 100px;" onclick="saveEditReason();">Done</button></td>
								</tr>
								<tr>
<%
If UBound(gasEditReasons) > 5 Then
	For i = 6 To UBound(gasEditReasons)
%>
									<td width="11%"><button style="width: 100px;" onclick="addToEditReason('<%=gasEditReasons(i)%>')"><%=gasEditReasons(i)%></button></td>
<%
		If i = 14 Then
			Exit For
		End If
	Next
	
	If UBound(gasEditReasons) < 14 Then
		For i = 13 To UBound(gasEditReasons) Step -1
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
										<button onclick="addToEditReason('+')">+</button><button onclick="addToEditReason('!')">!</button><button onclick="addToEditReason('@')">@</button><button onclick="addToEditReason('#')">#</button><button onclick="addToEditReason('$')">$</button><button onclick="addToEditReason('%')">%</button><button onclick="addToEditReason('^')">^</button><button onclick="addToEditReason('&')">&amp;</button><button onclick="addToEditReason('*')">*</button><button onclick="addToEditReason('(')">(</button><button onclick="addToEditReason(')')">)</button><button onclick="addToEditReason(':')">:</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="addToEditReason('=')">=</button><button onclick="addToEditReason('1')">1</button><button onclick="addToEditReason('2')">2</button><button onclick="addToEditReason('3')">3</button><button onclick="addToEditReason('4')">4</button><button onclick="addToEditReason('5')">5</button><button onclick="addToEditReason('6')">6</button><button onclick="addToEditReason('7')">7</button><button onclick="addToEditReason('8')">8</button><button onclick="addToEditReason('9')">9</button><button onclick="addToEditReason('0')">0</button><button onclick="addToEditReason('?')">?</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="addToEditReason('\'')">'</button><button name="ekey-q" id="ekey-q" onclick="addToEditReason('Q')">Q</button><button name="ekey-w" id="ekey-w" onclick="addToEditReason('W')">W</button><button name="ekey-e" id="ekey-e" onclick="addToEditReason('E')">E</button><button name="ekey-r" id="ekey-r" onclick="addToEditReason('R')">R</button><button name="ekey-t" id="ekey-t" onclick="addToEditReason('T')">T</button><button name="ekey-y" id="ekey-y" onclick="addToEditReason('Y')">Y</button><button name="ekey-u" id="ekey-u" onclick="addToEditReason('U')">U</button><button name="ekey-i" id="ekey-i" onclick="addToEditReason('I')">I</button><button name="ekey-o" id="ekey-o" onclick="addToEditReason('O')">O</button><button name="ekey-p" id="ekey-p" onclick="addToEditReason('P')">P</button><button onclick="addToEditReason('/')">/</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="addToEditReason('.')">.</button><button name="ekey-a" id="ekey-a" onclick="addToEditReason('A')">A</button><button name="ekey-s" id="ekey-s" onclick="addToEditReason('S')">S</button><button name="ekey-d" id="ekey-d" onclick="addToEditReason('D')">D</button><button name="ekey-f" id="ekey-f" onclick="addToEditReason('F')">F</button><button name="ekey-g" id="ekey-g" onclick="addToEditReason('G')">G</button><button name="ekey-h" id="ekey-h" onclick="addToEditReason('H')">H</button><button name="ekey-j" id="ekey-j" onclick="addToEditReason('J')">J</button><button name="ekey-k" id="ekey-k" onclick="addToEditReason('K')">K</button><button name="ekey-l" id="ekey-l" onclick="addToEditReason('L')">L</button><button onclick="addToEditReason(',')">,</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button onclick="shiftEditReason()">Shift</button><button onclick="addToEditReason('<')">&lt;</button><button name="ekey-z" id="ekey-z" onclick="addToEditReason('Z')">Z</button><button name="ekey-x" id="ekey-x" onclick="addToEditReason('X')">X</button><button name="ekey-c" id="ekey-c" onclick="addToEditReason('C')">C</button><button name="ekey-v" id="ekey-v" onclick="addToEditReason('V')">V</button><button name="ekey-b" id="ekey-b" onclick="addToEditReason('B')">B</button><button name="ekey-n" id="ekey-n" onclick="addToEditReason('N')">N</button><button name="ekey-m" id="ekey-m" onclick="addToEditReason('M')">M</button><button onclick="addToEditReason('>')">&gt;</button><button onclick="properEditReason()">Proper</button></div></td>
								</tr>
								<tr>
									<td><div align="center">
										<button style="width: 150px;" onclick="addToEditReason('1/2')">1/2</button>&nbsp;<button style="width: 600px;" onclick="addToEditReason(' ')">Space</button>&nbsp;<button style="width: 150px;" onclick="backspaceEditReason()">Bksp</button></div></td>
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
