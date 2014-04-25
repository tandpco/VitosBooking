<%
Option Explicit
Response.buffer = TRUE

If Session("SecurityID") = "" Then
	Response.Redirect("/default.asp")
End If

If Request("o").Count = 0 Then
	If Request("PayInAmount").Count = 0 Then
		If Request("Action") = "Apply" Then
			If Request("OpenCount").Count = 0 Then
				Response.Redirect("neworder.asp")
			Else
				If Not IsNumeric(Request("OpenCount")) Then
					Response.Redirect("neworder.asp")
				End If
			End If
		Else
			Response.Redirect("neworder.asp")
		End If
	Else
		If Request("PayInAmount").Count = 0 Then
			Response.Redirect("neworder.asp")
		Else
			If Not IsNumeric(Request("PayInAmount")) And Not IsNumeric(Request("PayInMethod")) And Not IsNumeric(Request("PayInCategory")) Then
				Response.Redirect("neworder.asp")
			End If
		End If
	End If
Else
	If Not IsNumeric(Request("o")) Then
		Response.Redirect("neworder.asp")
	End If
	
	If Request("s").Count = 0 Then
		Response.Redirect("neworder.asp")
	Else
		If Not IsNumeric(Request("s")) Then
			Response.Redirect("neworder.asp")
		End If
	End If
	
	If Request("v").Count = 0 Then
		Response.Redirect("neworder.asp")
	Else
		If Not IsNumeric(Request("v")) Then
			Response.Redirect("neworder.asp")
		End If
	End If
	
	If CLng(Request("v")) = 3 Then
		If Request("xcardnum").Count = 0 Then
			Response.Redirect("/error.asp?o=" & Session("OrderID") & "&err=" & Server.URLEncode("Missing credit card number."))
		Else
			If Len(Request("xcardnum")) < 17 Then
				If Not IsNumeric(Request("xcardnum")) Then
					Response.Redirect("/error.asp?o=" & Session("OrderID") & "&err=" & Server.URLEncode("Credit card number must be numeric."))
				Else
					If Len(Request("xcardnum")) < 15 Then
						Response.Redirect("/error.asp?o=" & Session("OrderID") & "&err=" & Server.URLEncode("Credit card number is too short."))
					End If
				End If
			End If
		End If
		
		If Len(Request("xcardnum")) < 17 Then
			If Request("xexpdate").Count = 0 Then
				Response.Redirect("/error.asp?o=" & Session("OrderID") & "&err=" & Server.URLEncode("Missing expiration date."))
			Else
				If Len(Request("xexpdate")) <> 4 Then
					Response.Redirect("/error.asp?o=" & Session("OrderID") & "&err=" & Server.URLEncode("Expiration date has an invalid length."))
				Else
					If Not IsNumeric(Request("expdate")) Then
						Response.Redirect("/error.asp?o=" & Session("OrderID") & "&err=" & Server.URLEncode("Expiration date must be numeric."))
					End If
				End If
			End If
		End If
		
		If Request("xtip").Count > 0 Then
			If Len(Request("xtip")) > 0 Then
				If Not IsNumeric(Request("xtip")) Then
					Response.Redirect("/error.asp?o=" & Session("OrderID") & "&err=" & Server.URLEncode("Tip amount must be numeric."))
				End If
			End If
		End If
	Else
		If Request("j").Count = 0 Then
			Response.Redirect("neworder.asp")
		Else
			If Not IsNumeric(Request("j")) Then
				Response.Redirect("neworder.asp")
			End If
		End If
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
<!-- #Include Virtual="include2/heartland.asp" -->
<!-- #Include Virtual="include2/mail.asp" -->
<%
Dim gnOrderID, gnSessionID, gsIPAddress, gnEmpID, gsRefID, gdtTransactionDate, gdtSubmitDate, gdtReleaseDate, gdtExpectedDate, gnStoreID, gnCustomerID, gsCustomerName, gsCustomerPhone, gnAddressID, gnOrderTypeID, gbIsPaid, gnPaymentTypeID, gsPaymentReference, gnAccountID, gdDeliveryCharge, gdDriverMoney, gdTax, gdTax2, gdTip, gnOrderStatusID, gsOrderNotes
Dim gbQuickMode
Dim gsOrderTypeDescription
Dim gsEMail, gsFirstName, gsLastName, gdtBirthdate, gnPrimaryAddressID, gsHomePhone, gsCellPhone, gsWorkPhone, gsFAXPhone, gbIsEMailList, gbIsTextList
Dim gsAddress1, gsAddress2, gsCity, gsState, gsPostalCode, gsAddressNotes
Dim gsAddressDescription, gsCustomerNotes
Dim ganUnitIDs(), gasDescriptions(), gasShortDescriptions()
Dim gdOrderTotal
Dim gdTenderTotal
Dim gsPrinterIPAddress, gbIsCashDrawer2
Dim i, gnTmp
Dim ganOrderIDs(), gsLocalErrorMsg
Dim gsCardNumber, gsExpDate
Dim gsPayInOutAccountNumber, gnPayAmount, gnPayInOutCategory, gsPayInOutCheckNumber, gsPayInFrom, gsPayMemo, gnPayInOutMethod
Dim gbNeedPrinterAlert
Dim gbSignaturePrint
Dim gsAccountName, gsPrimaryContactName, gsPrimaryContactEmail, gsSMSEmail, gsMailBody
Dim gsStoreName, gsStoreAddress1, gsStoreAddress2, gsStoreCity, gsStoreState, gsStorePostalCode, gsStorePhone, gsStoreFAX, gsStoreHours

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

If Request("q").Count <> 0 Then
	If Request("q") = "yes" Then
		Session("QuickMode") = TRUE
	End If
End If

If Request("signatureprint").Count <> 0 Then
	If Request("signatureprint") = "yes" Then
		gbSignaturePrint = TRUE
	Else
		gbSignaturePrint = FALSE
	End If
Else
	gbSignaturePrint = TRUE
End If

If Request("o").Count <> 0 Then
	gnOrderID = CLng(Request("o"))
	Session("OrderID") = gnOrderID
	If GetOrderDetails(gnOrderID, gnSessionID, gsIPAddress, gnEmpID, gsRefID, gdtTransactionDate, gdtSubmitDate, gdtReleaseDate, gdtExpectedDate, gnStoreID, gnCustomerID, gsCustomerName, gsCustomerPhone, gnAddressID, gnOrderTypeID, gbIsPaid, gnPaymentTypeID, gsPaymentReference, gnAccountID, gdDeliveryCharge, gdDriverMoney, gdTax, gdTax2, gdTip, gnOrderStatusID, gsOrderNotes) Then
'		If gnStoreID <> Session("StoreID") Then
'			Response.Redirect("otherstore.asp?t=" & gnOrderTypeID & "&s=" & gnStoreID & "&c=" & gnCustomerID & "&a=" & gnAddressID)
'		End If
		
		If DateValue(gdtTransactionDate) <> DateValue(Session("TransactionDate")) Then
			Response.Redirect("/error.asp?err=" & Server.URLEncode("Order is not From Today"))
		End If
		
' 2013-08-20 TAM: Allow edit order even if complete and paid
'		If gnOrderStatusID >= 10 Then
'			Response.Redirect("/error.asp?err=" & Server.URLEncode("Order is Complete"))
'		End If
'		
'		If gbIsPaid Then
'			Response.Redirect("/error.asp?err=" & Server.URLEncode("Order Has Been Paid"))
'		End If
	
	' Don't replace SessionID or IPAddress
	'	Session("SessionID") = gnSessionID
	'	Session("IPAddress") = gsIPAddress
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
		If Len(gsDBErrorMessage) > 0 Then
			Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
		Else
			Response.Redirect("/error.asp?err=" & Server.URLEncode("Invalid Order Specified"))
		End If
	End If
	
	gdOrderTotal = Session("OrderTotal")
	gnPaymentTypeID = CLng(Request("v"))
	gdTenderTotal = CDbl(Request("s"))
	If Request("r").Count > 0 Then
		If gnPaymentTypeID = 4 Then
			gsPaymentReference = ""
			gnAccountID = CLng(Request("r"))
		Else
			gsPaymentReference = Trim(Request("r"))
			gnAccountID = 0
		End If
	Else
		gsPaymentReference = ""
		gnAccountID = 0
	End If
	
	If Len(Session("EditReason")) > 0 Then
		If Not SetOrderEdited(CLng(Request("o")), Session("EmpID"), Session("EditReason")) Then
			Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
		End If
	End If
	
	If gnPaymentTypeID = 3 Then
		gsCardNumber = Request("xcardnum")
		gsExpDate = Request("xexpdate")
		If Request("xtip").Count > 0 Then
			If Len(Request("xtip")) > 0 Then
				gdTip = CDbl(Request("xtip"))
			Else
				gdTip = 0.00
			End If
		Else
			gdTip = 0.00
		End If
		
		If CCAuth(gnStoreID, gsCardNumber, gsExpDate, gsCustomerName, gsAddress1, gsPostalCode, gnOrderID, gnCustomerID, (gdOrderTotal + gdTip), gsPaymentReference) Then
			If gsPaymentReference = "DECLINED" Then
				Response.Redirect("/error.asp?o=" & Session("OrderID") & "&err=" & Server.URLEncode("The credit card has been DECLINED."))
			Else
				If Left(gsPaymentReference, 6) = "ERROR " Then
					Response.Redirect("/error.asp?o=" & Session("OrderID") & "&err=" & Server.URLEncode(gsPaymentReference))
				End If
			End If
		Else
			Response.Redirect("/error.asp?o=" & Session("OrderID") & "&err=" & Server.URLEncode(gsDBErrorMessage))
		End If
	Else
		gdTip = CDbl(Request("j"))
	End If
	
	If Not SetOrderPayment(gnOrderID, gnPaymentTypeID, gsPaymentReference, gnAccountID, gdTip, Session("EmpID")) Then
		Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
	End If
	
	If gnPaymentTypeID = 4 Then
' 2013-10-10 TAM: Debit at store closeout
'		If Not DebitAccountLedger(Session("StoreID"), gdtTransactionDate, gnAccountID, gnOrderID, "Order #" & gnOrderID, (gdOrderTotal + gdTip)) Then
'			Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
'		End If
		If Not IsAccountCollegeDebit(gnAccountID) Then
			If GetAccountContact(gnAccountID, gsAccountName, gsPrimaryContactName, gsPrimaryContactEmail, gsSMSEmail) Then
				If GetStoreDetails(Session("StoreID"), gsStoreName, gsStoreAddress1, gsStoreAddress2, gsStoreCity, gsStoreState, gsStorePostalCode, gsStorePhone, gsStoreFAX, gsStoreHours) Then
					gsMailBody = "On " & Session("TransactionDate") & " order # " & gnOrderID & " for " & FormatCurrency((gdOrderTotal + gdTip)) & " was placed on the " & gsAccountName & " account "
					gsMailBody = gsMailBody & "at Store # " & Session("StoreID") & ", " & gsStoreAddress1 & ". If you have questions about your order call " & Left(gsStorePhone, 3) & "-" & Mid(gsStorePhone, 4, 3) & "-" & Mid(gsStorePhone, 7) & ". "
					gsMailBody = gsMailBody & "If you have billing questions call 866-720-8486."
					
					If Len(gsSMSEmail) > 0 Then
						SendMail gsSMTPFrom, gsSMSEmail, "", "", "", gsMailBody, "", NULL
					End If
					
					If Len(gsPrimaryContactEmail) > 0 Then
						gsMailBody = gsMailBody & CHR(13) & CHR(10) & CHR(13) & CHR(10)
						gsMailBody = gsMailBody & "This is an unattended e-mail address, please do not reply. If you have any questions please contact Vito's Franchising at 866-720-8486." & CHR(13) & CHR(10) & CHR(13) & CHR(10)
						SendMail gsSMTPFrom, gsPrimaryContactEmail, "", "", "Vito's Pizza Order On Account", gsMailBody, "", NULL
					End If
				Else
					Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
				End If
			Else
				Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
			End If
		End If
	End If
	
	If Session("OrderEdited") Then
		If Not Session("QuickMode") Then
			If Not PrintOrder(Session("StoreID"), Session("OrderID"), Session("NewOrder")) Then
				gbNeedPrinterAlert = TRUE
				If Session("NewOrder") Then
					gsLocalErrorMsg = "PRINT FAILURE, SAVING AS A HOLD ORDER!"
				Else
					gsLocalErrorMsg = "PRINT FAILURE, CANNOT PRINT THIS ORDER!"
				End If
			End If
		End If
	End If
	
	If Session("NewOrder") Then
		If gbNeedPrinterAlert Then
			If Not SubmitHoldOrder(Session("OrderID"), DateAdd("n", 15, Now), 15, Session("CustomerID"), Session("AddressID")) Then
				Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
			End If
		Else
			If Not SubmitOrder(Session("OrderID"), Session("CustomerID"), Session("AddressID")) Then
				Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
			End If
		End If
	End If
	
	Session("OrderEdited") = FALSE
	Session("NewOrder") = FALSE
	
	If Not gbNeedPrinterAlert Then
		If gnPaymentTypeID = 1 Or gnPaymentTypeID = 2 Then
			If gnOrderTypeID <> 1 Then
				If Not SetOrderCompleted(gnOrderID) Then
					Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
				End If
			End If
		Else
			If gnOrderTypeID = 1 Then
				Response.Redirect("neworder.asp")
			Else
				If gbSignaturePrint Then
					If Not SetOrderCompleted(gnOrderID) Then
						Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
					End If
					
					If Not PrintSignatureCopies(Session("StoreID"), Session("OrderID"), gsIPAddress) Then
						Response.Redirect("/error.asp?err=" & Server.URLEncode("Could not print signature copies, check the printer."))
					End If
				End If
			End If
		End If
	End If
Else
	gnStoreID = Session("StoreID")
	gsIPAddress = Request.ServerVariables("REMOTE_ADDR")
	gnOrderID = -1
	gsCardNumber = Request("xcardnum")
	gsExpDate = Request("xexpdate")
	gdTip = 0.00
	
	If Request("Action") = "Apply" Then 'Form Submitted
		Dim intItem, intCounter, sqlInsertPayment, Discount, TotalPaying
		
		gsPayInFrom = GetAccountName(Request.Form("AccountID"))
		
		intCounter = Request.Form("OpenCount")
		TotalPaying = 0.00
		For intItem = 1 To intCounter
			If CDbl(Request.Form("Paying-" & intItem)) > 0 Then
				TotalPaying = TotalPaying + CDbl(Request.Form("Paying-" & intItem))
			End If
		Next
		
		gdTenderTotal = TotalPaying
		gdOrderTotal = TotalPaying
		gnPayAmount = TotalPaying
		
		If CCSale(gnStoreID, gsCardNumber, gsExpDate, Request.Form("AccountID"), "", "", 0, 0, TotalPaying, 0, 0, gsPaymentReference) Then
			If gsPaymentReference = "DECLINED" Then
				Response.Redirect("/error.asp?err=" & Server.URLEncode("The credit card has been DECLINED."))
			Else
				If Left(gsPaymentReference, 6) = "ERROR " Then
					Response.Redirect("/error.asp?err=" & Server.URLEncode(gsPaymentReference))
				End If
			End If
		Else
			Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
		End If
		
		For intItem = 1 To intCounter
			If Request.Form("Discount-" & intItem) = "" Then
				Discount = 0
			Else
				Discount = CDbl(Request.Form("Discount-" & intItem))
			End If

			If (CDbl(Request.Form("Paying-" & intItem)) + Discount) >0 Then
				sqlInsertPayment="Insert into tblLedger(AccountID, StoreID, OrderID, LedgerDate, LedgerDescription, Discount, Credit, ReferenceNumber, PaymentTypeID, PaymentAppliedBy) Values("&Request.Form("AccountID")&", " & Session("StoreID") & ", "&Request.Form("OrderID-" & intItem)&", '" & Session("TransactionDate") & "', 'Payment On Order #" & Request.Form("OrderID-" & intItem) & "', "&Discount&", "&Request.Form("Paying-" & intItem)&", '"& gsPaymentReference &"', "&Request.Form("PaymentTypeID")&", "&Session("EmpID")&")"

				If Not DBExecuteSQL(sqlInsertPayment) Then
					Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
				End If
			End If
		Next
	Else
		gsPayInOutAccountNumber = DBCleanLiteral(Request.Form("PayInAccountNumber"))
		gnPayAmount = CDbl(Request.Form("PayInAmount"))
		gnPayInOutCategory = CLng(Request.Form("PayInCategory"))
		gsPayInOutCheckNumber = DBCleanLiteral(Request.Form("PayInCheckNumber"))
		gsPayInFrom = DBCleanLiteral(Request.Form("PayInFrom"))
		gsPayMemo = DBCleanLiteral(Request.Form("PayInMemo"))
		gnPayInOutMethod = CLng(Request.Form("PayInMethod"))
		
		gdTenderTotal = gnPayAmount
		gdOrderTotal = gnPayAmount
		
		If CCSale(gnStoreID, gsCardNumber, gsExpDate, gsPayInFrom, "", "", 0, 0, gnPayAmount, 0, 0, gsPaymentReference) Then
			If gsPaymentReference = "DECLINED" Then
				Response.Redirect("/error.asp?err=" & Server.URLEncode("The credit card has been DECLINED."))
			Else
				If Left(gsPaymentReference, 6) = "ERROR " Then
					Response.Redirect("/error.asp?err=" & Server.URLEncode(gsPaymentReference))
				End If
			End If
		Else
			Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
		End If
		
		If Not CreatePayIn(gnStoreID, Session("EmpID"), gsPayInFrom, gnPayInOutMethod, gnPayAmount, gsPayInOutCheckNumber, gsPayInOutAccountNumber, gsPaymentReference, gnPayInOutCategory, gsPayMemo, Session("TransactionDate")) Then
			Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
		End If
	End If
	
	'----------------------------------------------------- PRINT RECEIPT --------------------------------------------------------------
	Dim msg, bRet, sPrinter, nFontSize, szText, sPrintData, sPrintData2, X, iCount, lasIPAddresses()

	If Not GetStoreStationPrinter(Session("StoreID"), Request.ServerVariables("REMOTE_ADDR"), sPrinter) Then
		If GetStoreMakeLinePrinters(Session("StoreID"), lasIPAddresses) Then
			If lasIPAddresses(0) <> "" Then
				sPrinter = lasIPAddresses(0)
			Else
				Response.Redirect("/error.asp?err=Payment accepted but could not determine where to print the receipt.")
			End If
		Else
			Response.Redirect("/error.asp?err=Payment accepted but could not determine where to print the receipt.")
		End If
	End If
	
	' Start a new line
	sPrintData = CHR(10)
	' Add logo
	sPrintData = sPrintData & CHR(27) & CHR(97) & CHR(1) & CHR(28) & CHR(112) & CHR(1) & CHR(0) & CHR(27) & CHR(97) & CHR(0)
	sPrintData = sPrintData & CHR(10) & CHR(27) & CHR(33) & CHR(14) 'Specify Font Size (last chr)
	sPrintData = sPrintData & CHR(27) & CHR(97) & CHR(1) 'Center Justify
	sPrintData = sPrintData & "vitos.com  #" & Session("StoreID")  & CHR(10) & CHR(10)
	If Request("Action") = "Apply" Then
		sPrintData = sPrintData & "Payment On Account Receipt" & CHR(10) & CHR(10)
	Else
		sPrintData = sPrintData & "Pay In Receipt" & CHR(10) & CHR(10)
	End If
	sPrintData = sPrintData & FormatDateTime(now,3) & " " & FormatDateTime(now,1) & CHR(10) & CHR(10)
	sPrintData = sPrintData & "Received By: " & GetEmployeeShortName(Session("EmpID")) & CHR(10)
	sPrintData = sPrintData & "-----------------------------------------" & CHR(10)
	sPrintData = sPrintData & "Received From: " & gsPayInFrom & CHR(10)
	sPrintData = sPrintData & "Payment Method: Credit Card" & CHR(10)
	sPrintData = sPrintData & "Amount Paid: " & FormatCurrency(gnPayAmount) & CHR(10)
	sPrintData = sPrintData & CHR(27) & CHR(97) & CHR(0) 'Left Justify
	sPrintData = sPrintData & "-----------------------------------------" & CHR(10) 
	If Request("Action") = "Apply" Then
		For intItem = 1 To intCounter
			If Request.Form("Discount-" & intItem) = "" Then
				Discount = 0
			Else
				Discount = CDbl(Request.Form("Discount-" & intItem))
			End If

			If (CDbl(Request.Form("Paying-" & intItem)) + Discount) >0 Then
				sPrintData = sPrintData & Left("Order # " & Request.Form("OrderID-" & intItem) & Space(10), 18) & "      "
				sPrintData = sPrintData & "Paid $" & Right(Space(10) & FormatNumber(CDbl(Request.Form("Paying-" & intItem)), 2), 10) & CHR(10)
				If Discount <> 0 Then
					sPrintData = sPrintData & Space(20) & "Discount $" & Right(Space(10) & FormatNumber(Discount, 2), 10) & CHR(10)
				End If
			End If
		Next
	Else
		sPrintData = sPrintData & "Category: "& GetPayInOutCategory(gnPayInOutCategory) & CHR(10)
		If Len(gsPayMemo) > 0 Then
			sPrintData = sPrintData & "Memo: " & gsPayMemo & CHR(10) 
		End If
	End If
	sPrintData = sPrintData & "-----------------------------------------" & CHR(10) & CHR(10)
	sPrintData = sPrintData & CHR(10)
	sPrintData = sPrintData & "Signature: ________________________" & CHR(10)
	sPrintData = sPrintData & "      I agree to above total amount" & CHR(10)
	sPrintData = sPrintData & "      as per card issuer agreement." & CHR(10)
	sPrintData = sPrintData & "Credit Card Auth #: " & gsPaymentReference & CHR(10)
	sPrintData = sPrintData & CHR(10)
	sPrintData = sPrintData & CHR(27) & CHR(97) & CHR(1) ' Center
	sPrintData2 = sPrintData
	sPrintData = sPrintData & CHR(10) & "Find Us on Facebook" & CHR(10)
	sPrintData = sPrintData & "vitos.com/facebook" & CHR(10)
	sPrintData = sPrintData & CHR(10) & "*** CUSTOMER COPY ***" & CHR(10)
	sPrintData2 = sPrintData2 & CHR(10) & "*** STORE COPY ***" & CHR(10)
	For x = 1 To 9 '# of returns
		sPrintData = sPrintData & CHR(10) 
		sPrintData2 = sPrintData2 & CHR(10) 
	Next 
	sPrintData = sPrintData & CHR(27) & CHR(97) & CHR(0)
	sPrintData2 = sPrintData2 & CHR(27) & CHR(97) & CHR(0)
	sPrintData = sPrintData & CHR(29) & CHR(86) & CHR(1) 'Cut
	sPrintData2 = sPrintData2 & CHR(29) & CHR(86) & CHR(1) 'Cut

	If Not SendToPrinter(sPrinter, sPrintData) Then
		Response.Redirect("/error.asp?err=Payment accepted but could not print the receipt.")
	End If
	If Not SendToPrinter(sPrinter, sPrintData2) Then
		Response.Redirect("/error.asp?err=Payment accepted but could not print the receipt.")
	End If
	
	'----------------------------------------------------- END PRINT RECEIPT --------------------------------------------------------------
End If

If Session("SecurityID") > 1 And gnPaymentTypeID <> 3 And ((Left(Request.ServerVariables("remote_addr"), 9) = "10.0.254." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.1." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.2.") OR Session("Swipe")) Then
	If GetStoreStationCashDrawer(gnStoreID, gsIPAddress, gsPrinterIPAddress, gbIsCashDrawer2) Then
		' Pop cash drawer
		If gbIsCashDrawer2 Then
			SendToPrinter gsPrinterIPAddress, CHR(27) & CHR(112) & CHR(1) & CHR(60) & CHR(60)
		Else
			SendToPrinter gsPrinterIPAddress, CHR(27) & CHR(112) & CHR(0) & CHR(60) & CHR(60)
		End If
	Else
		If Len(gsDBErrorMessage) > 0 Then
			Response.Redirect("/error.asp?err=" & Server.URLEncode(gsDBErrorMessage))
'		Else
'			Response.Redirect("/error.asp?err=" & Server.URLEncode("Order Accepted but No Cash Drawer Has Not Been Defined For This Station"))
		End If
	End If
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
<script src="/include2/redirect2.js" type="text/javascript"></script>
<script type="text/javascript">
<!--
var ie4=document.all;

function resetRedirect() {
	var loRedirectDiv;
	
	loRedirectDiv = ie4? eval("document.all.redirect") : document.getElementById("redirect");
	loRedirectDiv.innerHTML = <%=gnRedirectTime%>;
}
//-->
</script>
</head>

<body onload="clockInit(clockLocalStartTime, clockServerStartTime); clockOnLoad(); redirect('/default.asp')" onunload="clockOnUnload()">

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
						<span id="ClockTime" onclick="clockToggleSeconds()"><%=clockTimeString(Hour(gDate), Minute(gDate), Second(gDate))%></span> | 
						<span class="counter" id="redirect"><%=gnRedirectTime%></span>
					</div>
				</td>
			</tr>
			<tr height="733">
				<td valign="top" width="1010">
					<table cellpadding="0" cellspacing="0" width="1010" height="723">
						<tr>
							<td align="center" valign="top" width="1010">
<%
If gdOrderTotal < gdTenderTotal Then
%>
								<p><font size="72"><strong>CHANGE DUE:<br/><%=FormatCurrency(gdTenderTotal - gdOrderTotal)%></strong></font></p>
<%
End If
%>
								<p>&nbsp;</p>
								<p>&nbsp;</p>
<%
If Session("SecurityID") > 1 And gnPaymentTypeID < 3 And ((Left(Request.ServerVariables("remote_addr"), 9) = "10.0.254." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.1." Or Left(Request.ServerVariables("remote_addr"), 10) = "192.168.2.") OR Session("Swipe")) Then
%>
								<p><font size="72"><strong>CLOSE DRAWER</strong></font></p>
<%
Else
	If gbSignaturePrint Then
%>
								<p><font size="72"><strong>SALE COMPLETE</strong></font></p>
<%
	Else
%>
								<p><font size="72"><strong>ORDER SUBMITTED</strong></font></p>
<%
	End If
End If
%>
								<button style="width: 680px;" onclick="window.location = 'neworder.asp'">Done</button>
<%
If Request("o").Count <> 0 And gnCustomerID <> 1 Then
%>
								<p></p>
								<button style="width: 680px;" onclick="window.location = '/custmaint/editcustomer.asp?o=<%=gnOrderID%>&c=<%=gnCustomerID%>&a=<%=gnAddressID%>'">Edit Customer</button>
<%
End If
%>
							</td>
						</tr>
					</table>
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
