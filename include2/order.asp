<%
' **************************************************************************
' File: order.asp
' Purpose: Functions for order related activities.
' Created: 7/19/2011 - TAM
' Description:
'	Include this file on any page where order data is manipulated.
'	This file includes the following functions: GetOrderTypeDescription,
'		GetOrderDetails, GetOrderLines, GetOrderLineDetails, GetOrderLineItems,
'		GetOrderLineToppers, GetOrderLineFreeSides, GetOrderLineAddSides,
'		CreateOrder, CreateOrderLine, CreateOrderLineItem, CreateOrderLineTopper,
'		CreateOrderLineSide, DeleteOrderLine, DeleteOrder, CancelOrder, SubmitOrder,
'		SetOrderEmployee, SetOrderNotes, SetOrderType, SetOrderCustomer, SetOrderCustomerID,
'		SetOrderCustomerName, SetOrderAddress, SetOrderPayment, SetOrderPaymentType,
'		SetOrderCompleted, GetStoreOrderTypes, IsOrderTypeTaxable, SubmitHoldOrder,
'		GetHoldOrderTimes, ReleaseHoldOrders, DuplicateOrderLine, DuplicateOrder,
'		ResetHoldOrder, UpdateTransactionDate, GetVoidReasons, GetEditReasons.
'		GetMPOReasons, SetOrderEdited
'
' Revision History:
' 7/19/2011 - Created
' **************************************************************************

' **************************************************************************
' Function: GetOrderTypeDescription
' Purpose: Retrieves the descrition based on the OrderTypeID.
' Parameters:	pnOrderTypeID - The OrderTypeID to search for
' Return: Description as a string
' **************************************************************************
Function GetOrderTypeDescription(ByVal pnOrderTypeID)
	Dim lsRet, lsSQL, loRS
	
	lsRet = ""
	
	lsSQL = "select OrderTypeDescription from tlkpOrderTypes where OrderTypeID = " & pnOrderTypeID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lsRet = Trim(loRS("OrderTypeDescription"))
		End If
		
		DBCloseQuery loRS
	End If
	
	GetOrderTypeDescription = lsRet
End Function

' **************************************************************************
' Function: GetOrderDetails
' Purpose: Retrieves order details.
' Parameters:	pnOrderID - The OrderID to search for
'				pnSessionID - The ASP Session ID
'				psIPAddress - The IP Address
'				pnEmpID - The employee ID
'				psRefID - The reference ID
'				pdtTransaction - The transaction date
'				pdtSubmitDate - The submit date
'				pdtReleaseDate - The release date
'				pdtExpectedDate - The expected date
'				pnStoreID - The StoreID
'				pnCustomerID - The CustomerID
'				psCustomerName - The customer name
'				psCustomerPhone - The customer phone number
'				pnAddressID - The AddressID
'				pnOrderTypeID - The OrderTypeID
'				pbIsPaid - Flag indicating paid
'				pnPaymentTypeID - The PaymentTypeID
'				psPaymentReference - The payment reference
'				pnAccountID - The AccountID
'				pdDeliveryCharge - The delivery charge
'				pdDriveMoney - The driver money
'				pdTax - The sales tax
'				pdTax2 - The 2nd sales tax
'				pdTip - The tip
'				pnOrderStatusID - The OrderStatusID
'				psOrderNotes - The order notes
' Return: True if sucessful, False if not
' **************************************************************************
Function GetOrderDetails(ByVal pnOrderID, ByRef pnSessionID, ByRef psIPAddress, ByRef pnEmpID, ByRef psRefID, ByRef pdtTransactionDate, ByRef pdtSubmitDate, ByRef pdtReleaseDate, ByRef pdtExpectedDate, ByRef pnStoreID, ByRef pnCustomerID, ByRef psCustomerName, ByRef psCustomerPhone, ByRef pnAddressID, ByRef pnOrderTypeID, ByRef pbIsPaid, ByRef pnPaymentTypeID, ByRef psPaymentReference, ByRef pnAccountID, ByRef pdDeliveryCharge, ByRef pdDriverMoney, ByRef pdTax, ByRef pdTax2, ByRef pdTip, ByRef pnOrderStatusID, ByRef psOrderNotes)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	lsSQL = "select SessionID, IPAddress, EmpID, RefID, TransactionDate, SubmitDate, ReleaseDate, ExpectedDate, StoreID, CustomerID, CustomerName, CustomerPhone, AddressID, OrderTypeID, IsPaid, PaymentTypeID, PaymentReference, PaymentAuthorization, AccountID, DeliveryCharge, DriverMoney, Tax, Tax2, Tip, OrderStatusID, OrderNotes from tblOrders where OrderID = " & pnOrderID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lbRet = TRUE
			
			pnSessionID = loRS("SessionID")
			psIPAddress = Trim(loRS("IPAddress"))
			pnEmpID = loRS("EmpID")
			If IsNull(loRS("RefID")) Then
				psRefID = ""
			Else
				psRefID = Trim(loRS("RefID"))
			End If
			pdtTransactionDate = loRS("TransactionDate")
			If IsNull(loRS("SubmitDate")) Then
				pdtSubmitDate = DateValue("1/1/1900")
			Else
				pdtSubmitDate = loRS("SubmitDate")
			End If
			If IsNull(loRS("ReleaseDate")) Then
				pdtReleaseDate = DateValue("1/1/1900")
			Else
				pdtReleaseDate = loRS("ReleaseDate")
			End If
			If IsNull(loRS("ExpectedDate")) Then
				pdtExpectedDate = DateValue("1/1/1900")
			Else
				pdtExpectedDate = loRS("ExpectedDate")
			End If
			pnStoreID = loRS("StoreID")
			pnCustomerID = loRS("CustomerID")
			If IsNull(loRS("CustomerName")) Then
				psCustomerName = ""
			Else
				psCustomerName = Trim(loRS("CustomerName"))
			End If
			If IsNull(loRS("CustomerPhone")) Then
				psCustomerPhone = ""
			Else
				psCustomerPhone = Trim(loRS("CustomerPhone"))
			End If
			pnAddressID = loRS("AddressID")
			pnOrderTypeID = loRS("OrderTypeID")
			If loRS("IsPaid") <> 0 Then
				pbIsPaid = TRUE
			Else
				pbIsPaid = FALSE
			End If
			pnPaymentTypeID = loRS("PaymentTypeID")
			If IsNull(loRS("PaymentAuthorization")) Then
				psPaymentReference = ""
			Else
				psPaymentReference = Trim(loRS("PaymentAuthorization"))
			End If
			If IsNull(loRS("AccountID")) Then
				pnAccountID = 0
			Else
				pnAccountID = loRS("AccountID")
			End If
			pdDeliveryCharge = loRS("DeliveryCharge")
			pdDriverMoney = loRS("DriverMoney")
			pdTax = loRS("Tax")
			pdTax2 = loRS("Tax2")
			pdTip = loRS("Tip")
			pnOrderStatusID = loRS("OrderStatusID")
			If IsNull(loRS("OrderNotes")) Then
				psOrderNotes = ""
			Else
				psOrderNotes = Trim(loRS("OrderNotes"))
			End If
		End If
		
		DBCloseQuery loRS
	End If
	
	GetOrderDetails = lbRet
End Function

' **************************************************************************
' Function: GetOrderLines
' Purpose: Retrieves order details.
' Parameters:	pnOrderID - The OrderID to search for
'				panOrderLineIDs - Array of OrderLineIDs
'				pasDescriptions - Array of descriptions
'				panQuantity - Array of quantities
'				padCost - Array of costs
'				padDiscount - Array of discounts
' Return: True if sucessful, False if not
' **************************************************************************
Function GetOrderLines(ByVal pnOrderID, ByRef panOrderLineIDs, ByRef pasDescriptions, ByRef panQuantity, ByRef padCost, ByRef padDiscount)
	Dim lbRet, lsSQL, loRS, loRS2, lnPos, lsDescription, lbNeedSeparator, lnUnitID
	
	lbRet = FALSE
	
	lsSQL = "select OrderLineID, UnitID, SpecialtyID, SizeID, StyleID, Half1SauceID, Half2SauceID, Half1SauceModifierID, Half2SauceModifierID, OrderLineNotes, Quantity, Cost, Discount, Description, ShortDescription from tblOrderLines left outer join tblCoupons on tblOrderLines.CouponID = tblCoupons.CouponID where OrderID = " & pnOrderID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				lnUnitID = loRS("UnitID")
				
				If Not IsNull(loRS("SizeID")) Then
					lsDescription = GetSizeShortDescription(loRS("SizeID"))
				End If
				If Not IsNull(loRS("StyleID")) Then
					If Len(lsDescription) > 0 Then
						lsDescription = lsDescription & " "
					End If
					lsDescription = lsDescription & GetStyleShortDescription(loRS("StyleID"))
				End If
				If Not IsNull(loRS("SpecialtyID")) Then
					If Len(lsDescription) > 0 Then
						lsDescription = lsDescription & " "
					End If
					lsDescription = lsDescription & GetSpecialtyShortDescription(loRS("SpecialtyID"))
				End If
				If Not IsNull(loRS("UnitID")) Then
					If Len(lsDescription) > 0 Then
						lsDescription = lsDescription & " "
					End If
					lsDescription = lsDescription & GetUnitShortDescription(loRS("UnitID"))
				End If
				lsDescription = lsDescription & "<br/>"
				
				If Not IsNull(loRS("Half1SauceID")) Then
					If IsNull(loRS("Half2SauceID")) Then
						lsDescription = lsDescription & "1st Half Sauce: "
					Else
						If IsNull(loRS("Half1SauceModifierID")) And IsNull(loRS("Half2SauceModifierID")) Then
							If (loRS("Half1SauceID") = loRS("Half2SauceID")) Then
								lsDescription = lsDescription & "Whole Sauce: "
							Else
								lsDescription = lsDescription & "1st Half Sauce: "
							End If
						Else
							If IsNull(loRS("Half1SauceModifierID")) Or IsNull(loRS("Half2SauceModifierID")) Then
								lsDescription = lsDescription & "1st Half Sauce: "
							Else
								If (loRS("Half1SauceID") = loRS("Half2SauceID")) And (loRS("Half1SauceModifierID") = loRS("Half2SauceModifierID")) Then
									lsDescription = lsDescription & "Whole Sauce: "
								Else
									lsDescription = lsDescription & "1st Half Sauce: "
								End If
							End If
						End If
					End If
					lsDescription = lsDescription & GetSauceShortDescription(loRS("Half1SauceID"))
					If Not IsNull(loRS("Half1SauceModifierID")) Then
						lsDescription = lsDescription & " (" & GetSauceModifierShortDescription(loRS("Half1SauceModifierID")) & ")"
					End If
					lsDescription = lsDescription & "<br/>"
				End If
				If Not IsNull(loRS("Half2SauceID")) Then
					If IsNull(loRS("Half1SauceID")) Then
						lsDescription = lsDescription & "2nd Half Sauce: " & GetSauceShortDescription(loRS("Half2SauceID"))
						If Not IsNull(loRS("Half2SauceModifierID")) Then
							lsDescription = lsDescription & " (" & GetSauceModifierShortDescription(loRS("Half2SauceModifierID")) & ")"
						End If
						lsDescription = lsDescription & "<br/>"
					Else
						If IsNull(loRS("Half1SauceModifierID")) And IsNull(loRS("Half2SauceModifierID")) Then
							If (loRS("Half1SauceID") <> loRS("Half2SauceID")) Then
								lsDescription = lsDescription & "2st Half Sauce: " & GetSauceShortDescription(loRS("Half2SauceID"))
								If Not IsNull(loRS("Half2SauceModifierID")) Then
									lsDescription = lsDescription & " (" & GetSauceModifierShortDescription(loRS("Half2SauceModifierID")) & ")"
								End If
								lsDescription = lsDescription & "<br/>"
							End If
						Else
							If IsNull(loRS("Half1SauceModifierID")) Or IsNull(loRS("Half2SauceModifierID")) Then
								lsDescription = lsDescription & "2nd Half Sauce: " & GetSauceShortDescription(loRS("Half2SauceID"))
								If Not IsNull(loRS("Half2SauceModifierID")) Then
									lsDescription = lsDescription & " (" & GetSauceModifierShortDescription(loRS("Half2SauceModifierID")) & ")"
								End If
								lsDescription = lsDescription & "<br/>"
							Else
								If (loRS("Half1SauceID") <> loRS("Half2SauceID")) Or (loRS("Half1SauceModifierID") <> loRS("Half2SauceModifierID")) Then
									lsDescription = lsDescription & "2nd Half Sauce: " & GetSauceShortDescription(loRS("Half2SauceID"))
									If Not IsNull(loRS("Half2SauceModifierID")) Then
										lsDescription = lsDescription & " (" & GetSauceModifierShortDescription(loRS("Half2SauceModifierID")) & ")"
									End If
									lsDescription = lsDescription & "<br/>"
								End If
							End If
						End If
					End If
				End If
				
				lsSQL = "select tblOrderLineItems.ItemID from tblOrderLineItems inner join trelUnitItems on tblOrderLineItems.ItemID = trelUnitItems.ItemID and trelUnitItems.UnitID = " & lnUnitID & " where OrderLineID = " & loRS("OrderLineID") & " and HalfID = 0 order by UnitItemPrintOrder, tblOrderLineItems.ItemID"
				If DBOpenQuery(lsSQL, FALSE, loRS2) Then
					If Not loRS2.bof And Not loRS2.eof Then
						lsDescription = lsDescription & "Whole Items: "
						
						lbNeedSeparator = FALSE
						Do While Not loRS2.eof
							If lbNeedSeparator Then
								lsDescription = lsDescription & ", "
							End If
							lsDescription = lsDescription & GetItemShortDescription(loRS2("ItemID"))
							lbNeedSeparator = TRUE
							loRS2.MoveNext
						Loop
						
						lsDescription = lsDescription & "<br/>"
					End If
					
					DBCloseQuery loRS2
				End If
				
				lsSQL = "select tblOrderLineItems.ItemID from tblOrderLineItems inner join trelUnitItems on tblOrderLineItems.ItemID = trelUnitItems.ItemID and trelUnitItems.UnitID = " & lnUnitID & " where OrderLineID = " & loRS("OrderLineID") & " and HalfID = 1 order by UnitItemPrintOrder, tblOrderLineItems.ItemID"
				If DBOpenQuery(lsSQL, FALSE, loRS2) Then
					If Not loRS2.bof And Not loRS2.eof Then
						lsDescription = lsDescription & "1st Half Items: "
						
						lbNeedSeparator = FALSE
						Do While Not loRS2.eof
							If lbNeedSeparator Then
								lsDescription = lsDescription & ", "
							End If
							lsDescription = lsDescription & GetItemShortDescription(loRS2("ItemID"))
							lbNeedSeparator = TRUE
							loRS2.MoveNext
						Loop
						
						lsDescription = lsDescription & "<br/>"
					End If
					
					DBCloseQuery loRS2
				End If
				
				lsSQL = "select tblOrderLineItems.ItemID from tblOrderLineItems inner join trelUnitItems on tblOrderLineItems.ItemID = trelUnitItems.ItemID and trelUnitItems.UnitID = " & lnUnitID & " where OrderLineID = " & loRS("OrderLineID") & " and HalfID = 2 order by UnitItemPrintOrder, tblOrderLineItems.ItemID"
				If DBOpenQuery(lsSQL, FALSE, loRS2) Then
					If Not loRS2.bof And Not loRS2.eof Then
						lsDescription = lsDescription & "2nd Half Items: "
						
						lbNeedSeparator = FALSE
						Do While Not loRS2.eof
							If lbNeedSeparator Then
								lsDescription = lsDescription & ", "
							End If
							lsDescription = lsDescription & GetItemShortDescription(loRS2("ItemID"))
							lbNeedSeparator = TRUE
							loRS2.MoveNext
						Loop
						
						lsDescription = lsDescription & "<br/>"
					End If
					
					DBCloseQuery loRS2
				End If
				
				lsSQL = "select tblOrderLineItems.ItemID from tblOrderLineItems inner join trelUnitItems on tblOrderLineItems.ItemID = trelUnitItems.ItemID and trelUnitItems.UnitID = " & lnUnitID & " where OrderLineID = " & loRS("OrderLineID") & " and HalfID = 3 order by UnitItemPrintOrder, tblOrderLineItems.ItemID"
				If DBOpenQuery(lsSQL, FALSE, loRS2) Then
					If Not loRS2.bof And Not loRS2.eof Then
						lsDescription = lsDescription & "On Side Items: "
						
						lbNeedSeparator = FALSE
						Do While Not loRS2.eof
							If lbNeedSeparator Then
								lsDescription = lsDescription & ", "
							End If
							lsDescription = lsDescription & GetItemShortDescription(loRS2("ItemID"))
							lbNeedSeparator = TRUE
							loRS2.MoveNext
						Loop
						
						lsDescription = lsDescription & "<br/>"
					End If
					
					DBCloseQuery loRS2
				End If
				
				lsSQL = "select TopperID from tblOrderLineToppers where OrderLineID = " & loRS("OrderLineID") & " and TopperHalfID = 0"
				If DBOpenQuery(lsSQL, FALSE, loRS2) Then
					If Not loRS2.bof And Not loRS2.eof Then
						lsDescription = lsDescription & "Whole Toppers: "
						
						lbNeedSeparator = FALSE
						Do While Not loRS2.eof
							If lbNeedSeparator Then
								lsDescription = lsDescription & ", "
							End If
							lsDescription = lsDescription & GetTopperShortDescription(loRS2("TopperID"))
							lbNeedSeparator = TRUE
							loRS2.MoveNext
						Loop
						
						lsDescription = lsDescription & "<br/>"
					End If
					
					DBCloseQuery loRS2
				End If
				
				lsSQL = "select TopperID from tblOrderLineToppers where OrderLineID = " & loRS("OrderLineID") & " and TopperHalfID = 1"
				If DBOpenQuery(lsSQL, FALSE, loRS2) Then
					If Not loRS2.bof And Not loRS2.eof Then
						lsDescription = lsDescription & "1st Half Toppers: "
						
						lbNeedSeparator = FALSE
						Do While Not loRS2.eof
							If lbNeedSeparator Then
								lsDescription = lsDescription & ", "
							End If
							lsDescription = lsDescription & GetTopperShortDescription(loRS2("TopperID"))
							lbNeedSeparator = TRUE
							loRS2.MoveNext
						Loop
						
						lsDescription = lsDescription & "<br/>"
					End If
					
					DBCloseQuery loRS2
				End If
				
				lsSQL = "select TopperID from tblOrderLineToppers where OrderLineID = " & loRS("OrderLineID") & " and TopperHalfID = 2"
				If DBOpenQuery(lsSQL, FALSE, loRS2) Then
					If Not loRS2.bof And Not loRS2.eof Then
						lsDescription = lsDescription & "2nd Half Toppers: "
						
						lbNeedSeparator = FALSE
						Do While Not loRS2.eof
							If lbNeedSeparator Then
								lsDescription = lsDescription & ", "
							End If
							lsDescription = lsDescription & GetTopperShortDescription(loRS2("TopperID"))
							lbNeedSeparator = TRUE
							loRS2.MoveNext
						Loop
						
						lsDescription = lsDescription & "<br/>"
					End If
					
					DBCloseQuery loRS2
				End If
				
				lsSQL = "select TopperID from tblOrderLineToppers where OrderLineID = " & loRS("OrderLineID") & " and TopperHalfID = 3"
				If DBOpenQuery(lsSQL, FALSE, loRS2) Then
					If Not loRS2.bof And Not loRS2.eof Then
						lsDescription = lsDescription & "On Side Toppers: "
						
						lbNeedSeparator = FALSE
						Do While Not loRS2.eof
							If lbNeedSeparator Then
								lsDescription = lsDescription & ", "
							End If
							lsDescription = lsDescription & GetTopperShortDescription(loRS2("TopperID"))
							lbNeedSeparator = TRUE
							loRS2.MoveNext
						Loop
						
						lsDescription = lsDescription & "<br/>"
					End If
					
					DBCloseQuery loRS2
				End If
				
				lsSQL = "select SideID from tblOrderLineSides where OrderLineID = " & loRS("OrderLineID") & " and IsFreeSide = 1"
				If DBOpenQuery(lsSQL, FALSE, loRS2) Then
					If Not loRS2.bof And Not loRS2.eof Then
						lsDescription = lsDescription & "Included Sides: "
						
						lbNeedSeparator = FALSE
						Do While Not loRS2.eof
							If lbNeedSeparator Then
								lsDescription = lsDescription & ", "
							End If
							lsDescription = lsDescription & GetSideShortDescription(loRS2("SideID"))
							lbNeedSeparator = TRUE
							loRS2.MoveNext
						Loop
						
						lsDescription = lsDescription & "<br/>"
					End If
					
					DBCloseQuery loRS2
				End If
				
				lsSQL = "select SideID from tblOrderLineSides where OrderLineID = " & loRS("OrderLineID") & " and IsFreeSide = 0"
				If DBOpenQuery(lsSQL, FALSE, loRS2) Then
					If Not loRS2.bof And Not loRS2.eof Then
						lsDescription = lsDescription & "Added Sides: "
						
						lbNeedSeparator = FALSE
						Do While Not loRS2.eof
							If lbNeedSeparator Then
								lsDescription = lsDescription & ", "
							End If
							lsDescription = lsDescription & GetSideShortDescription(loRS2("SideID"))
							lbNeedSeparator = TRUE
							loRS2.MoveNext
						Loop
						
						lsDescription = lsDescription & "<br/>"
					End If
					
					DBCloseQuery loRS2
				End If
				
				If Not IsNull(loRS("OrderLineNotes")) Then
					lsDescription = lsDescription & Trim(loRS("OrderLineNotes")) & "<br/>"
				End If
				
				If loRS("Quantity") > 1 Then
					lsDescription = lsDescription & loRS("Quantity") & " @ "
				End If
				lsDescription = lsDescription & FormatCurrency((loRS("Cost") - loRS("Discount")))
				If Not IsNull(loRS("Description")) Then
					lsDescription = lsDescription & " (" & loRS("Description") & ")"
				End If
				
				ReDim Preserve panOrderLineIDs(lnPos), pasDescriptions(lnPos), panQuantity(lnPos), padCost(lnPos), padDiscount(lnPos)
				
				panOrderLineIDs(lnPos) = loRS("OrderLineID")
				pasDescriptions(lnPos) = lsDescription
				panQuantity(lnPos) = loRS("Quantity")
				padCost(lnPos) = loRS("Cost")
				padDiscount(lnPos) = loRS("Discount")
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panOrderLineIDs(0), pasDescriptions(0), panQuantity(0), padCost(0), padDiscount(0)
			panOrderLineIDs(0) = 0
			pasDescriptions(0) = ""
			panQuantity(0) = 0
			padCost(0) = 0.00
			padDiscount(0) = 0.00
		End If
		
		DBCloseQuery loRS
	End If
	
	GetOrderLines = lbRet
End Function

' **************************************************************************
' Function: GetOrderLineDetails
' Purpose: Retrieves order line details.
' Parameters:	pnOrderLineID - The OrderLineID to search for
'				pnUnitID - The UnitID
'				pnSpecialtyID - The SpecialtyID
'				pnSizeID - The SizeID
'				pnStyleID - The StyleID
'				pnHalf1SauceID - The Half1SauceID
'				pnHalf2SauceID - The Half2SauceID
'				pnHalf1SauceModifierID - The Half1SauceModifierID
'				pnHalf2SauceModifierID - The Half2SauceModifierID
'				psOrderLineNotes - The order line notes
'				pnQuantity - The order line quantity
'				pdCost - The order line cost
'				pdDiscount - The order line discount
'				pnCouponID - The CouponID
' Return: True if sucessful, False if not
' **************************************************************************
Function GetOrderLineDetails(ByVal pnOrderLineID, ByRef pnUnitID, ByRef pnSpecialtyID, ByRef pnSizeID, ByRef pnStyleID, ByRef pnHalf1SauceID, ByRef pnHalf2SauceID, ByRef pnHalf1SauceModifierID, ByRef pnHalf2SauceModifierID, ByRef psOrderLineNotes, ByRef pnQuantity, ByRef pdCost, ByRef pdDiscount, ByRef pnCouponID)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	lsSQL = "select UnitID, SpecialtyID, SizeID, StyleID, Half1SauceID, Half2SauceID, Half1SauceModifierID, Half2SauceModifierID, OrderLineNotes, Quantity, Cost, Discount, CouponID from tblOrderLines where OrderLineID = " & pnOrderLineID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lbRet = TRUE
			
			pnUnitID = loRS("UnitID")
			If IsNull(loRS("SpecialtyID")) Then
				pnSpecialtyID = 0
			Else
				pnSpecialtyID = loRS("SpecialtyID")
			End If
			If IsNull(loRS("SizeID")) Then
				pnSizeID = 0
			Else
				pnSizeID = loRS("SizeID")
			End If
			If IsNull(loRS("StyleID")) Then
				pnStyleID = 0
			Else
				pnStyleID = loRS("StyleID")
			End If
			If IsNull(loRS("Half1SauceID")) Then
				pnHalf1SauceID = 0
			Else
				pnHalf1SauceID = loRS("Half1SauceID")
			End If
			If IsNull(loRS("Half2SauceID")) Then
				pnHalf2SauceID = 0
			Else
				pnHalf2SauceID = loRS("Half2SauceID")
			End If
			If IsNull(loRS("Half1SauceModifierID")) Then
				pnHalf1SauceModifierID = 0
			Else
				pnHalf1SauceModifierID = loRS("Half1SauceModifierID")
			End If
			If IsNull(loRS("Half2SauceModifierID")) Then
				pnHalf2SauceModifierID = 0
			Else
				pnHalf2SauceModifierID = loRS("Half2SauceModifierID")
			End If
			If IsNull(loRS("OrderLineNotes")) Then
				psOrderLineNotes = ""
			Else
				psOrderLineNotes = loRS("OrderLineNotes")
			End If
			pnQuantity = loRS("Quantity")
			pdCost = loRS("Cost")
			pdDiscount = loRS("Discount")
			If IsNull(loRS("CouponID")) Then
				pnCouponID = 0
			Else
				pnCouponID = loRS("CouponID")
			End If
		End If
		
		DBCloseQuery loRS
	End If
	
	GetOrderLineDetails = lbRet
End Function

' **************************************************************************
' Function: GetOrderLineItems
' Purpose: Retrieves order line items.
' Parameters:	pnOrderLineID - The OrderLineID to search for
'				pnUnitID - The UnitID to search for
'				panOrderLineItemIDs - Array of OrderLineItemIDs
'				panItemIDs - Array of ItemIDs
'				panHalfIDs - Array of HalfIDs
'				pasItemDescriptions - Array of item descriptions
'				pasItemShortDescriptions - Array of item short descriptions
' Return: True if sucessful, False if not
' **************************************************************************
Function GetOrderLineItems(ByVal pnOrderLineID, ByVal pnUnitID, ByRef panOrderLineItemIDs, ByRef panItemIDs, ByRef panHalfIDs, ByRef pasItemDescriptions, ByRef pasItemShortDescriptions)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select OrderLineItemID, tblOrderLineItems.ItemID, HalfID, ItemDescription, ItemShortDescription from tblOrderLineItems inner join tblItems on tblOrderLineItems.ItemID = tblItems.ItemID inner join trelUnitItems on tblOrderLineItems.ItemID = trelUnitItems.ItemID and trelUnitItems.UnitID = " & pnUnitID & " where OrderLineID = " & pnOrderLineID & " order by HalfID, UnitItemPrintOrder, tblOrderLineItems.ItemID"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve panOrderLineItemIDs(lnPos), panItemIDs(lnPos), panHalfIDs(lnPos), pasItemDescriptions(lnPos), pasItemShortDescriptions(lnPos)
				
				panOrderLineItemIDs(lnPos) = loRS("OrderLineItemID")
				panItemIDs(lnPos) = loRS("ItemID")
				panHalfIDs(lnPos) = loRS("HalfID")
				pasItemDescriptions(lnPos) = Trim(loRS("ItemDescription"))
				pasItemShortDescriptions(lnPos) = Trim(loRS("ItemShortDescription"))
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panOrderLineItemIDs(0), panItemIDs(0), panHalfIDs(0), pasItemDescriptions(0), pasItemShortDescriptions(0)
			
			panOrderLineItemIDs(0) = 0
			panItemIDs(0) = 0
			panHalfIDs(0) = 0
			pasItemDescriptions(0) = ""
			pasItemShortDescriptions(0) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetOrderLineItems = lbRet
End Function

' **************************************************************************
' Function: GetOrderLineToppers
' Purpose: Retrieves order line toppers.
' Parameters:	pnOrderLineID - The OrderLineID to search for
'				panOrderLineTopperIDs - Array of OrderLineTopperIDs
'				panTopperIDs - Array of TopperIDs
'				panTopperHalfIDs - Array of TopperHalfIDs
'				pasTopperDescriptions - Array of topper descriptions
'				pasTopperShortDescriptions - Array of topper short descriptions
' Return: True if sucessful, False if not
' **************************************************************************
Function GetOrderLineToppers(ByVal pnOrderLineID, ByRef panOrderLineTopperIDs, ByRef panTopperIDs, ByRef panTopperHalfIDs, ByRef pasTopperDescriptions, ByRef pasTopperShortDescriptions)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select OrderLineTopperID, tblOrderLineToppers.TopperID, TopperHalfID, TopperDescription, TopperShortDescription from tblOrderLineToppers inner join tblTopper on tblOrderLineToppers.TopperID = tblTopper.TopperID where OrderLineID = " & pnOrderLineID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve panOrderLineTopperIDs(lnPos), panTopperIDs(lnPos), panTopperHalfIDs(lnPos), pasTopperDescriptions(lnPos), pasTopperShortDescriptions(lnPos)
				
				panOrderLineTopperIDs(lnPos) = loRS("OrderLineTopperID")
				panTopperIDs(lnPos) = loRS("TopperID")
				panTopperHalfIDs(lnPos) = loRS("TopperHalfID")
				pasTopperDescriptions(lnPos) = Trim(loRS("TopperDescription"))
				pasTopperShortDescriptions(lnPos) = Trim(loRS("TopperShortDescription"))
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panOrderLineTopperIDs(0), panTopperIDs(0), panTopperHalfIDs(0), pasTopperDescriptions(0), pasTopperShortDescriptions(0)
			
			panOrderLineTopperIDs(0) = 0
			panTopperIDs(0) = 0
			panTopperHalfIDs(0) = 0
			pasTopperDescriptions(0) = ""
			pasTopperShortDescriptions(0) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetOrderLineToppers = lbRet
End Function

' **************************************************************************
' Function: GetOrderLineFreeSides
' Purpose: Retrieves order line free sides.
' Parameters:	pnOrderLineID - The OrderLineID to search for
'				panOrderLineSideIDs - Array of OrderLineSideIDs
'				panFreeSideIDs - Array of SideIDs
'				pasFreeSideDescriptions - Array of side descriptions
'				pasFreeSideShortDescriptions - Array of side short descriptions
' Return: True if sucessful, False if not
' **************************************************************************
Function GetOrderLineFreeSides(ByVal pnOrderLineID, ByRef panOrderLineSideIDs, ByRef panFreeSideIDs, ByRef pasFreeSideDescriptions, ByRef pasFreeSideShortDescriptions)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select OrderLineSideID, tblOrderLineSides.SideID, SideDescription, SideShortDescription from tblOrderLineSides inner join tblSides on tblOrderLineSides.SideID = tblSides.SideID where OrderLineID = " & pnOrderLineID & " and IsFreeSide <> 0"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve panOrderLineSideIDs(lnPos), panFreeSideIDs(lnPos), pasFreeSideDescriptions(lnPos), pasFreeSideShortDescriptions(lnPos)
				
				panOrderLineSideIDs(lnPos) = loRS("OrderLineSideID")
				panFreeSideIDs(lnPos) = loRS("SideID")
				pasFreeSideDescriptions(lnPos) = Trim(loRS("SideDescription"))
				pasFreeSideShortDescriptions(lnPos) = Trim(loRS("SideShortDescription"))
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panOrderLineSideIDs(0), panFreeSideIDs(0), pasFreeSideDescriptions(0), pasFreeSideShortDescriptions(0)
			
			panOrderLineSideIDs(0) = 0
			panFreeSideIDs(0) = 0
			pasFreeSideDescriptions(0) = ""
			pasFreeSideShortDescriptions(0) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetOrderLineFreeSides = lbRet
End Function

' **************************************************************************
' Function: GetOrderLineAddSides
' Purpose: Retrieves order line additional sides.
' Parameters:	pnOrderLineID - The OrderLineID to search for
'				panOrderLineSideIDs - Array of OrderLineSideIDs
'				panAddSideIDs - Array of SideIDs
'				pasAddSideDescriptions - Array of side descriptions
'				pasAddSideShortDescriptions - Array of side short descriptions
' Return: True if sucessful, False if not
' **************************************************************************
Function GetOrderLineAddSides(ByVal pnOrderLineID, ByRef panOrderLineSideIDs, ByRef panAddSideIDs, ByRef pasAddSideDescriptions, ByRef pasAddSideShortDescriptions)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select OrderLineSideID, tblOrderLineSides.SideID, SideDescription, SideShortDescription from tblOrderLineSides inner join tblSides on tblOrderLineSides.SideID = tblSides.SideID where OrderLineID = " & pnOrderLineID & " and IsFreeSide = 0"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve panOrderLineSideIDs(lnPos), panAddSideIDs(lnPos), pasAddSideDescriptions(lnPos), pasAddSideShortDescriptions(lnPos)
				
				panOrderLineSideIDs(lnPos) = loRS("OrderLineSideID")
				panAddSideIDs(lnPos) = loRS("SideID")
				pasAddSideDescriptions(lnPos) = Trim(loRS("SideDescription"))
				pasAddSideShortDescriptions(lnPos) = Trim(loRS("SideShortDescription"))
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panOrderLineSideIDs(0), panAddSideIDs(0), pasAddSideDescriptions(0), pasAddSideShortDescriptions(0)
			
			panOrderLineSideIDs(0) = 0
			panAddSideIDs(0) = 0
			pasAddSideDescriptions(0) = ""
			pasAddSideShortDescriptions(0) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetOrderLineAddSides = lbRet
End Function

' **************************************************************************
' Function: CreateOrder
' Purpose: Creates an order.
' Parameters:	pnSessionID - The SessionID
'				psIPAddress - The IP Address
'				pnEmpID - The EmpID
'				psRefID - The referral ID
'				psTransactionDate - The transaction date
'				pnStoreID - The StoreID
'				pnCustomerID - The CustomerID
'				psCustomerName - The customer's name
'				psCustomerPhone - The customer's phone number
'				pnAddressID - The AddressID
'				pnOrderTypeID - The OrderTypeID
'				pdDeliveryCharge - The delivery charge
'				pdDriverMoney - The driver money
'				psOrderNotes - The order notes
' Return: The OrderID
' **************************************************************************
Function CreateOrder(ByVal pnSessionID, ByVal psIPAddress, ByVal pnEmpID, ByVal psRefID, ByVal psTransactionDate, ByVal pnStoreID, ByVal pnCustomerID, ByVal psCustomerName, ByVal psCustomerPhone, ByVal pnAddressID, ByVal pnOrderTypeID, ByVal pdDeliveryCharge, ByVal pdDriverMoney, ByVal psOrderNotes)
	Dim lnRet, lsSQL, loRS
	
	lnRet = 0
	
	lsSQL = "EXEC AddOrder @pSessionID = " & Session.SessionID & ", @pIPAddress = '" & Request.ServerVariables("REMOTE_ADDR") & "', @pEmpID = " & pnEmpID & ", @pRefID = "
	If Len(Trim(psRefID)) = 0 Then
		lsSQL = lsSQL & "NULL"
	Else
		lsSQL = lsSQL & "'" & DBCleanLiteral(Trim(psRefID)) & "'"
	End If
	lsSQL = lsSQL & ", @pTransactionDate = '" & DBCleanLiteral(psTransactionDate) & "', @pStoreID = " & pnStoreID & ", @pCustomerID = " & pnCustomerID & ", @pCustomerName = "
	If Len(Trim(psCustomerName)) = 0 Then
		lsSQL = lsSQL & "NULL"
	Else
		lsSQL = lsSQL & "'" & DBCleanLiteral(Trim(psCustomerName)) & "'"
	End If
	lsSQL = lsSQL & ", @pCustomerPhone = "
	If Len(Trim(psCustomerPhone)) = 0 Then
		lsSQL = lsSQL & "NULL"
	Else
		lsSQL = lsSQL & "'" & DBCleanLiteral(Trim(psCustomerPhone)) & "'"
	End If
	lsSQL = lsSQL & ", @pAddressID = " & pnAddressID & ", @pOrderTypeID = " & pnOrderTypeID & ", @pDeliveryCharge = " & pdDeliveryCharge & ", @pDriverMoney = " & pdDriverMoney & ", @pOrderNotes = "
	If Len(Trim(psOrderNotes)) = 0 Then
		lsSQL = lsSQL & "NULL"
	Else
		lsSQL = lsSQL & "'" & DBCleanLiteral(Trim(psOrderNotes)) & "'"
	End If
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lnRet = CLng(loRS(0))
		End If
		
		DBCloseQuery loRS
	End If
	
	CreateOrder = lnRet
End Function

' **************************************************************************
' Function: CreateOrderLine
' Purpose: Creates an order line.
' Parameters:	pnOrderID - The OrderID
'				pnUnitID - The UnitID
'				pnSpecialtyID - The SpecialtyID
'				pnSizeID - The SizeID
'				pnStyleID - The StyleID
'				pnHalf1SauceID - The Half1SauceID
'				pnHalf2SauceID - The Half2SauceID
'				pnHalf1SauceModifierID - The Half1SauceModifier
'				pnHalf2SauceModifierID - The Half2SauceModifier
'				psOrderLineNotes - The order line notes
'				pnQuantity - The order line quantity
' Return: The OrderLineID
' **************************************************************************
Function CreateOrderLine(ByVal pnOrderID, ByVal pnUnitID, ByVal pnSpecialtyID, ByVal pnSizeID, ByVal pnStyleID, ByVal pnHalf1SauceID, ByVal pnHalf2SauceID, ByVal pnHalf1SauceModifierID, ByVal pnHalf2SauceModifierID, ByVal psOrderLineNotes, ByVal pnQuantity)
	Dim lnRet, lsSQL, loRS
	
	lnRet = 0
	
	lsSQL = "EXEC AddOrderLine @pOrderID = " & pnOrderID & ", @pUnitID = " & pnUnitID & ", @pSpecialtyID = "
	If pnSpecialtyID = 0 Then
		lsSQL = lsSQL & "NULL"
	Else
		lsSQL = lsSQL & pnSpecialtyID
	End If
	lsSQL = lsSQL & ", @pSizeID = "
	If pnSizeID = 0 Then
		lsSQL = lsSQL & "NULL"
	Else
		lsSQL = lsSQL & pnSizeID
	End If
	lsSQL = lsSQL & ", @pStyleID = "
	If pnStyleID = 0 Then
		lsSQL = lsSQL & "NULL"
	Else
		lsSQL = lsSQL & pnStyleID
	End If
	lsSQL = lsSQL & ", @pHalf1SauceID = "
	If pnHalf1SauceID = 0 Then
		lsSQL = lsSQL & "NULL"
	Else
		lsSQL = lsSQL & pnHalf1SauceID
	End If
	lsSQL = lsSQL & ", @pHalf2SauceID = "
	If pnHalf2SauceID = 0 Then
		lsSQL = lsSQL & "NULL"
	Else
		lsSQL = lsSQL & pnHalf2SauceID
	End If
	lsSQL = lsSQL & ", @pHalf1SauceModifierID = "
	If pnHalf1SauceModifierID = 0 Then
		lsSQL = lsSQL & "NULL"
	Else
		lsSQL = lsSQL & pnHalf1SauceModifierID
	End If
	lsSQL = lsSQL & ", @pHalf2SauceModifierID = "
	If pnHalf2SauceModifierID = 0 Then
		lsSQL = lsSQL & "NULL"
	Else
		lsSQL = lsSQL & pnHalf2SauceModifierID
	End If
	lsSQL = lsSQL & ", @pOrderLineNotes = "
	If Len(Trim(psOrderLineNotes)) = 0 Then
		lsSQL = lsSQL & "NULL"
	Else
		lsSQL = lsSQL & "'" & DBCleanLiteral(Trim(psOrderLineNotes)) & "'"
	End If
	lsSQL = lsSQL & ", @pQuantity = " & pnQuantity & ", @pInternetDescription = NULL"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lnRet = CLng(loRS(0))
		End If
		
		DBCloseQuery loRS
	End If
	
	CreateOrderLine = lnRet
End Function

' **************************************************************************
' Function: CreateOrderLineItem
' Purpose: Creates an order line item.
' Parameters:	pnOrderLineID - The OrderLineID
'				pnItemID - The ItemID
'				pnHalfID - The HalfID
' Return: The OrderLineItemID
' **************************************************************************
Function CreateOrderLineItem(ByVal pnOrderLineID, ByVal pnItemID, ByVal pnHalfID)
	Dim lnRet, lsSQL, loRS
	
	lnRet = 0
	
	lsSQL = "EXEC AddOrderLineItem @pOrderLineID = " & pnOrderLineID & ", @pItemID = " & pnItemID & ", @pHalfID = " & pnHalfID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lnRet = CLng(loRS(0))
		End If
		
		DBCloseQuery loRS
	End If
	
	CreateOrderLineItem = lnRet
End Function

' **************************************************************************
' Function: CreateOrderLineTopper
' Purpose: Creates an order line topper.
' Parameters:	pnOrderLineID - The OrderLineID
'				pnTopperID - The TopperID
'				pnTopperHalfID - The TopperHalfID
' Return: The OrderLineTopperID
' **************************************************************************
Function CreateOrderLineTopper(ByVal pnOrderLineID, ByVal pnTopperID, ByVal pnTopperHalfID)
	Dim lnRet, lsSQL, loRS
	
	lnRet = 0
	
	lsSQL = "EXEC AddOrderLineTopper @pOrderLineID = " & pnOrderLineID & ", @pTopperID = " & pnTopperID & ", @pTopperHalfID = " & pnTopperHalfID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lnRet = CLng(loRS(0))
		End If
		
		DBCloseQuery loRS
	End If
	
	CreateOrderLineTopper = lnRet
End Function

' **************************************************************************
' Function: CreateOrderLineSide
' Purpose: Creates an order line side.
' Parameters:	pnOrderLineID - The OrderLineID
'				pnSideID - The SideID
'				pbIsFreeSide - Flag if free side
' Return: The OrderLineSideID
' **************************************************************************
Function CreateOrderLineSide(ByVal pnOrderLineID, ByVal pnSideID, ByVal pbIsFreeSide)
	Dim lnRet, lsSQL, loRS
	
	lnRet = 0
	
	lsSQL = "EXEC AddOrderLineSide @pOrderLineID = " & pnOrderLineID & ", @pSideID = " & pnSideID & ", @pIsFreeSide = "
	If pbIsFreeSide Then
		lsSQL = lsSQL & "1"
	Else
		lsSQL = lsSQL & "0"
	End If
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lnRet = CLng(loRS(0))
		End If
		
		DBCloseQuery loRS
	End If
	
	CreateOrderLineSide = lnRet
End Function

' **************************************************************************
' Function: DeleteOrderLine
' Purpose: Deletes an order line.
' Parameters:	pnOrderLineID - The OrderLineID
' Return: True if sucessful, False if not
' **************************************************************************
Function DeleteOrderLine(ByVal pnOrderLineID)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	lsSQL = "delete from tblOrderLines where OrderLineID = " & pnOrderLineID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If
	
	DeleteOrderLine = lbRet
End Function

' **************************************************************************
' Function: DeleteOrder
' Purpose: Deletes an order.
' Parameters:	pnOrderID - The OrderID
' Return: True if sucessful, False if not
' **************************************************************************
Function DeleteOrder(ByVal pnOrderID)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	lsSQL = "delete from tblOrders where OrderID = " & pnOrderID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If
	
	DeleteOrder = lbRet
End Function

' **************************************************************************
' Function: CancelOrder
' Purpose: Cancels an order.
' Parameters:	pnOrderID - The OrderID
'				pnEmpID - The EmpID
' Return: True if sucessful, False if not
' **************************************************************************
Function CancelOrder(ByVal pnOrderID, ByVal pnEmpID, ByVal psVoidReason)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	lsSQL = "update tblOrders set OrderStatusID = 11, VoidDate = getdate(), VoidEmpID = " & pnEmpID & ", VoidReason = '" & psVoidReason & "' where OrderID = " & pnOrderID
	If DBExecuteSQL(lsSQL) Then
		lsSQL = "delete from tblDelivery where OrderID = " & pnOrderID
		If DBExecuteSQL(lsSQL) Then
			lbRet = TRUE
		End If
	End If
	
	CancelOrder = lbRet
End Function

' **************************************************************************
' Function: SubmitOrder
' Purpose: Submits an order.
' Parameters:	pnOrderID - The OrderID
'				pnCustomerID - The CustomerID
'				pnAddressID = The AddressID
' Return: True if sucessful, False if not
' **************************************************************************
Function SubmitOrder(ByVal pnOrderID, ByVal pnCustomerID, ByVal pnAddressID)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	lsSQL = "update tblOrders set SubmitDate = getdate(), ReleaseDate = getdate(), OrderStatusID = 3 where OrderID = " & pnOrderID
	If DBExecuteSQL(lsSQL) Then
		If pnCustomerID = 1 Or pnAddressID = 1 Then
			lbRet = TRUE
		Else
			lsSQL = "update tblCustomers set PrimaryAddressID = " & pnAddressID & " where CustomerID = " & pnCustomerID
			If DBExecuteSQL(lsSQL) Then
				lbRet = TRUE
			End If
		End If
	End If
	
	SubmitOrder = lbRet
End Function

' **************************************************************************
' Function: SetOrderEmployee
' Purpose: Assigns an employee to an order.
' Parameters:	pnOrderID - The OrderLineID
'				pnEmpID - The EmpID
' Return: True if sucessful, False if not
' **************************************************************************
Function SetOrderEmployee(ByVal pnOrderID, ByVal pnEmpID)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	lsSQL = "update tblOrders set EmpID = " & pnEmpID & " where OrderID = " & pnOrderID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If
	
	SetOrderEmployee = lbRet
End Function

' **************************************************************************
' Function: SetOrderNotes
' Purpose: Assigns notes to an order.
' Parameters:	pnOrderID - The OrderLineID
'				psNotes - The notes
' Return: True if sucessful, False if not
' **************************************************************************
Function SetOrderNotes(ByVal pnOrderID, ByVal psNotes)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	lsSQL = "update tblOrders set OrderNotes = '" & DBCleanLiteral(psNotes) & "' where OrderID = " & pnOrderID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If
	
	SetOrderNotes = lbRet
End Function

' **************************************************************************
' Function: SetOrderType
' Purpose: Assigns a type to an order.
' Parameters:	pnOrderID - The OrderLineID
'				pnStoreID - The StoreID
'				pnOrderTypeID - The order type
'				pdDeliveryCharge - The delivery charge
'				pdDriverMoney - The driver money
' Return: True if sucessful, False if not
' **************************************************************************
Function SetOrderType(ByVal pnOrderID, ByVal pnStoreID, ByVal pnOrderTypeID, ByVal pdDeliveryCharge, ByVal pdDriverMoney)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	lsSQL = "update tblOrders set StoreID = " & pnStoreID & ", OrderTypeID = " & pnOrderTypeID & ", DeliveryCharge = " & pdDeliveryCharge & ", DriverMoney = " & pdDriverMoney & " where OrderID = " & pnOrderID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If
	
	SetOrderType = lbRet
End Function

' **************************************************************************
' Function: SetOrderCustomer
' Purpose: Assigns a customer to an order.
' Parameters:	pnOrderID - The OrderLineID
'				pnCustomerID - The order type
'				psCustomerName - The customer name
'				psCustomerPhone - The customer phone
' Return: True if sucessful, False if not
' **************************************************************************
Function SetOrderCustomer(ByVal pnOrderID, ByVal pnCustomerID, ByVal psCustomerName, ByVal psCustomerPhone)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	lsSQL = "update tblOrders set CustomerID = " & pnCustomerID & ", CustomerName = "
	If Len(psCustomerName) = 0 Then
		lsSQL = lsSQL & "NULL"
	Else
		lsSQL = lsSQL & "'" & DBCleanLiteral(psCustomerName) & "'"
	End If
	lsSQL = lsSQL & ", CustomerPhone = "
	If Len(psCustomerPhone) = 0 Then
		lsSQL = lsSQL & "NULL"
	Else
		lsSQL = lsSQL & "'" & DBCleanLiteral(psCustomerPhone) & "'"
	End If
	lsSQL = lsSQL & " where OrderID = " & pnOrderID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If
	
	SetOrderCustomer = lbRet
End Function

' **************************************************************************
' Function: SetOrderCustomerID
' Purpose: Assigns a CustomerID to an order.
' Parameters:	pnOrderID - The OrderLineID
'				pnCustomerID - The order type
' Return: True if sucessful, False if not
' **************************************************************************
Function SetOrderCustomerID(ByVal pnOrderID, ByVal pnCustomerID)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	lsSQL = "update tblOrders set CustomerID = " & pnCustomerID & " where OrderID = " & pnOrderID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If
	
	SetOrderCustomerID = lbRet
End Function

' **************************************************************************
' Function: SetOrderCustomerName
' Purpose: Assigns a customer name to an order.
' Parameters:	pnOrderID - The OrderLineID
'				psCustomerName - The customer name
' Return: True if sucessful, False if not
' **************************************************************************
Function SetOrderCustomerName(ByVal pnOrderID, ByVal psCustomerName)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	lsSQL = "update tblOrders set CustomerName = "
	If Len(psCustomerName) = 0 Then
		lsSQL = lsSQL & "NULL"
	Else
		lsSQL = lsSQL & "'" & DBCleanLiteral(psCustomerName) & "'"
	End If
	lsSQL = lsSQL & " where OrderID = " & pnOrderID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If
	
	SetOrderCustomerName = lbRet
End Function

' **************************************************************************
' Function: SetOrderAddress
' Purpose: Assigns an address to an order.
' Parameters:	pnOrderID - The OrderLineID
'				pnAddressID - The AddressID
' Return: True if sucessful, False if not
' **************************************************************************
Function SetOrderAddress(ByVal pnOrderID, ByVal pnAddressID)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	lsSQL = "update tblOrders set AddressID = " & pnAddressID & " where OrderID = " & pnOrderID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If
	
	SetOrderAddress = lbRet
End Function

' **************************************************************************
' Function: SetOrderPayment
' Purpose: Marks an order as paid.
' Parameters:	pnOrderID - The OrderLineID
'				pnPaymentTypeID - The PaymentTypeID
'				psPaymentReference - The PaymentReference
'				pnAccountID - The AccountID
'				pdTipAmount - The tip amount
'				pnPaymentEmpID - The EmpID who accepted the payment
' Return: True if sucessful, False if not
' **************************************************************************
Function SetOrderPayment(ByVal pnOrderID, ByVal pnPaymentTypeID, ByVal psPaymentReference, ByVal pnAccountID, ByVal pdTipAmount, ByVal pnPaymentEmpID)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	lsSQL = "update tblOrders set IsPaid = 1, PaymentTypeID = " & pnPaymentTypeID
	If Len(psPaymentReference) = 0 Then
		lsSQL = lsSQL & ", PaymentReference = NULL, PaymentAuthorization = NULL"
	Else
		lsSQL = lsSQL & ", PaymentReference = NULL, PaymentAuthorization = '" & psPaymentReference & "'"
	End If
	If pnAccountID = 0 Then
		lsSQL = lsSQL & ", AccountID = NULL"
	Else
		lsSQL = lsSQL & ", AccountID = " & pnAccountID
	End If
	lsSQL = lsSQL & ", Tip = " & pdTipAmount & ", PaidDate = getdate(), PaymentEmpID = " & pnPaymentEmpID & " where OrderID = " & pnOrderID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If
	
	SetOrderPayment = lbRet
End Function

' **************************************************************************
' Function: SetOrderPaymentType
' Purpose: Sets an order's planned payment without marking as paid.
' Parameters:	pnOrderID - The OrderID
'				pnPaymentTypeID - The PaymentTypeID
'				psPaymentReference - The PaymentReference
'				pnAccountID - The AccountID
' Return: True if sucessful, False if not
' **************************************************************************
Function SetOrderPaymentType(ByVal pnOrderID, ByVal pnPaymentTypeID, ByVal psPaymentReference, ByVal pnAccountID)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	lsSQL = "update tblOrders set PaymentTypeID = " & pnPaymentTypeID
	If pnPaymentTypeID <= 2 Then
		lsSQL = lsSQL & ", IsPaid = 0"
	End If
	If Len(psPaymentReference) = 0 Then
		lsSQL = lsSQL & ", PaymentReference = NULL, PaymentAuthorization = NULL"
	Else
		lsSQL = lsSQL & ", PaymentReference = NULL, PaymentAuthorization = '" & psPaymentReference & "'"
	End If
	If pnAccountID = 0 Then
		lsSQL = lsSQL & ", AccountID = NULL"
	Else
		lsSQL = lsSQL & ", AccountID = " & pnAccountID
	End If
	lsSQL = lsSQL & " where OrderID = " & pnOrderID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If
	
	SetOrderPaymentType = lbRet
End Function

' **************************************************************************
' Function: SetOrderCompleted
' Purpose: Sets an order's status to completed.
' Parameters:	pnOrderID - The OrderID
' Return: True if sucessful, False if not
' **************************************************************************
Function SetOrderCompleted(ByVal pnOrderID)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	lsSQL = "update tblOrders set OrderStatusID = 10 where OrderID = " & pnOrderID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If
	
	SetOrderCompleted = lbRet
End Function

' **************************************************************************
' Function: GetStoreOrderTypes
' Purpose: Retrieves order types for a store.
' Parameters:	pnStoreID - The StoreID to search for
'				panOrderTypeIDs - Array of OrderTypeIDs
'				pasOrderTypeDescriptions - Array of order type descriptions
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStoreOrderTypes(ByVal pnStoreID, ByRef panOrderTypeIDs, ByRef pasOrderTypeDescriptions)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select tlkpOrderTypes.OrderTypeID, OrderTypeDescription from trelStoreOrderTypes inner join tlkpOrderTypes on trelStoreOrderTypes.OrderTypeID = tlkpOrderTypes.OrderTypeID and trelStoreOrderTypes.StoreID = " & pnStoreID & " order by tlkpOrderTypes.OrderTypeID"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve panOrderTypeIDs(lnPos), pasOrderTypeDescriptions(lnPos)
				
				panOrderTypeIDs(lnPos) = loRS("OrderTypeID")
				pasOrderTypeDescriptions(lnPos) = Trim(loRS("OrderTypeDescription"))
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panOrderTypeIDs(0), pasOrderTypeDescriptions(0)
			panOrderTypeIDs(0) = 0
			pasOrderTypeDescriptions(0) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStoreOrderTypes = lbRet
End Function

' **************************************************************************
' Function: IsOrderTypeTaxable
' Purpose: Determines if an order type is taxable.
' Parameters:	pnStoreID - The StoreID to search for
'				pnOrderTypeID - The OrderTypeID to search for
' Return: True or false
' **************************************************************************
Function IsOrderTypeTaxable(ByVal pnStoreID, ByVal pnOrderTypeID)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	lsSQL = "select IsTaxable from trelStoreOrderTypes where StoreID = " & pnStoreID & " and OrderTypeID = " & pnOrderTypeID & " and IsTaxable <> 0"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lbRet = TRUE
		End If
		
		DBCloseQuery loRS
	End If
	
	IsOrderTypeTaxable = lbRet
End Function

' **************************************************************************
' Function: SubmitHoldOrder
' Purpose: Submits a hold order.
' Parameters:	pnOrderID - The OrderID
'				psExpectedDateTime - The expected date and time of the hold order
'				pnReleaseMinutes - The number of minutes ahead to release the order
'				pnCustomerID - The CustomerID
'				pnAddressID = The AddressID
' Return: True if sucessful, False if not
' **************************************************************************
Function SubmitHoldOrder(ByVal pnOrderID, ByVal psExpectedDateTime, ByVal pnReleaseMinutes, ByVal pnCustomerID, ByVal pnAddressID)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	' TODO: Take into account future day hold orders and adjust transaction date
	
	lsSQL = "update tblOrders set SubmitDate = getdate(), ReleaseDate = DATEADD(mi, -" & pnReleaseMinutes & ", '" & DBCleanLiteral(psExpectedDateTime) & "'), ExpectedDate = '" & DBCleanLiteral(psExpectedDateTime) & "', OrderStatusID = 2 where OrderID = " & pnOrderID
	If DBExecuteSQL(lsSQL) Then
		If pnCustomerID = 1 Or pnAddressID = 1 Then
			lbRet = TRUE
		Else
			lsSQL = "update tblCustomers set PrimaryAddressID = " & pnAddressID & " where CustomerID = " & pnCustomerID
			If DBExecuteSQL(lsSQL) Then
				lbRet = TRUE
			End If
		End If
	End If
	
	SubmitHoldOrder = lbRet
End Function

' **************************************************************************
' Function: ResetHoldOrder
' Purpose: Resets a hold order.
' Parameters:	pnOrderID - The OrderID
' Return: True if sucessful, False if not
' **************************************************************************
Function ResetHoldOrder(ByVal pnOrderID)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	' TODO: Take into account future day hold orders and adjust transaction date
	
	lsSQL = "update tblOrders set OrderStatusID = 2 where OrderID = " & pnOrderID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If
	
	ResetHoldOrder = lbRet
End Function

' **************************************************************************
' Function: GetHoldOrderTimes
' Purpose: Retrieves upcoming hold orders for a store.
' Parameters:	pnStoreID - The StoreID to search for
'				pasHoldOrderTimes - Array of times
' Return: True if sucessful, False if not
' **************************************************************************
Function GetHoldOrderTimes(ByVal pnStoreID, ByRef pasHoldOrderTimes)
	Dim lbRet, lsSQL, loRS, lnPos, ldtTargetDate
	
	lbRet = FALSE
	
	ldtTargetDate = DateAdd("h", gnHoldOrderDisplayHours, Now())
	
	lsSQL = "select ReleaseDate from tblOrders where StoreID = " & pnStoreID & " and OrderStatusID = 2 and ReleaseDate <= '" & ldtTargetDate & "' order by ReleaseDate"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve pasHoldOrderTimes(lnPos)
				
				If Hour(loRS("ReleaseDate")) = 0 Then
					pasHoldOrderTimes(lnPos) = "12:" & Left(Minute(loRS("ReleaseDate")) & "0", 2) & " AM"
				Else
					If Hour(loRS("ReleaseDate")) = 12 Then
						pasHoldOrderTimes(lnPos) = "12:" & Left(Minute(loRS("ReleaseDate")) & "0", 2) & " PM"
					Else
						If Hour(loRS("ReleaseDate")) > 12 Then
							pasHoldOrderTimes(lnPos) = (Hour(loRS("ReleaseDate")) - 12) & ":" & Left(Minute(loRS("ReleaseDate")) & "0", 2) & " PM"
						Else
							pasHoldOrderTimes(lnPos) = Hour(loRS("ReleaseDate")) & ":" & Left(Minute(loRS("ReleaseDate")) & "0", 2) & " AM"
						End If
					End If
				End If
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim pasHoldOrderTimes(0)
			pasHoldOrderTimes(0) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetHoldOrderTimes = lbRet
End Function

' **************************************************************************
' Function: ReleaseHoldOrders
' Purpose: Releases hold orders and return the OrderIDs that need to be printed.
' Parameters:	pnStoreID - The StoreID
'				pdtTransactionDate - The current transaction date
'				panOrderIDs - Array of OrderIDs
' Return: True if sucessful, False if not
' **************************************************************************
Function ReleaseHoldOrders(ByVal pnStoreID, ByVal pdtTransactionDate, ByRef panOrderIDs)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "EXEC ReleaseHoldOrders @pStoreID = " & pnStoreID & ", @pTransactionDate = '" & pdtTransactionDate & "'"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve panOrderIDs(lnPos)
				
				panOrderIDs(lnPos) = loRS(0)
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panOrderIDs(0)
			panOrderIDs(0) = 0
		End If
		
		DBCloseQuery loRS
	End If
	
	ReleaseHoldOrders = lbRet
End Function

' **************************************************************************
' Function: DuplicateOrderLine
' Purpose: Creates a new order line from an existing order line.
' Parameters:	pnOrderID - The OrderID
'				pnOrderLineID - The OrderLineID to duplicate
'				pnNewOrderLineID - The new order line ID
' Return: True if sucessful, False if not
' **************************************************************************
Function DuplicateOrderLine(ByVal pnOrderID, ByVal pnOrderLineID, ByRef pnNewOrderLineID)
	Dim lbRet, lnTmp, i
	Dim lnUnitID, lnSpecialtyID, lnSizeID, lnStyleID, lnHalf1SauceID, lnHalf2SauceID, lnHalf1SauceModifierID, lnHalf2SauceModifierID, lsOrderLineNotes, lnQuantity, ldOrderLineCost, ldOrderLineDiscount, lnCouponID
	Dim lanOrderLineItemIDs(), lanItemIDs(), lanHalfIDs(), lasItemDescriptions(), lasItemShortDescriptions()
	Dim lanOrderLineTopperIDs(), lanTopperIDs(), lanTopperHalfIDs(), lasTopperDescriptions(), lasTopperShortDescriptions()
	Dim lanOrderLineFreeSideIDs(), lanFreeSideIDs(), lasFreeSideDescriptions(), lasFreeSideShortDescriptions()
	Dim lanOrderLineAddSideIDs(), lanAddSideIDs(), lasAddSideDescriptions(), lasAddSideShortDescriptions()
	
	lbRet = FALSE
	
	If GetOrderLineDetails(pnOrderLineID, lnUnitID, lnSpecialtyID, lnSizeID, lnStyleID, lnHalf1SauceID, lnHalf2SauceID, lnHalf1SauceModifierID, lnHalf2SauceModifierID, lsOrderLineNotes, lnQuantity, ldOrderLineCost, ldOrderLineDiscount, lnCouponID) Then
		If GetOrderLineItems(pnOrderLineID, lnUnitID, lanOrderLineItemIDs, lanItemIDs, lanHalfIDs, lasItemDescriptions, lasItemShortDescriptions) Then
			If GetOrderLineToppers(pnOrderLineID, lanOrderLineTopperIDs, lanTopperIDs, lanTopperHalfIDs, lasTopperDescriptions, lasTopperShortDescriptions) Then
				If GetOrderLineFreeSides(pnOrderLineID, lanOrderLineFreeSideIDs, lanFreeSideIDs, lasFreeSideDescriptions, lasFreeSideShortDescriptions) Then
					If GetOrderLineAddSides(pnOrderLineID, lanOrderLineAddSideIDs, lanAddSideIDs, lasAddSideDescriptions, lasAddSideShortDescriptions) Then
						pnNewOrderLineID = CreateOrderLine(pnOrderID, lnUnitID, lnSpecialtyID, lnSizeID, lnStyleID, lnHalf1SauceID, lnHalf2SauceID, lnHalf1SauceModifierID, lnHalf2SauceModifierID, lsOrderLineNotes, lnQuantity)
						If pnNewOrderLineID > 0 Then
							lbRet = TRUE
							
							If lanItemIDs(0) <> 0 Then
								For i = 0 To UBound(lanItemIDs)
									lnTmp = CreateOrderLineItem(pnNewOrderLineID, lanItemIDs(i), lanHalfIDs(i))
									If lnTmp = 0 Then
										lbRet = FALSE
										Exit For
									End If
								Next
							End If
							
							If lbRet Then
								If lanTopperIDs(0) <> 0 Then
									For i = 0 To UBound(lanTopperIDs)
										lnTmp = CreateOrderLineTopper(pnNewOrderLineID, lanTopperIDs(i), lanTopperHalfIDs(i))
										If lnTmp = 0 Then
											lbRet = FALSE
											Exit For
										End If
									Next
								End If
							End If
							
							If lbRet Then
								If lanFreeSideIDs(0) <> 0 Then
									For i = 0 To UBound(lanFreeSideIDs)
										lnTmp = CreateOrderLineSide(pnNewOrderLineID, lanFreeSideIDs(i), 1)
										If lnTmp = 0 Then
											lbRet = FALSE
											Exit For
										End If
									Next
								End If
							End If
							
							If lbRet Then
								If lanAddSideIDs(0) <> 0 Then
									For i = 0 To UBound(lanAddSideIDs)
										lnTmp = CreateOrderLineSide(pnNewOrderLineID, lanAddSideIDs(i), 0)
										If lnTmp = 0 Then
											lbRet = FALSE
											Exit For
										End If
									Next
								End If
							End If
						End If
					End If
				End If
			End If
		End If
	End If
	
	DuplicateOrderLine = lbRet
End Function

' **************************************************************************
' Function: DuplicateOrder
' Purpose: Creates a new order from an existing order.
' Parameters:	pnOrderID - The OrderID
'				pnNewOrderID - The new order ID
'				psCouponIDs - The coupon IDs applied
' Return: True if sucessful, False if not
' **************************************************************************
Function DuplicateOrder(ByVal pnOrderID, ByRef pnNewOrderID, ByRef psCouponIDs)
	Dim lbRet, lnTmp, h, i, lnNewOrderLineID
	Dim lnSessionID, lsIPAddress, lnEmpID, lsRefID, ldtTransactionDate, ldtSubmitDate, ldtReleaseDate, ldtExpectedDate, lnStoreID, lnCustomerID, lsCustomerName, lsCustomerPhone, lnAddressID, lnOrderTypeID, lbIsPaid, lnPaymentTypeID, lsPaymentReference, lnAccountID, ldDeliveryCharge, ldDriverMoney, ldTax, ldTax2, ldTip, lnOrderStatusID, lsOrderNotes
	Dim lanOrderLineIDs(), lasOrderLineDescriptions(), lanQuantity(), ladCost(), ladDiscount()
	Dim lnUnitID, lnSpecialtyID, lnSizeID, lnStyleID, lnHalf1SauceID, lnHalf2SauceID, lnHalf1SauceModifierID, lnHalf2SauceModifierID, lsOrderLineNotes, lnQuantity, ldOrderLineCost, ldOrderLineDiscount, lnCouponID
	Dim lanOrderLineItemIDs(), lanItemIDs(), lanHalfIDs(), lasItemDescriptions(), lasItemShortDescriptions()
	Dim lanOrderLineTopperIDs(), lanTopperIDs(), lanTopperHalfIDs(), lasTopperDescriptions(), lasTopperShortDescriptions()
	Dim lanOrderLineFreeSideIDs(), lanFreeSideIDs(), lasFreeSideDescriptions(), lasFreeSideShortDescriptions()
	Dim lanOrderLineAddSideIDs(), lanAddSideIDs(), lasAddSideDescriptions(), lasAddSideShortDescriptions()
	
	lbRet = FALSE
	
	If GetOrderDetails(pnOrderID, lnSessionID, lsIPAddress, lnEmpID, lsRefID, ldtTransactionDate, ldtSubmitDate, ldtReleaseDate, ldtExpectedDate, lnStoreID, lnCustomerID, lsCustomerName, lsCustomerPhone, lnAddressID, lnOrderTypeID, lbIsPaid, lnPaymentTypeID, lsPaymentReference, lnAccountID, ldDeliveryCharge, ldDriverMoney, ldTax, ldTax2, ldTip, lnOrderStatusID, lsOrderNotes) Then
		If GetOrderLines(pnOrderID, lanOrderLineIDs, lasOrderLineDescriptions, lanQuantity, ladCost, ladDiscount) Then
			pnNewOrderID = CreateOrder(Session.SessionID, Request.ServerVariables("REMOTE_ADDR"), Session("EmpID"), lsRefID, Session("TransactionDate"), Session("StoreID"), lnCustomerID, lsCustomerName, lsCustomerPhone, lnAddressID, lnOrderTypeID, ldDeliveryCharge, ldDriverMoney, lsOrderNotes)
			If pnNewOrderID > 0 Then
				lbRet = TRUE
				
				If lanOrderLineIDs(0) <> 0 Then
					For h = 0 To UBound(lanOrderLineIDs)
						If GetOrderLineDetails(lanOrderLineIDs(h), lnUnitID, lnSpecialtyID, lnSizeID, lnStyleID, lnHalf1SauceID, lnHalf2SauceID, lnHalf1SauceModifierID, lnHalf2SauceModifierID, lsOrderLineNotes, lnQuantity, ldOrderLineCost, ldOrderLineDiscount, lnCouponID) Then
							If lnCouponID <> 0 Then
								If Len(psCouponIDs) = 0 Then
									psCouponIDs = lnCouponID
								Else
									psCouponIDs = psCouponIDs & "," & lnCouponID
								End If
							End If
							
							If GetOrderLineItems(lanOrderLineIDs(h), lnUnitID, lanOrderLineItemIDs, lanItemIDs, lanHalfIDs, lasItemDescriptions, lasItemShortDescriptions) Then
								If GetOrderLineToppers(lanOrderLineIDs(h), lanOrderLineTopperIDs, lanTopperIDs, lanTopperHalfIDs, lasTopperDescriptions, lasTopperShortDescriptions) Then
									If GetOrderLineFreeSides(lanOrderLineIDs(h), lanOrderLineFreeSideIDs, lanFreeSideIDs, lasFreeSideDescriptions, lasFreeSideShortDescriptions) Then
										If GetOrderLineAddSides(lanOrderLineIDs(h), lanOrderLineAddSideIDs, lanAddSideIDs, lasAddSideDescriptions, lasAddSideShortDescriptions) Then
											lnNewOrderLineID = CreateOrderLine(pnNewOrderID, lnUnitID, lnSpecialtyID, lnSizeID, lnStyleID, lnHalf1SauceID, lnHalf2SauceID, lnHalf1SauceModifierID, lnHalf2SauceModifierID, lsOrderLineNotes, lnQuantity)
											If lnNewOrderLineID > 0 Then
												If lanItemIDs(0) <> 0 Then
													For i = 0 To UBound(lanItemIDs)
														lnTmp = CreateOrderLineItem(lnNewOrderLineID, lanItemIDs(i), lanHalfIDs(i))
														If lnTmp = 0 Then
															lbRet = FALSE
															Exit For
														End If
													Next
												End If
												
												If lbRet Then
													If lanTopperIDs(0) <> 0 Then
														For i = 0 To UBound(lanTopperIDs)
															lnTmp = CreateOrderLineTopper(lnNewOrderLineID, lanTopperIDs(i), lanTopperHalfIDs(i))
															If lnTmp = 0 Then
																lbRet = FALSE
																Exit For
															End If
														Next
													End If
												End If
												
												If lbRet Then
													If lanFreeSideIDs(0) <> 0 Then
														For i = 0 To UBound(lanFreeSideIDs)
															lnTmp = CreateOrderLineSide(lnNewOrderLineID, lanFreeSideIDs(i), 1)
															If lnTmp = 0 Then
																lbRet = FALSE
																Exit For
															End If
														Next
													End If
												End If
												
												If lbRet Then
													If lanAddSideIDs(0) <> 0 Then
														For i = 0 To UBound(lanAddSideIDs)
															lnTmp = CreateOrderLineSide(lnNewOrderLineID, lanAddSideIDs(i), 0)
															If lnTmp = 0 Then
																lbRet = FALSE
																Exit For
															End If
														Next
													End If
												End If
												
												If Not lbRet Then
													Exit For
												End If
											Else
												lbRet = FALSE
												Exit For
											End If
										Else
											lbRet = FALSE
											Exit For
										End If
									Else
										lbRet = FALSE
										Exit For
									End If
								Else
									lbRet = FALSE
									Exit For
								End If
							Else
								lbRet = FALSE
								Exit For
							End If
						Else
							lbRet = FALSE
							Exit For
						End If
					Next
				End If
			End If
		End If
	End If
	
	DuplicateOrder = lbRet
End Function

' **************************************************************************
' Function: UpdateTransactionDate
' Purpose: Sets the transaction date of an order.
' Parameters:	pnOrderID - The OrderID to update
'				pdtTransactionDate - The transaction date
' Return: True if sucessful, False if not
' **************************************************************************
Function UpdateTransactionDate(ByVal pnOrderID, ByVal pdtTransactionDate)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	lsSQL = "update tblOrders set TransactionDate = '" & pdtTransactionDate & "' where OrderID = " & pnOrderID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If
	
	UpdateTransactionDate = lbRet
End Function

' **************************************************************************
' Function: GetVoidReasons
' Purpose: Retrieves a list of void reasons.
' Parameters:	pasVoidReasons - Array of void reasons
' Return: True if sucessful, False if not
' **************************************************************************
Function GetVoidReasons(ByRef pasVoidReasons)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select VoidReason from tlkpVoidReasons"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve pasVoidReasons(lnPos)
				
				pasVoidReasons(lnPos) = loRS("VoidReason")
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim pasVoidReasons(0)
			pasVoidReasons(0) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetVoidReasons = lbRet
End Function

' **************************************************************************
' Function: GetEditReasons
' Purpose: Retrieves a list of edit reasons.
' Parameters:	pasEditReasons - Array of edit reasons
' Return: True if sucessful, False if not
' **************************************************************************
Function GetEditReasons(ByRef pasEditReasons)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select EditReason from tlkpEditReasons"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve pasEditReasons(lnPos)
				
				pasEditReasons(lnPos) = loRS("EditReason")
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim pasEditReasons(0)
			pasEditReasons(0) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetEditReasons = lbRet
End Function

' **************************************************************************
' Function: GetMPOReasons
' Purpose: Retrieves a list of MPO reasons.
' Parameters:	pasMPOReasons - Array of MPO reasons
' Return: True if sucessful, False if not
' **************************************************************************
Function GetMPOReasons(ByRef pasMPOReasons)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select MPOReason from tlkpMPOReasons"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve pasMPOReasons(lnPos)
				
				pasMPOReasons(lnPos) = loRS("MPOReason")
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim pasMPOReasons(0)
			pasMPOReasons(0) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetMPOReasons = lbRet
End Function

' **************************************************************************
' Function: SetOrderEdited
' Purpose: Records that an order has been edited.
' Parameters:	pnOrderID - The OrderID
'				pnEmpID - The EmpID
'				psEditReason - The edit reason
' Return: True if sucessful, False if not
' **************************************************************************
Function SetOrderEdited(ByVal pnOrderID, ByVal pnEmpID, ByVal psEditReason)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	lsSQL = "update tblOrders set EditEmpID = " & pnEmpID
	lsSQL = lsSQL & ", EditReason = '" & DBCleanLiteral(psEditReason) & "'"
	lsSQL = lsSQL & " where OrderID = " & pnOrderID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If
	
	SetOrderEdited = lbRet
End Function
%>