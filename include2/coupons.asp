<%
' **************************************************************************
' File: coupon.asp
' Purpose: Functions for coupon related activities.
' Created: 9/22/2011 - TAM
' Description:
'	Include this file on any page where coupons are handled.
'	This file includes the following functions: GetActiveCoupons,
'		GetOrderCoupons, SetManagerPriceOverride, GetManagerPrice,
'		GetCouponDetails, GetCouponAppliesTo, GetCouponCombos, SetCouponDiscount,
'		ClearAllCoupons, RecalculateOrderDiscounts, GetActiveCouponsByUnitSizeSpecialtyStyle,
'		ZeroOutOrder, GetActiveCouponsUnitSize, GetPromoCouponID,
'		GetOrderFreeCoupons, GetCouponDescription, PromoCodeExists,
'		AddPromoCode
'
' Revision History:
' 7/19/201 - Created
' **************************************************************************

' **************************************************************************
' Function: GetActiveCoupons
' Purpose: Retrieves a list of active coupons.
' Parameters:	pnStoreID - The StoreID to search for
'				pnOrderTypeID - The order type to seach for
'				pdTransactionDate - The current transaction date
'				pbIsInternet - Is this for showing on the internet
'				panCouponIDs - Array of coupon IDs
'				pasCouponDescriptions - Array of coupon descriptions
'				pasCouponShortDescriptions - Array of coupon short descriptions
' Return: True if sucessful, False if not
' **************************************************************************
Function GetActiveCoupons(ByVal pnStoreID, ByVal pnOrderTypeID, ByVal pdTransactionDate, ByVal pbIsInternet, ByRef panCouponIDs, ByRef pasCouponDescriptions, ByRef pasCouponShortDescriptions)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select tblCoupons.CouponID, Description, ShortDescription from tblCoupons inner join trelCouponStore on tblCoupons.CouponID = trelCouponStore.CouponID and trelCouponStore.StoreID = " & pnStoreID
	lsSQL = lsSQL & " inner join tblCouponDateRange on tblCoupons.CouponID = tblCouponDateRange.CouponID where '" & pdTransactionDate & "' between ValidFrom and ValidTo and "
	Select Case pnOrderTypeID
		Case 1
			lsSQL = lsSQL & "ValidForDelivery <> 0"
		Case 2
			lsSQL = lsSQL & "ValidForPickup <> 0"
		Case 3
			lsSQL = lsSQL & "ValidForDineIn <> 0"
		Case 4
			lsSQL = lsSQL & "ValidForWalkInOrder <> 0"
	End Select
	If pbIsInternet Then
		lsSQL = lsSQL & " and ValidForInternetOrder <> 0 and ShowOnWeb <> 0"
	End If
	lsSQL = lsSQL & " order by IsFree DESC, Description"
	
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve panCouponIDs(lnPos), pasCouponDescriptions(lnPos), pasCouponShortDescriptions(lnPos)
				
				panCouponIDs(lnPos) = loRS("CouponID")
				pasCouponDescriptions(lnPos) = Trim(loRS("Description"))
				pasCouponShortDescriptions(lnPos) = Trim(loRS("ShortDescription"))
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panCouponIDs(0), pasCouponDescriptions(0), pasCouponShortDescriptions(0)
			
			panCouponIDs(0) = 0
			pasCouponDescriptions(0) = ""
			pasCouponShortDescriptions(0) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetActiveCoupons = lbRet
End Function

' **************************************************************************
' Function: GetOrderCoupons
' Purpose: Retrieves a list of all coupons applied to an order.
' Parameters:	pnOrderID - The OrderID
'				psCoupons - The list of coupons applied
' Return: True if sucessful, False if not
' **************************************************************************
Function GetOrderCoupons(ByVal pnOrderID, ByRef psCoupons)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	psCoupons = ""
	
	lsSQL = "select distinct Description, ShortDescription from tblOrderLines inner join tblCoupons on tblOrderLines.CouponID = tblCoupons.CouponID where OrderID = " & pnOrderID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			Do While Not loRS.eof
				If Len(psCoupons) = 0 Then
					psCoupons = loRS("Description")
				Else
					psCoupons = psCoupons & "," & loRS("Description")
				End If
				
				loRS.MoveNext
			Loop
		End If
		
		DBCloseQuery loRS
	End If
	
	GetOrderCoupons = lbRet
End Function

' **************************************************************************
' Function: SetManagerPriceOverride
' Purpose: Sets a manager price on a line item.
' Parameters:	pnOrderLineID - The OrderLineID
'				pdNewPrice - The new price.
'				psMPOReason - The MPO reason
' Return: True if sucessful, False if not
' **************************************************************************
Function SetManagerPriceOverride(ByVal pnOrderLineID, ByVal pdNewPrice, ByVal psMPOReason)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	lsSQL = "update tblOrderLines set Discount = (Cost - " & pdNewPrice & "), CouponID = 1, MPOReason = '" & psMPOReason & "' where OrderLineID = " & pnOrderLineID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If
	
	SetManagerPriceOverride = lbRet
End Function

' **************************************************************************
' Function: GetManagerPrice
' Purpose: Retrieves the manager price on a line item.
' Parameters:	pnOrderLineID - The OrderLineID
'				pdOriginalPrice - The original price
'				pdManagerPrice - The manager price (will be 0 if not specified).
'				psMPOReason - The reason for the MPO
' Return: True if sucessful, False if not
' **************************************************************************
Function GetManagerPrice(ByVal pnOrderLineID, ByRef pdOriginalPrice, ByRef pdManagerPrice, ByRef psMPOReason)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	lsSQL = "select Cost, Discount, CouponID, MPOReason from tblOrderLines where OrderLineID = " & pnOrderLineID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			pdOriginalPrice = loRS("Cost")
			
			If Not IsNull(loRS("CouponID")) Then
				If loRS("CouponID") = 1 Then
					pdManagerPrice = loRS("Cost") - loRS("Discount")
				Else
					pdManagerPrice = 0.00
				End If
			Else
				pdManagerPrice = 0.00
			End If
			If Not IsNull(loRS("MPOReason")) Then
				psMPOReason = loRS("MPOReason")
			Else
				psMPOReason = ""
			End If
		Else
			pdOriginalPrice = 0.00
			pdManagerPrice = 0.00
			psMPOReason = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetManagerPrice = lbRet
End Function

' **************************************************************************
' Function: GetCouponDetails
' Purpose: Retrieves coupon details.
' Parameters:	pnCouponID - The CouponID to search for
'				pnCouponTypeIDs - The CouponTypeID
'				pdCouponPercentageOff - The percentage off
'				pdCouponDollarOff - The dollar off
'				pdCouponMinimumPurchase - The minimum purchase
'				pbCouponIsFree - Is this coupon for a free unit
' Return: True if sucessful, False if not
' **************************************************************************
Function GetCouponDetails(ByVal pnCouponID, ByRef pnCouponTypeID, ByRef pdCouponPercentageOff, ByRef pdCouponDollarOff, ByRef pdCouponMinimumPurchase, ByRef pbCouponIsFree)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	lsSQL = "select CouponTypeID, PercentageOff, DollarOff, MinimumPurchase, IsFree from tblCoupons where CouponID = " & pnCouponID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lbRet = TRUE
			
			pnCouponTypeID = loRS("CouponTypeID")
			If Not IsNull(loRS("PercentageOff")) Then
				pdCouponPercentageOff = loRS("PercentageOff")
			Else
				pdCouponPercentageOff = 0.00
			End If
			If Not IsNull(loRS("PercentageOff")) Then
				pdCouponDollarOff = loRS("DollarOff")
			Else
				pdCouponDollarOff = 0.00
			End If
			If Not IsNull(loRS("MinimumPurchase")) Then
				pdCouponMinimumPurchase = loRS("MinimumPurchase")
			Else
				pdCouponMinimumPurchase = 0.00
			End If
			If loRS("IsFree") <> 0 Then
				pbCouponIsFree = TRUE
			Else
				pbCouponIsFree = FALSE
			End If
		End If
		
		DBCloseQuery loRS
	End If
	
	GetCouponDetails = lbRet
End Function

' **************************************************************************
' Function: GetCouponAppliesTo
' Purpose: Gets the details of what a coupon applies to.
' Parameters:	pnCouponID - The CouponID to search for
'				panCouponAppliesToID - Array of CouponAppliesToID
'				panCouponAppliesUnitID - Array of UnitIDs
'				panCouponAppliesSizeID - Array of SizeIDs
'				panCouponAppliesSpecialtyID - Array of SpecialtyIDs
'				panCouponAppliesStyleID - Array of StyleIDs
'				padCouponAppliesFixedPrice - Array of Fixed Prices
'				padCouponAppliesDollarOff - Array of Dollars Off
'				padCouponAppliesMinimumPrice - Array of Minimum Prices
'				padCouponAppliesAddForSpecialty - Array of add for specialty
'				panCouponAppliesComboChoice - Array of number of choices
'				panCouponAppliesComboQuantity - Array of quantities
' Return: True if sucessful, False if not
' **************************************************************************
Function GetCouponAppliesTo(ByVal pnCouponID, panCouponAppliesToID, panCouponAppliesUnitID, panCouponAppliesSizeID, panCouponAppliesSpecialtyID, panCouponAppliesStyleID, padCouponAppliesFixedPrice, padCouponAppliesDollarOff, padCouponAppliesMinimumPrice, padCouponAppliesAddForSpecialty, panCouponAppliesComboChoice, panCouponAppliesComboQuantity)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select CouponAppliesToID, UnitID, SizeID, SpecialtyID, StyleID, FixedPrice, DollarOff, MinimumPrice, AddForSpecialty, ComboChoice, ComboQuantity from tblCouponAppliesTo where CouponID = " & pnCouponID
	
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve panCouponAppliesToID(lnPos), panCouponAppliesUnitID(lnPos), panCouponAppliesSizeID(lnPos), panCouponAppliesSpecialtyID(lnPos), panCouponAppliesStyleID(lnPos), padCouponAppliesFixedPrice(lnPos), padCouponAppliesDollarOff(lnPos), padCouponAppliesMinimumPrice(lnPos), padCouponAppliesAddForSpecialty(lnPos), panCouponAppliesComboChoice(lnPos), panCouponAppliesComboQuantity(lnPos)
				
				panCouponAppliesToID(lnPos) = loRS("CouponAppliesToID")
				If IsNull(loRS("UnitID")) Then
					panCouponAppliesUnitID(lnPos) = 0
				Else
					panCouponAppliesUnitID(lnPos) = loRS("UnitID")
				End If
				If IsNull(loRS("SizeID")) Then
					panCouponAppliesSizeID(lnPos) = 0
				Else
					panCouponAppliesSizeID(lnPos) = loRS("SizeID")
				End If
				If IsNull(loRS("SpecialtyID")) Then
					panCouponAppliesSpecialtyID(lnPos) = 0
				Else
					panCouponAppliesSpecialtyID(lnPos) = loRS("SpecialtyID")
				End If
				If IsNull(loRS("StyleID")) Then
					panCouponAppliesStyleID(lnPos) = 0
				Else
					panCouponAppliesStyleID(lnPos) = loRS("StyleID")
				End If
				If IsNull(loRS("FixedPrice")) Then
					padCouponAppliesFixedPrice(lnPos) = 0.00
				Else
					padCouponAppliesFixedPrice(lnPos) = loRS("FixedPrice")
				End If
				If IsNull(loRS("DollarOff")) Then
					padCouponAppliesDollarOff(lnPos) = 0.00
				Else
					padCouponAppliesDollarOff(lnPos) = loRS("DollarOff")
				End If
				padCouponAppliesMinimumPrice(lnPos) = loRS("MinimumPrice")
				padCouponAppliesAddForSpecialty(lnPos) = loRS("AddForSpecialty")
				If IsNull(loRS("ComboChoice")) Then
					panCouponAppliesComboChoice(lnPos) = 0
				Else
					panCouponAppliesComboChoice(lnPos) = loRS("ComboChoice")
				End If
				If IsNull(loRS("ComboQuantity")) Then
					panCouponAppliesComboQuantity(lnPos) = 0
				Else
					panCouponAppliesComboQuantity(lnPos) = loRS("ComboQuantity")
				End If
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panCouponAppliesToID(0), panCouponAppliesUnitID(0), panCouponAppliesSizeID(0), panCouponAppliesSpecialtyID(0), panCouponAppliesStyleID(0), padCouponAppliesFixedPrice(0), padCouponAppliesDollarOff(0), padCouponAppliesMinimumPrice(0), padCouponAppliesAddForSpecialty(0), panCouponAppliesComboChoice(0), panCouponAppliesComboQuantity(0)
			
			panCouponAppliesToID(0) = 0
			panCouponAppliesUnitID(0) = 0
			panCouponAppliesSizeID(0) = 0
			panCouponAppliesSpecialtyID(0) = 0
			panCouponAppliesStyleID(0) = 0
			padCouponAppliesFixedPrice(0) = 0
			padCouponAppliesDollarOff(0) = 0
			padCouponAppliesMinimumPrice(0) = 0
			padCouponAppliesAddForSpecialty(0) = 0
			panCouponAppliesComboChoice(0) = 0
			panCouponAppliesComboQuantity(0) = 0
		End If
		
		DBCloseQuery loRS
	End If
	
	GetCouponAppliesTo = lbRet
End Function


' **************************************************************************
' Function: GetCouponCombos
' Purpose: Gets the details of what a coupon combos are.
' Parameters:	pnCouponID - The CouponID to search for
'				panCouponCombosID - Array of CouponCombosID
'				panCouponCombosUnitID - Array of UnitIDs
'				panCouponCombosSizeID - Array of SizeIDs
'				panCouponCombosSpecialtyID - Array of SpecialtyIDs
'				panCouponCombosStyleID - Array of StyleIDs
'				padCouponCombosFixedPrice - Array of Fixed Prices
'				padCouponCombosDollarOff - Array of Dollars Off
'				padCouponCombosMinimumPrice - Array of Minimum Prices
'				padCouponCombosAddForSpecialty - Array of add for specialty
' Return: True if sucessful, False if not
' **************************************************************************
Function GetCouponCombos(ByVal pnCouponID, panCouponCombosID, panCouponCombosUnitID, panCouponCombosSizeID, panCouponCombosSpecialtyID, panCouponCombosStyleID, padCouponCombosFixedPrice, padCouponCombosDollarOff, padCouponCombosMinimumPrice, padCouponCombosAddForSpecialty)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select CouponCombosID, UnitID, SizeID, SpecialtyID, StyleID, FixedPrice, DollarOff, MinimumPrice, AddForSpecialty from tblCouponCombos where CouponID = " & pnCouponID
	
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve panCouponCombosID(lnPos), panCouponCombosUnitID(lnPos), panCouponCombosSizeID(lnPos), panCouponCombosSpecialtyID(lnPos), panCouponCombosStyleID(lnPos), padCouponCombosFixedPrice(lnPos), padCouponCombosDollarOff(lnPos), padCouponCombosMinimumPrice(lnPos), padCouponCombosAddForSpecialty(lnPos)
				
				panCouponCombosID(lnPos) = loRS("CouponCombosID")
				If IsNull(loRS("UnitID")) Then
					panCouponCombosUnitID(lnPos) = 0
				Else
					panCouponCombosUnitID(lnPos) = loRS("UnitID")
				End If
				If IsNull(loRS("SizeID")) Then
					panCouponCombosSizeID(lnPos) = 0
				Else
					panCouponCombosSizeID(lnPos) = loRS("SizeID")
				End If
				If IsNull(loRS("SpecialtyID")) Then
					panCouponCombosSpecialtyID(lnPos) = 0
				Else
					panCouponCombosSpecialtyID(lnPos) = loRS("SpecialtyID")
				End If
				If IsNull(loRS("StyleID")) Then
					panCouponCombosStyleID(lnPos) = 0
				Else
					panCouponCombosStyleID(lnPos) = loRS("StyleID")
				End If
				If IsNull(loRS("FixedPrice")) Then
					padCouponCombosFixedPrice(lnPos) = 0.00
				Else
					padCouponCombosFixedPrice(lnPos) = loRS("FixedPrice")
				End If
				If IsNull(loRS("DollarOff")) Then
					padCouponCombosDollarOff(lnPos) = 0.00
				Else
					padCouponCombosDollarOff(lnPos) = loRS("DollarOff")
				End If
				padCouponCombosMinimumPrice(lnPos) = loRS("MinimumPrice")
				padCouponCombosAddForSpecialty(lnPos) = loRS("AddForSpecialty")
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panCouponCombosID(0), panCouponCombosUnitID(0), panCouponCombosSizeID(0), panCouponCombosSpecialtyID(0), panCouponCombosStyleID(0), padCouponCombosFixedPrice(0), padCouponCombosDollarOff(0), padCouponCombosMinimumPrice(0), padCouponCombosAddForSpecialty(0)
			
			panCouponCombosID(0) = 0
			panCouponCombosUnitID(0) = 0
			panCouponCombosSizeID(0) = 0
			panCouponCombosSpecialtyID(0) = 0
			panCouponCombosStyleID(0) = 0
			padCouponCombosFixedPrice(0) = 0.00
			padCouponCombosDollarOff(lnPos) = 0.00
			padCouponCombosMinimumPrice(lnPos) = 0.00
			padCouponCombosAddForSpecialty(lnPos) = 0.00
		End If
		
		DBCloseQuery loRS
	End If
	
	GetCouponCombos = lbRet
End Function

' **************************************************************************
' Function: SetCouponDiscount
' Purpose: Sets a coupon discount.
' Parameters:	pnOrderLineID - The OrderLineID
'				pnCouponID - The CouponID
'				pdDiscount - The discount
' Return: True if sucessful, False if not
' **************************************************************************
Function SetCouponDiscount(ByVal pnOrderLineID, ByVal pnCouponID, ByVal pdDiscount)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	lsSQL = "update tblOrderLines set Discount = " & pdDiscount & ", CouponID = " & pnCouponID & " where OrderLineID = " & pnOrderLineID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If
	
	SetCouponDiscount = lbRet
End Function

' **************************************************************************
' Function: ClearAllCoupons
' Purpose: Removes all coupon discounts from an order.
' Parameters:	pnOrderID - The OrderID
' Return: True if sucessful, False if not
' **************************************************************************
Function ClearAllCoupons(ByVal pnOrderID)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	lsSQL = "update tblOrderLines set Discount = 0, CouponID = NULL, MPOReason = NULL where OrderID = " & pnOrderID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If
	
	ClearAllCoupons = lbRet
End Function

' **************************************************************************
' Function: ClearAllNMOCoupons
' Purpose: Removes all coupon discounts (except manager price overrides) from an order.
' Parameters:	pnOrderID - The OrderID
' Return: True if sucessful, False if not
' **************************************************************************
Function ClearAllNMOCoupons(ByVal pnOrderID)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	lsSQL = "update tblOrderLines set Discount = 0, CouponID = NULL where OrderID = " & pnOrderID & " and CouponID <> 1"
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If
	
	ClearAllNMOCoupons = lbRet
End Function

' **************************************************************************
' Function: RecalculateOrderDiscounts
' Purpose: Recalculates discounts for all lines of an order.
' Parameters:	pnOrderID - The OrderID
'				psCouponIDs - Comma separated list of coupon IDs to apply
' Return: True if coupons were applied, False if not
' **************************************************************************
Function RecalculateOrderDiscounts(ByVal pnStoreID, ByVal pnOrderID, ByVal psCouponIDs)
	Dim lbRet, lasCouponIDs, lnPos, i, j, k, l, m
	Dim lnCouponTypeID, ldCouponPercentageOff, ldCouponDollarOff, ldCouponMinimumPurchase, lbCouponIsFree
	Dim lanCouponIDs(), lanCouponTypeIDs(), ladCouponPercentageOff(), ladCouponDollarOff(), ladCouponMinimumPurchase(), ladCouponIsFree()
	Dim lanCouponAppliesToID(), lanCouponAppliesUnitID(), lanCouponAppliesSizeID(), lanCouponAppliesSpecialtyID(), lanCouponAppliesStyleID(), ladCouponAppliesFixedPrice(), ladCouponAppliesDollarOff(), ladCouponAppliesMinimumPrice(), ladCouponAppliesAddForSpecialty(), lanCouponAppliesComboChoice(), lanCouponAppliesComboQuantity()
	Dim lnSessionID, lsIPAddress, lnEmpID, lsRefID, ldtTransactionDate, ldtSubmitDate, ldtReleaseDate, ldtExpectedDate, lnStoreID, lnCustomerID, lsCustomerName, lsCustomerPhone, lnAddressID, lnOrderTypeID, lbIsPaid, lnPaymentTypeID, lsPaymentReference, lnAccountID, ldDeliveryCharge, ldDriverMoney, ldTax, ldTax2, ldTip, lnOrderStatusID, lsOrderNotes
	Dim lanOrderLineIDs(), lasOrderLineDescriptions(), lanQuantity(), ladCost(), ladDiscount(), lanOrderCouponIDs(), labOrderCouponIsCombo()
	Dim lnOrderLineID, lnUnitID, lnSpecialtyID, lnSizeID, lnStyleID, lnHalf1SauceID, lnHalf2SauceID, lnHalf1SauceModifierID, lnHalf2SauceModifierID, lsOrderLineNotes, lnQuantity, ldOrderLineCost, ldOrderLineDiscount, lnCouponID
	Dim lanOrderLineItemIDs(), lanItemIDs(), lanHalfIDs(), lasItemDescriptions(), lasItemShortDescriptions()
	Dim lanUnitItemIDs(), lasUnitItemDescriptions(), lasUnitItemShortDescriptions(), ladUnitItemOnSidePrice(), lanUnitItemCounts(), labUnitFreeItemFlags(), labUnitItemIsCheeses(), labUnitItemIsBaseCheeses(), labUnitItemIsExtraCheeses()
	Dim lnAppliesToIndex, ldTmpDiscount, lnItemCount, ldTmpPrice
	Dim lnFreeApplyToIndex, lnFreeApplyToDiscount
	Dim lnBOGOFoundIndex, lanBOGOAlreadyUsed()
	Dim lbBOGOIsAlsoBuy, lnBOGOFoundSkip
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	RecalculateOrderDiscounts = lbRet
End Function

' **************************************************************************
' Function: GetActiveCouponsByUnitSizeSpecialtyStyle
' Purpose: Retrieves a list of active coupons for a set of units and sizes.
' Parameters:	pnStoreID - The StoreID to search for
'				pnOrderTypeID - The order type to seach for
'				pdTransactionDate - The current transaction date
'				pbIsInternet - Is this for showing on the internet
'				panUnitIDs - Array of unit IDs to look for
'				panSizeIDs - Array of size IDs to look for
'				panSpecialtyIDs - Array of specialty IDs to look for
'				panStyleIDs - Array of style IDs to look for
'				panCouponIDs - Array of coupon IDs
'				pasCouponDescriptions - Array of coupon descriptions
'				pasCouponShortDescriptions - Array of coupon short descriptions
' Return: True if sucessful, False if not
' **************************************************************************
Function GetActiveCouponsByUnitSizeSpecialtyStyle(ByVal pnStoreID, ByVal pnOrderTypeID, ByVal pdTransactionDate, ByVal pbIsInternet, ByVal panUnitIDs, ByVal panSizeIDs, ByVal panStyleIDs, ByVal panSpecialtyIDs, ByRef panCouponIDs, ByRef pasCouponDescriptions, ByRef pasCouponShortDescriptions)
	Dim lbRet, lsSQL, loRS, lnPos, i, loRS2, lbAdd
	
	lbRet = FALSE
	
	lsSQL = "select tblCoupons.CouponID, Description, ShortDescription from tblCoupons inner join trelCouponStore on tblCoupons.CouponID = trelCouponStore.CouponID and trelCouponStore.StoreID = " & pnStoreID
	lsSQL = lsSQL & " inner join tblCouponDateRange on tblCoupons.CouponID = tblCouponDateRange.CouponID where '" & pdTransactionDate & "' between ValidFrom and ValidTo and "
	Select Case pnOrderTypeID
		Case 1
			lsSQL = lsSQL & "ValidForDelivery <> 0"
		Case 2
			lsSQL = lsSQL & "ValidForPickup <> 0"
		Case 3
			lsSQL = lsSQL & "ValidForDineIn <> 0"
		Case 4
			lsSQL = lsSQL & "ValidForWalkInOrder <> 0"
	End Select
	If pbIsInternet Then
		lsSQL = lsSQL & " and ValidForInternetOrder <> 0 and ShowOnWeb <> 0"
	End If
	lsSQL = lsSQL & " order by IsFree DESC, Description"
	
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			ReDim panCouponIDs(0), pasCouponDescriptions(0), pasCouponShortDescriptions(0)
			
			panCouponIDs(0) = 0
			pasCouponDescriptions(0) = ""
			pasCouponShortDescriptions(0) = ""
			
			lnPos = 0
			
			Do While Not loRS.eof
				lbAdd = FALSE
				
				If panUnitIDs(0) = 0 Then
					lbAdd = TRUE
				Else
					For i = 0 To UBound(panUnitIDs)
						lsSQL = "select UnitID, SizeID, StyleID, SpecialtyID from tblCouponAppliesTo where CouponID = " & loRS("CouponID")
						
						If DBOpenQuery(lsSQL, FALSE, loRS2) Then
							If Not loRS2.bof And Not loRS2.eof Then
								Do While Not loRS2.eof
									lbAdd = TRUE
									
									If Not IsNull(loRS2("UnitID")) Then
										If panUnitIDs(i) <> loRS2("UnitID") Then
											lbAdd = FALSE
										End If
									End If
									If panSizeIDs(i) <> 0 Then
										If Not IsNull(loRS2("SizeID")) Then
											If panSizeIDs(i) <> loRS2("SizeID") Then
												lbAdd = FALSE
											End If
										End If
									End If
									If panStyleIDs(i) <> 0 Then
										If Not IsNull(loRS2("StyleID")) Then
											If panStyleIDs(i) <> loRS2("StyleID") Then
												lbAdd = FALSE
											End If
										End If
									End If
									If panSpecialtyIDs(i) <> 0 Then
										If Not IsNull(loRS2("SpecialtyID")) Then
											If panSpecialtyIDs(i) <> loRS2("SpecialtyID") Then
												lbAdd = FALSE
											End If
										End If
									End If
									
									If lbAdd Then
										Exit Do
									Else
										loRS2.MoveNext
									End If
								Loop
							Else
								lbAdd = TRUE
							End If
							
							DBCloseQuery loRS2
						End If
						
						If lbAdd Then
							Exit For
						End If
					Next
				End If
				
				If lbAdd Then
					ReDim Preserve panCouponIDs(lnPos), pasCouponDescriptions(lnPos), pasCouponShortDescriptions(lnPos)
					
					panCouponIDs(lnPos) = loRS("CouponID")
					pasCouponDescriptions(lnPos) = Trim(loRS("Description"))
					pasCouponShortDescriptions(lnPos) = Trim(loRS("ShortDescription"))
					
					lnPos = lnPos + 1
				End If
				
				loRS.MoveNext
			Loop
		Else
			ReDim panCouponIDs(0), pasCouponDescriptions(0), pasCouponShortDescriptions(0)
			
			panCouponIDs(0) = 0
			pasCouponDescriptions(0) = ""
			pasCouponShortDescriptions(0) = ""
		End If
		
		DBCloseQuery loRS
	Else
		ReDim panCouponIDs(0), pasCouponDescriptions(0), pasCouponShortDescriptions(0)
		
		panCouponIDs(0) = 0
		pasCouponDescriptions(0) = ""
		pasCouponShortDescriptions(0) = ""
	End If
	
	GetActiveCouponsByUnitSizeSpecialtyStyle = lbRet
End Function

' **************************************************************************
' Function: ZeroOutOrder
' Purpose: Sets an order to 0 cost including delivery and tax.
' Parameters:	pnOrderID - The OrderLineID
'				psMPOReason - The reason the order is being zeroed
' Return: True if sucessful, False if not
' **************************************************************************
Function ZeroOutOrder(ByVal pnOrderID, ByVal psMPOReason)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	lsSQL = "update tblOrderLines set Discount = Cost, CouponID = 1, MPOReason = '" & psMPOReason & "' where OrderID = " & pnOrderID
	If DBExecuteSQL(lsSQL) Then
		lsSQL = "update tblOrders set DeliveryCharge = 0, Tax = 0, Tax2 = 0 where OrderID = " & pnOrderID
		If DBExecuteSQL(lsSQL) Then
			lbRet = TRUE
		End If
	End If
	
	ZeroOutOrder = lbRet
End Function

' **************************************************************************
' Function: GetActiveCouponsUnitSize
' Purpose: Retrieves a list of active coupons and their respective unit/sizes.
' Parameters:	pnStoreID - The StoreID to search for
'				pnOrderTypeID - The order type to seach for
'				pdTransactionDate - The current transaction date
'				pbIsInternet - Is this for showing on the internet
'				panCouponIDs - Array of coupon IDs
'				pasCouponDescriptions - Array of coupon descriptions
'				panUnitIDs - Array of UnitIDs
'				pasUnitDescriptions - Array of unit descriptions
'				panSizeIDs - Array of SizeIDs
'				pasSizeDescriptions - Array of size descriptions
' Return: True if sucessful, False if not
' **************************************************************************
Function GetActiveCouponsUnitSize(ByVal pnStoreID, ByVal pnOrderTypeID, ByVal pdTransactionDate, ByVal pbIsInternet, ByRef panCouponIDs, ByRef pasCouponDescriptions, ByRef panUnitIDs, ByRef pasUnitDescriptions, ByRef panSizeIDs, ByRef pasSizeDescriptions)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select tblCoupons.CouponID, Description, ShortDescription, tblCouponAppliesTo.UnitID, tblCouponAppliesTo.SizeID, UnitDescription, SizeDescription from tblCoupons inner join trelCouponStore on tblCoupons.CouponID = trelCouponStore.CouponID and trelCouponStore.StoreID = " & pnStoreID
	lsSQL = lsSQL & " inner join tblCouponDateRange on tblCoupons.CouponID = tblCouponDateRange.CouponID "
	lsSQL = lsSQL & "inner join tblCouponAppliesTo on tblCoupons.CouponID = tblCouponAppliesTo.CouponID "
	lsSQL = lsSQL & "inner join tblUnit on tblCouponAppliesTo.UnitID = tblUnit.UnitID "
	lsSQL = lsSQL & "inner join tblSizes on tblCouponAppliesTo.SizeID = tblSizes.SizeID "
	lsSQL = lsSQL & "where '" & pdTransactionDate & "' between ValidFrom and ValidTo "
	Select Case pnOrderTypeID
		Case 1
			lsSQL = lsSQL & "and ValidForDelivery <> 0 "
		Case 2
			lsSQL = lsSQL & "and ValidForPickup <> 0 "
		Case 3
			lsSQL = lsSQL & "and ValidForDineIn <> 0 "
		Case 4
			lsSQL = lsSQL & "and ValidForWalkInOrder <> 0 "
	End Select
	If pbIsInternet Then
		lsSQL = lsSQL & "and ValidForInternetOrder <> 0 and ShowOnWeb <> 0 "
	End If
	lsSQL = lsSQL & "order by tblCouponAppliesTo.UnitID, tblCouponAppliesTo.SizeID"
	
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve panCouponIDs(lnPos), pasCouponDescriptions(lnPos), panUnitIDs(lnPos), pasUnitDescriptions(lnPos), panSizeIDs(lnPos), pasSizeDescriptions(lnPos)
				
				panCouponIDs(lnPos) = loRS("CouponID")
				pasCouponDescriptions(lnPos) = Trim(loRS("Description"))
				panUnitIDs(lnPos) = loRS("UnitID")
				pasUnitDescriptions(lnPos) = Trim(loRS("UnitDescription"))
				panSizeIDs(lnPos) = loRS("SizeID")
				pasSizeDescriptions(lnPos) = Trim(loRS("SizeDescription"))
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panCouponIDs(0), pasCouponDescriptions(0), panUnitIDs(0), pasUnitDescriptions(0), panSizeIDs(0), pasSizeDescriptions(0)
			
			panCouponIDs(0) = 0
			pasCouponDescriptions(0) = ""
			panUnitIDs(0) = 0
			pasUnitDescriptions(0) = ""
			panSizeIDs(0) = 0
			pasSizeDescriptions(0) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetActiveCouponsUnitSize = lbRet
End Function

' **************************************************************************
' Function: GetPromoCouponID
' Purpose: Retrieves the CouponID of a promo code.
' Parameters:	psPromoCode - The promo code to search for
' Return: The CouponID found or 0 if not found
' **************************************************************************
Function GetPromoCouponID(ByVal psPromoCode)
	Dim lnRet, lsSQL, loRS
	
	lnRet = 0
	
	lsSQL = "select CouponID from tblCouponsPromoCodes where PromoCode = '" & DBCleanLiteral(psPromoCode) & "'"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lnRet = loRS("CouponID")
		End If
		
		DBCloseQuery loRS
	End If
	
	GetPromoCouponID = lnRet
End Function

' **************************************************************************
' Function: GetOrderFreeCoupons
' Purpose: Retrieves a list of all free coupons applied to an order.
' Parameters:	pnOrderID - The OrderID
'				panFreeCouponIDs - Array of coupon IDs applied
'				pasFreeCouponDescriptions - Array coupon descriptions applied
'				panFreeCouponCount - Array of how many of these coupons applied
' Return: True if sucessful, False if not
' **************************************************************************
Function GetOrderFreeCoupons(ByVal pnOrderID, ByRef panFreeCouponIDs, ByRef pasFreeCouponDescriptions, ByRef panFreeCouponCount)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select tblOrderLines.CouponID, Description, COUNT(*) as CouponCount from tblOrderLines inner join tblCoupons on tblOrderLines.CouponID = tblCoupons.CouponID where OrderID = " & pnOrderID & " and IsFree <> 0 group by tblOrderLines.CouponID, Description"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve panFreeCouponIDs(lnPos), pasFreeCouponDescriptions(lnPos), panFreeCouponCount(lnPos)
				
				panFreeCouponIDs(lnPos) = loRS("CouponID")
				pasFreeCouponDescriptions(lnPos) = loRS("Description")
				panFreeCouponCount(lnPos) = loRS("CouponCount")
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panFreeCouponIDs(0), pasFreeCouponDescriptions(0), panFreeCouponCount(0)
			
			panFreeCouponIDs(0) = 0
			pasFreeCouponDescriptions(0) = ""
			panFreeCouponCount(0) = 0
		End If
		
		DBCloseQuery loRS
	End If
	
	GetOrderFreeCoupons = lbRet
End Function

' **************************************************************************
' Function: GetCouponDescription
' Purpose: Retrieves the description of a coupon.
' Parameters:	pnCouponID - The coupon ID to search for
'				psDescription - The coupon description
'				psShortDescription - The coupon short description
' Return: True if sucessful, False if not
' **************************************************************************
Function GetCouponDescription(ByVal pnCouponID, ByRef psDescription, ByRef psShortDescription)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	lsSQL = "select Description, ShortDescription from tblCoupons where CouponID = " & pnCouponID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lbRet = TRUE
			
			If IsNull(loRS("Description")) Then
				psDescription = ""
			Else
				psDescription = loRS("Description")
			End If
			If IsNull(loRS("ShortDescription")) Then
				psShortDescription = ""
			Else
				psShortDescription = loRS("ShortDescription")
			End If
		End If
		
		DBCloseQuery loRS
	End If
	
	GetCouponDescription = lbRet
End Function

' **************************************************************************
' Function: PromoCodeExists
' Purpose: Determines if a promo code exists.
' Parameters:	psPromoCode - The promo code to search for
' Return: True if found, False if not
' **************************************************************************
Function PromoCodeExists(ByVal psPromoCode)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	lsSQL = "select CouponID from tblCouponsPromoCodes where PromoCode = '" & DBCleanLiteral(psPromoCode) & "'"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lbRet = TRUE
		End If
		
		DBCloseQuery loRS
	End If
	
	PromoCodeExists = lbRet
End Function

' **************************************************************************
' Function: AddPromoCode
' Purpose: Adds a promo code.
' Parameters:	psPromoCode - The promo code to add
'				pnCouponID - The coupon ID
'				pnMaxUses - The max uses
'				psValidFrom - The valid from date
'				psValidTo - The valid to date
'				pbIsMassMailer - Flag for is code is for a mass mailer
' Return: True if sucessful, False if not
' **************************************************************************
Function AddPromoCode(ByVal psPromoCode, ByVal pnCouponID, ByVal pnMaxUses, ByVal psValidFrom, ByVal psValidTo, ByVal pbIsMassMailer)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	lsSQL = "insert into tblCouponsPromoCodes (PromoCode, CouponID, MaxUses"
	If pbIsMassMailer Then
		lsSQL = lsSQL & ", IsMassMailer"
	End If
	lsSQL = lsSQL & ") values ('" & DBCleanLiteral(psPromoCode) & "', " & pnCouponID & ", " & pnMaxUses
	If pbIsMassMailer Then
		lsSQL = lsSQL & ", 1"
	End If
	lsSQL = lsSQL & ")"
	
	If DBExecuteSQL(lsSQL) Then
		lsSQL = "insert into tblCouponPromoCodeDateRange (PromoCode, ValidFrom, ValidTo) values ('" & DBCleanLiteral(psPromoCode) & "', '" & DBCleanLiteral(psValidFrom) & "', '" & DBCleanLiteral(psValidTo) & "')"
		
		If DBExecuteSQL(lsSQL) Then
			lbRet = TRUE
		End If
	End If
	
	AddPromoCode = lbRet
End Function
%>