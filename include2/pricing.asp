<%
' **************************************************************************
' File: pricing.asp
' Purpose: Functions for price related activities.
' Created: 7/28/2011 - TAM
' Description:
'	Include this file on any page where price data is manipulated.
'	This file includes the following functions: RecalculateOrderTax,
'		RecalculateOrderPrice, RecalculateOrderLinePrice, 
'		RecalculateOrderLineIdealCost
'
' Revision History:
' 7/28/201 - Created
' **************************************************************************

' **************************************************************************
' Function: RecalculateOrderTax
' Purpose: Adds an amount to the tax on an order.
' Parameters:	pnStoreID - StoreID to search for
'				pnOrderID - The OrderID to search for
' Return: True if sucessful, False if not
' **************************************************************************
Function RecalculateOrderTax(ByVal pnStoreID, ByVal pnOrderID)
	Dim lbRet, lsSQL, loRS, loRS2, ldTaxRate, ldTaxRate2, lbIsDeliveryTaxable, ldTaxablePrice, ldTaxablePrice2, ldTax, ldTax2
	
	lbRet = FALSE
	
	ldTaxRate = GetStoreTaxRate(pnStoreID)
	ldTaxRate2 = GetStoreTaxRate2(pnStoreID)
	lbIsDeliveryTaxable = IsDeliveryTaxable(pnStoreID)
	
	If ldTaxRate <> -1 Then
		lsSQL = "select OrderTypeID, Tax, Tax2, DeliveryCharge from tblOrders where StoreID = " & pnStoreID & " and OrderID = " & pnOrderID
		If DBOpenQuery(lsSQL, TRUE, loRS) Then
			If Not loRS.bof And Not loRS.eof Then
				ldTax = 0.00
				ldTax2 = 0.00
				
				' Recalc tax here
				If IsOrderTypeTaxable(pnStoreID, loRS("OrderTypeID")) Then
					lsSQL = "select (Quantity * (Cost - Discount)) as Price from tblOrderLines where OrderID = " & pnOrderID
				Else
					lsSQL = "select (Quantity * (Cost - Discount)) as Price from tblOrderLines inner join trelStoreUnitSize on trelStoreUnitSize.StoreID = " & pnStoreID & " and tblOrderLines.UnitID = trelStoreUnitSize.UnitID and tblOrderLines.SizeID = trelStoreUnitSize.SizeID where OrderID = " & pnOrderID & " and IsTaxable <> 0"
				End If
				If DBOpenQuery(lsSQL, FALSE, loRS2) Then
					ldTaxablePrice = 0.00
					ldTaxablePrice2 = 0.00
					
					If Not loRS2.bof And Not loRS2.eof Then
						Do While Not loRS2.eof
							ldTaxablePrice = ldTaxablePrice + loRS2("Price")
							ldTaxablePrice2 = ldTaxablePrice2 + loRS2("Price")
							
							loRS2.MoveNext
						Loop
					End If
					
					DBCloseQuery loRS2
					
					If lbIsDeliveryTaxable Then
						ldTaxablePrice = ldTaxablePrice + loRS("DeliveryCharge")
					End If
					
					ldTax = Round2((ldTaxablePrice * ldTaxRate), 2)
					ldTax2 = Round2((ldTaxablePrice2 * ldTaxRate2), 2)
					If ldTax = loRS("Tax") And ldTax2 = loRS("Tax2") Then
						lbRet = TRUE
					Else
						On Error Resume Next
						If ldTax <> loRS("Tax") Then
							loRS("Tax") = ldTax
							If Err.Number = 0 Then
								If ldTax2 <> loRS("Tax2") Then
									loRS("Tax2") = ldTax2
									If Err.Number = 0 Then
										loRS.Update
										If Err.Number = 0 Then
											lbRet = TRUE
										Else
											gsDBErrorMessage = Err.Description
										End If
									Else
										gsDBErrorMessage = Err.Description
									End If
								Else
									loRS.Update
									If Err.Number = 0 Then
										lbRet = TRUE
									Else
										gsDBErrorMessage = Err.Description
									End If
								End If
							Else
								gsDBErrorMessage = Err.Description
							End If
						Else
							loRS("Tax2") = ldTax2
							If Err.Number = 0 Then
								loRS.Update
								If Err.Number = 0 Then
									lbRet = TRUE
								Else
									gsDBErrorMessage = Err.Description
								End If
							Else
								gsDBErrorMessage = Err.Description
							End If
						End If
						On Error Goto 0
					End If
				End If
			End If
			
			DBCloseQuery loRS
		End If
	End If
	
	RecalculateOrderTax = lbRet
End Function

' **************************************************************************
' Function: RecalculateOrderPrice
' Purpose: Recalculates the price for an entire order.
' Parameters:	pnStoreID - The StoreID to search for
' 				pnOrderID - The OrderID to search for
' Return: True if sucessful, False if not
' **************************************************************************
Function RecalculateOrderPrice(ByVal pnStoreID, ByVal pnOrderID)
	Dim lbRet, lsSQL, loRS, ldNewPrice
	
	lbRet = FALSE
	
	lsSQL = "select OrderLineID from tblOrderLines where OrderID = " & pnOrderID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lbRet = TRUE
			
			Do While Not loRS.eof
				If Not RecalculateOrderLinePrice(pnStoreID, loRS("OrderLineID"), ldNewPrice) Then
					lbRet = FALSE
				End If
				
				If lbRet Then
					loRS.MoveNext
				Else
					Exit Do
				End If
			Loop
		End If
		
		DBCloseQuery loRS
	End If
	
	RecalculateOrderPrice = lbRet
End Function

' **************************************************************************
' Function: RecalculateOrderLinePrice
' Purpose: Recalculates the price for one line on an order.
' Parameters:	pnStoreID - The StoreID to search for
' 				pnOrderLineID - The OrderLineID to search for
'				pdNewPrice - The new price for the order line
' Return: True if sucessful, False if not
' **************************************************************************
Function RecalculateOrderLinePrice(ByVal pnStoreID, ByVal pnOrderLineID, ByRef pdNewPrice)
	Dim lbRet, lsSQL, loRS, i, j, k, l, ldPrice, lnIncludedItems, lnTmp, lnItemVariance, ldPerItemPrices, lnItemCount, ldPremiumSurcharge, lbIgnoreSpecialty
	Dim lnUnitID, lnSpecialtyID, lnSizeID, lnStyleID, lnHalf1SauceID, lnHalf2SauceID, lnHalf1SauceModifierID, lnHalf2SauceModifierID
	Dim lanOrderLineItemIDs(), lanItemIDs(), lanHalfIDs(), lasItemDescriptions(), lasItemShortDescriptions()
	Dim lanOrderLineTopperIDs(), lanTopperIDs(), lanTopperHalfIDs(), lasTopperDescriptions(), lasTopperShortDescriptions()
	Dim lanOrderLineAddSideIDs(), lanAddSideIDs(), lasAddSideDescriptions(), lasAddSideShortDescriptions()
	Dim lanUnitSizeIDs(), lasUnitSizeDescriptions(), lasUnitSizeShortDescriptions(), ladUnitSizeStandardBasePrice(), lanUnitSizeStandardNumberIncludedItems(), ladUnitSizeSpecialtyBasePrice(), lanUnitSizeSpecialtyNumberIncludedItems(), lanUnitSizePercentSpecialtyItemVariances(), ladPerAdditionalItemPrices(), labIsTaxable()
	Dim lanSizeStyleIDs(), lasSizeStyleDescriptions(), lasSizeStyleShortDescriptions(), lasSizeStyleSpecialMessage(), lanSizeStyleSizeIDs(), ladSizeStyleSurcharges()
	Dim lanSauceIDs(), lasSauceDescriptions(), lasSauceShortDescriptions()
	Dim lanSauceModifierIDs(), lasSauceModifierDescriptions(), lasSauceModifierShortDescriptions()
	Dim lanUnitItemIDs(), lasUnitItemDescriptions(), lasUnitItemShortDescriptions(), ladUnitItemOnSidePrice(), lanUnitItemCounts(), labUnitFreeItemFlags(), labUnitItemIsCheeses(), labUnitItemIsBaseCheeses(), labUnitItemIsExtraCheeses()
	Dim lanUnitTopperIDs(), lasUnitTopperDescriptions(), lasUnitTopperShortDescriptions()
	Dim lanUnitSideIDs(), lasUnitSideDescriptions(), lasUnitSideShortDescriptions(), ladUnitSidePrices()
	Dim lanUnitSpecialtyIDs(), lasUnitSpecialtyDescriptions(), lasUnitSpecialtyShortDescriptions(), lanUnitSpecialtySauceID(), lanUnitSpecialtyStyleID(), labSpecialtyNoBaseCheese(), lanUnitSpecialtyItemIDs(), lanUnitSpecialtyItemQuantity()
	Dim lanUpchargeSizeIDs(), lanUpchargeItemIDs(), ladUpchargePrice()
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	RecalculateOrderLinePrice = lbRet
End Function

' **************************************************************************
' Function: RecalculateOrderLineIdealCost
' Purpose: Recalculates the ideal cost for one line on an order.
' Parameters:	pnStoreID - The StoreID to search for
' 				pnOrderLineID - The OrderLineID to search for
'				pnOrderTypeID - The OrderTypeID
' Return: True if sucessful, False if not
' **************************************************************************
Function RecalculateOrderLineIdealCost(ByVal pnStoreID, ByVal pnOrderLineID, ByVal pnOrderTypeID)
	Dim lbRet, lsSQL, loRS
	Dim lnUnitID, lnSpecialtyID, lnSizeID, lnStyleID, lnHalf1SauceID, lnHalf2SauceID, lnHalf1SauceModifierID, lnHalf2SauceModifierID, lsOrderLineNotes, lnQuantity, ldOrderLineCost, ldOrderLineDiscount, lnCouponID
	Dim lanOrderLineItemIDs(), lanItemIDs(), lanHalfIDs(), lasItemDescriptions(), lasItemShortDescriptions()
	Dim lanOrderLineTopperIDs(), lanTopperIDs(), lanTopperHalfIDs(), lasTopperDescriptions(), lasTopperShortDescriptions()
	Dim lanOrderLineFreeSideIDs(), lanFreeSideIDs(), lasFreeSideDescriptions(), lasFreeSideShortDescriptions()
	Dim lanOrderLineAddSideIDs(), lanAddSideIDs(), lasAddSideDescriptions(), lasAddSideShortDescriptions()
	Dim lanUnitSideIDs(), lasUnitSideDescriptions(), lasUnitSideShortDescriptions(), ladUnitSidePrices()
	Dim lanUnitItemIDs(), lasUnitItemDescriptions(), lasUnitItemShortDescriptions(), ladUnitItemOnSidePrice(), lanUnitItemCounts(), labUnitFreeItemFlags(), labUnitItemIsCheeses(), labUnitItemIsBaseCheeses(), labUnitItemIsExtraCheeses()
	Dim lanWeightItemIDs(), lanItemCounts1(), ladItemWeights1(), lanItemCounts2(), ladItemWeights2(), ladItemWeights3()
	Dim lanWeightSauceIDs(), lanWeightSauceModifierIDs(), ladSauceWeights()
	Dim lanWeightStyleIDs(), ladStyleRecipeIDs(), ladStyleWeights()
	Dim lanWeightTopperIDs(), ladTopperWeights()
	Dim lanWeightSideIDs(), ladSideWeights()
	Dim lnItemCount, lbHalveBaseCheese, lbIsBaseCheese, lnExtraCheeseItemID, ldBaseCheeseWeight, ldWeight, ldBaseCost, ldThisCost, ldTotalCost, ldStyleWeight, ldStyleCost
	
	Dim i, j, k, lnTmp
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	RecalculateOrderLineIdealCost = lbRet
End Function
%>
