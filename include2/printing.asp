<%
' **************************************************************************
' File: printing.asp
' Purpose: Functions for printing.
' Created: 7/19/2011 - TAM
' Description:
'	Include this file on any page where printing is necessary.
'	This file includes the following functions: SendToPrinter
'		PrintOrder, GetStoreMakeLinePrinters, GetStoreAltMakeLinePrinters, 
'		GetPrintStoreHeader, IsUnitAltPrinter, PrintSignatureCopies,
'		GetStoreCCPrinters, GetStoreStationPrinter, GetStorePrinters
'
' Revision History:
' 7/19/201 - Created
' **************************************************************************

' **************************************************************************
' Function: SendToPrinter
' Purpose: Sends data to a printer.
' Parameters:	psPrinterIP - The IP address of the printer
'				psPrintQueue - The print queue name
'				psData - The data to print
' Return: True if sucessful, False if not
' **************************************************************************
Function SendToPrinter(ByVal psPrinterIP, ByVal psData)
	Dim lbRet, loPrinter
	
	lbRet = FALSE
	On Error Resume Next
	
	Set loPrinter = CreateObject("ARLEpsonTM.ARLEpsonTMPrinter")
	If Err.Number = 0 And IsObject(loPrinter) Then
		If loPrinter.OpenPrinter(psPrinterIP) Then
			If loPrinter.GetPrinterStatus(psPrinterIP) = 0 Then
				If loPrinter.SendToPrinter(psData, Len(psData)) Then
					lbRet = TRUE
				End If
			End If
			
			loPrinter.ClosePrinter()
		End If
		
		Set loPrinter = Nothing
	End If
	
	SendToPrinter = lbRet
End Function

' **************************************************************************
' Function: PrintOrder
' Purpose: Sends an order to the appropriate printers.
' Parameters:	pnStoreUD - The StoreID
'				pnOrderID - The OrderID
'				pbNewOrder - Flag if this is a new order
' Return: True if sucessful, False if not
' **************************************************************************
Function PrintOrder(ByVal pnStoreID, ByVal pnOrderID, ByVal pbNewOrder)
	Dim lbRet, i, j, k, lnTmp, lasIPAddresses(), lsOutput, lsOutput2, lbDoAltMakeLine, lbAltMakeLine, ldTotalCost, ldTotalDiscount, lsOrderTypeDescription
	Dim lnSessionID, lsIPAddress, lnEmpID, lsRefID, ldtTransactionDate, ldtSubmitDate, ldtReleaseDate, ldtExpectedDate, lnStoreID, lnCustomerID, lsCustomerName, lsCustomerPhone, lnAddressID, lnOrderTypeID, lbIsPaid, lnPaymentTypeID, lsPaymentReference, lnAccountID, ldDeliveryCharge, ldDriverMoney, ldTax, ldTax2, ldTip, lnOrderStatusID, lsOrderNotes
	Dim lanOrderLineIDs(), lasOrderLineDescriptions(), lanQuantity(), ladCost(), ladDiscount()
	Dim lnUnitID, lnSpecialtyID, lnSizeID, lnStyleID, lnHalf1SauceID, lnHalf2SauceID, lnHalf1SauceModifierID, lnHalf2SauceModifierID, lsOrderLineNotes, lnQuantity, ldOrderLineCost, ldOrderLineDiscount, lnCouponID
	Dim lanOrderLineItemIDs(), lanItemIDs(), lanHalfIDs(), lasItemDescriptions(), lasItemShortDescriptions()
	Dim lanOrderLineTopperIDs(), lanTopperIDs(), lanTopperHalfIDs(), lasTopperDescriptions(), lasTopperShortDescriptions()
	Dim lanOrderLineFreeSideIDs(), lanFreeSideIDs(), lasFreeSideDescriptions(), lasFreeSideShortDescriptions()
	Dim lanOrderLineAddSideIDs(), lanAddSideIDs(), lasAddSideDescriptions(), lasAddSideShortDescriptions()
	Dim lanUnitSideIDs(), lasUnitSideDescriptions(), lasUnitSideShortDescriptions(), ladUnitSidePrices()
	Dim lanUnitItemIDs(), lasUnitItemDescriptions(), lasUnitItemShortDescriptions(), ladUnitItemOnSidePrice(), lanUnitItemCounts(), labUnitFreeItemFlags(), labUnitItemIsCheeses(), labUnitItemIsBaseCheeses(), labUnitItemIsExtraCheeses()
	Dim lanWeightItemIDs(), lanItemCounts1(), ladItemWeights1(), lanItemCounts2(), ladItemWeights2(), ladItemWeights3()
	Dim lanWeightSauceIDs(), lanWeightSauceModifierIDs(), ladSauceWeights()
	Dim lsAddress1, lsAddress2, lsCity, lsState, lsPostalCode, lsAddressNotes
	Dim lsAddressDescription, lsCustomerNotes
	Dim lnItemCount, lbHalveBaseCheese, lbIsBaseCheese, lnExtraCheeseItemID, ldBaseCheeseWeight, lsWeight, lsCoupons, lsUnit, lasUnitList(), lbUnitFound
	Dim lanFreeCouponIDs(), lasFreeCouponDescriptions(), lanFreeCouponCount()
	Dim lsBaseCheeseAfter
	
	lbRet = FALSE
	
	If GetOrderDetails(pnOrderID, lnSessionID, lsIPAddress, lnEmpID, lsRefID, ldtTransactionDate, ldtSubmitDate, ldtReleaseDate, ldtExpectedDate, lnStoreID, lnCustomerID, lsCustomerName, lsCustomerPhone, lnAddressID, lnOrderTypeID, lbIsPaid, lnPaymentTypeID, lsPaymentReference, lnAccountID, ldDeliveryCharge, ldDriverMoney, ldTax, ldTax2, ldTip, lnOrderStatusID, lsOrderNotes) Then
		If GetAddressDetails(lnAddressID, lnTmp, lsAddress1, lsAddress2, lsCity, lsState, lsPostalCode, lsAddressNotes) Then
			lsAddressDescription = ""
			lsCustomerNotes = ""
			
			If lnAddressID > 1 Then
				GetCustomerAddressDetails lnCustomerID, lnAddressID, lsAddressDescription, lsCustomerNotes
			End If
			
			If GetOrderLines(pnOrderID, lanOrderLineIDs, lasOrderLineDescriptions, lanQuantity, ladCost, ladDiscount) Then
				If GetOrderCoupons(pnOrderID, lsCoupons) Then
					If GetOrderFreeCoupons(pnOrderID, lanFreeCouponIDs, lasFreeCouponDescriptions, lanFreeCouponCount) Then
						lbRet = TRUE
						lbDoAltMakeLine = FALSE
						
						' Build for make line and alternate make line
						ldTotalCost = 0.00
						ldTotalDiscount = 0.00
						lsOrderTypeDescription = GetOrderTypeDescription(lnOrderTypeID)
						ReDim lasUnitList(0)
						lasUnitList(0) = ""
						
						' Start a new line
						lsOutput = CHR(10)
						
						' Add logo
						lsOutput = lsOutput & CHR(27) & CHR(97) & CHR(1) & CHR(28) & CHR(112) & CHR(1) & CHR(0) & CHR(27) & CHR(97) & CHR(0)
						
						' Set font size to 12
						lsOutput = lsOutput & CHR(10) & CHR(27) & CHR(33) & CHR(12)
						
						' Add header with center justification
						lsOutput = lsOutput & CHR(27) & CHR(97) & CHR(1)
						lsOutput = lsOutput & "vitos.com #" & pnStoreID & " " & GetEmployeeShortName(lnEmpID) & CHR(10)
						lsOutput = lsOutput & FormatDateTime(Now()) & " #" & pnOrderID & CHR(10)
						
						' Set font size to 18
						lsOutput = lsOutput & CHR(27) & CHR(33) & CHR(18)
						
						' If not a new order, indicate as a reprint in reverse mode
						If Not pbNewOrder Then
							lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(1)
							lsOutput = lsOutput & "**** REPRINT ****" & CHR(10)
							lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(0)
						End If
						lsOutput = lsOutput & CHR(10)
						
						' Return to left justify and add line items
						lsOutput = lsOutput & CHR(27) & CHR(97) & CHR(0)
						lsOutput2 = lsOutput
						For i = 0 To UBound(lanOrderLineIDs)
							If GetOrderLineDetails(lanOrderLineIDs(i), lnUnitID, lnSpecialtyID, lnSizeID, lnStyleID, lnHalf1SauceID, lnHalf2SauceID, lnHalf1SauceModifierID, lnHalf2SauceModifierID, lsOrderLineNotes, lnQuantity, ldOrderLineCost, ldOrderLineDiscount, lnCouponID) Then
								If GetOrderLineItems(lanOrderLineIDs(i), lnUnitID, lanOrderLineItemIDs, lanItemIDs, lanHalfIDs, lasItemDescriptions, lasItemShortDescriptions) Then
									If GetOrderLineToppers(lanOrderLineIDs(i), lanOrderLineTopperIDs, lanTopperIDs, lanTopperHalfIDs, lasTopperDescriptions, lasTopperShortDescriptions) Then
										If GetOrderLineFreeSides(lanOrderLineIDs(i), lanOrderLineFreeSideIDs, lanFreeSideIDs, lasFreeSideDescriptions, lasFreeSideShortDescriptions) Then
											If GetOrderLineAddSides(lanOrderLineIDs(i), lanOrderLineAddSideIDs, lanAddSideIDs, lasAddSideDescriptions, lasAddSideShortDescriptions) Then
												If GetStoreUnitSides(pnStoreID, lnUnitID, lanUnitSideIDs, lasUnitSideDescriptions, lasUnitSideShortDescriptions, ladUnitSidePrices) Then
													If GetStoreUnitItems(pnStoreID, lnUnitID, lanUnitItemIDs, lasUnitItemDescriptions, lasUnitItemShortDescriptions, ladUnitItemOnSidePrice, lanUnitItemCounts, labUnitFreeItemFlags, labUnitItemIsCheeses, labUnitItemIsBaseCheeses, labUnitItemIsExtraCheeses) Then
														If GetUnitSizeItemWeights(lnUnitID, lnSizeID, lanWeightItemIDs, lanItemCounts1, ladItemWeights1, lanItemCounts2, ladItemWeights2, ladItemWeights3) Then
															If GetUnitSizeSauceWeights(lnUnitID, lnSizeID, lanWeightSauceIDs, lanWeightSauceModifierIDs, ladSauceWeights) Then
																lbAltMakeLine = IsUnitAltPrinter(pnStoreID, lnUnitID)
																If lbAltMakeLine Then
																	lbDoAltMakeLine = TRUE
																End If
																
																ldTotalCost = ldTotalCost + (lnQuantity * ldOrderLineCost)
																ldTotalDiscount = ldTotalDiscount + (lnQuantity * ldOrderLineDiscount)
																
																lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(1)
																If lbAltMakeLine Then
																	lsOutput2 = lsOutput2 & CHR(29) & CHR(66) & CHR(1)
																End If
																If lnSizeID > 0 Then
																	lsOutput = lsOutput & GetSizeShortDescription(lnSizeID) & " "
																	If lbAltMakeLine Then
																		lsOutput2 = lsOutput2 & GetSizeShortDescription(lnSizeID) & " "
																	End If
																End If
																lsUnit = GetUnitShortDescription(lnUnitID)
																lbUnitFound = FALSE
																For k = 0 To UBound(lasUnitList)
																	If lasUnitList(k) = lsUnit Then
																		lbUnitFound = TRUE
																		Exit For
																	End If
																Next
																If Not lbUnitFound Then
																	If Len(lasUnitList(0)) > 0 Then
																		ReDim Preserve lasUnitList(UBound(lasUnitList) + 1)
																	End If
																	
																	lasUnitList(UBound(lasUnitList)) = lsUnit
																End If
																lsOutput = lsOutput & lsUnit
																If lnQuantity > 1 Then
																	lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(0)
																	lsOutput = lsOutput & CHR(27) & CHR(69) & CHR(1)
																	lsOutput = lsOutput & " " & lnQuantity & " @ "
																	lsOutput = lsOutput & CHR(27) & CHR(69) & CHR(0)
																	lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(1)
																Else
																	lsOutput = lsOutput & " "
																End If
																lsOutput = lsOutput & FormatCurrency(ldOrderLineCost - ldOrderLineDiscount) & CHR(10)
																If lbAltMakeLine Then
																	lsOutput2 = lsOutput2 & lsUnit
																	If lnQuantity > 1 Then
																		lsOutput2 = lsOutput2 & CHR(29) & CHR(66) & CHR(0)
																		lsOutput2 = lsOutput2 & CHR(27) & CHR(69) & CHR(1)
																		lsOutput2 = lsOutput2 & " " & lnQuantity & " @ "
																		lsOutput2 = lsOutput2 & CHR(27) & CHR(69) & CHR(0)
																		lsOutput2 = lsOutput2 & CHR(29) & CHR(66) & CHR(1)
																	Else
																		lsOutput2 = lsOutput2 & " "
																	End If
																	lsOutput2 = lsOutput2 & FormatCurrency(ldOrderLineCost - ldOrderLineDiscount) & CHR(10)
																End If
																If lnSpecialtyID > 0 Then
																	lsOutput = lsOutput & GetSpecialtyShortDescription(lnSpecialtyID) & CHR(10)
																	If lbAltMakeLine Then
																		lsOutput2 = lsOutput2 & GetSpecialtyShortDescription(lnSpecialtyID) & CHR(10)
																	End If
																End If
																If lnStyleID > 0 Then
																	lsOutput = lsOutput & GetStyleShortDescription(lnStyleID) & CHR(10)
																	If lbAltMakeLine Then
																		lsOutput2 = lsOutput2 & GetStyleShortDescription(lnStyleID) & CHR(10)
																	End If
																End If
																lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(0)
																If lbAltMakeLine Then
																	lsOutput2 = lsOutput2 & CHR(29) & CHR(66) & CHR(0)
																End If
																
																If lnUnitID = 1 Then
																	lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(1)
																	If lbAltMakeLine Then
																		lsOutput2 = lsOutput2 & CHR(29) & CHR(66) & CHR(1)
																	End If
																	If lnHalf1SauceID > 0 Then
																		lsWeight = "     "
																		For k = 0 To UBound(lanWeightSauceIDs)
																			If lnHalf1SauceID = lanWeightSauceIDs(k) And lnHalf1SauceModifierID = lanWeightSauceModifierIDs(k) Then
																				If (lnHalf1SauceID = lnHalf2SauceID) And (lnHalf1SauceModifierID = lnHalf2SauceModifierID) Then
																					lsWeight = FormatCurrency(ladSauceWeights(k))
																				Else
																					lsWeight = FormatCurrency(ladSauceWeights(k) / 2)
																				End If
																				lsWeight = Right(lsWeight, (Len(lsWeight) - 1))
																				lsWeight = Right("     " & lsWeight, 5)
																				
																				Exit For
																			End If
																		Next
																		
																		If (lnHalf1SauceID = lnHalf2SauceID) And (lnHalf1SauceModifierID = lnHalf2SauceModifierID) Then
																			If lnHalf1SauceModifierID > 0 Then
																				lsOutput = lsOutput & Left(GetSauceShortDescription(lnHalf1SauceID) & " " & GetSauceModifierShortDescription(lnHalf1SauceModifierID) & String(27, " "), 27) & " " & lsWeight
																				If lbAltMakeLine Then
																					lsOutput2 = lsOutput2 & Left(GetSauceShortDescription(lnHalf1SauceID) & " " & GetSauceModifierShortDescription(lnHalf1SauceModifierID) & String(27, " "), 27) & " " & lsWeight
																				End If
																			Else
																				lsOutput = lsOutput & Left(GetSauceShortDescription(lnHalf1SauceID) & String(27, " "), 27) & " " & lsWeight
																				If lbAltMakeLine Then
																					lsOutput2 = lsOutput2 & Left(GetSauceShortDescription(lnHalf1SauceID) & String(27, " "), 27) & " " & lsWeight
																				End If
																			End If
																		Else
																			If lnHalf1SauceModifierID > 0 Then
																				lsOutput = lsOutput & "HALF 1: " & Left(GetSauceShortDescription(lnHalf1SauceID) & " " & GetSauceModifierShortDescription(lnHalf1SauceModifierID) & String(19, " "), 19) & " " & lsWeight
																				If lbAltMakeLine Then
																					lsOutput2 = lsOutput2 & "HALF 1: " & Left(GetSauceShortDescription(lnHalf1SauceID) & " " & GetSauceModifierShortDescription(lnHalf1SauceModifierID) & String(19, " "), 19) & " " & lsWeight
																				End If
																			Else
																				lsOutput = lsOutput & "HALF 1: " & Left(GetSauceShortDescription(lnHalf1SauceID) & String(19, " "), 19) & " " & lsWeight
																				If lbAltMakeLine Then
																					lsOutput2 = lsOutput2 & "HALF 1: " & Left(GetSauceShortDescription(lnHalf1SauceID) & String(19, " "), 19) & " " & lsWeight
																				End If
																			End If
																		End If
																		lsOutput = lsOutput & CHR(10)
																		If lbAltMakeLine Then
																			lsOutput2 = lsOutput2 & CHR(10)
																		End If
																	End If
																	If lnHalf2SauceID > 0 And ((lnHalf1SauceID <> lnHalf2SauceID) Or (lnHalf1SauceModifierID <> lnHalf2SauceModifierID)) Then
																		lsWeight = "     "
																		For k = 0 To UBound(lanWeightSauceIDs)
																			If lnHalf2SauceID = lanWeightSauceIDs(k) And lnHalf2SauceModifierID = lanWeightSauceModifierIDs(k) Then
																				lsWeight = FormatCurrency(ladSauceWeights(k) / 2)
																				lsWeight = Right(lsWeight, (Len(lsWeight) - 1))
																				lsWeight = Right("     " & lsWeight, 5)
																				
																				Exit For
																			End If
																		Next
																		
																		If lnHalf2SauceModifierID > 0 Then
																			lsOutput = lsOutput & "HALF 2: " & Left(GetSauceShortDescription(lnHalf2SauceID) & " " & GetSauceModifierShortDescription(lnHalf2SauceModifierID) & String(19, " "), 19) & " " & lsWeight
																			If lbAltMakeLine Then
																				lsOutput2 = lsOutput2 & "HALF 2: " & Left(GetSauceShortDescription(lnHalf2SauceID) & " " & GetSauceModifierShortDescription(lnHalf2SauceModifierID) & String(19, " "), 19) & " " & lsWeight
																			End If
																		Else
																			lsOutput = lsOutput & "HALF 2: " & Left(GetSauceShortDescription(lnHalf2SauceID) & String(19, " "), 19) & " " & lsWeight
																			If lbAltMakeLine Then
																				lsOutput2 = lsOutput2 & "HALF 2: " & Left(GetSauceShortDescription(lnHalf2SauceID) & String(19, " "), 19) & " " & lsWeight
																			End If
																		End If
																		lsOutput = lsOutput & CHR(10)
																		If lbAltMakeLine Then
																			lsOutput2 = lsOutput2 & CHR(10)
																		End If
																	End If
																	lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(0)
																	If lbAltMakeLine Then
																		lsOutput2 = lsOutput2 & CHR(29) & CHR(66) & CHR(0)
																	End If
																End If
																
																lnItemCount = 0
																lbHalveBaseCheese = FALSE
																For j = 0 To UBound(lanItemIDs)
																	If lanItemIDs(j) > 0 Then
																		For k = 0 To UBound(lanUnitItemIDs)
																			If lanItemIDs(j) = lanUnitItemIDs(k) Then
																				If (Not labUnitFreeItemFlags(k)) And (Not labUnitItemIsCheeses(k)) Then
																					Select Case lanHalfIDs(j)
																						Case 0:
																							lnItemCount = lnItemCount + lanUnitItemCounts(k)
																						Case 1:
																							lnItemCount = lnItemCount + (lanUnitItemCounts(k) / 2)
																						Case 2:
																							lnItemCount = lnItemCount + (lanUnitItemCounts(k) / 2)
																					End Select
																				End If
																				
																				If labUnitItemIsCheeses(k) And (Not labUnitItemIsBaseCheeses(k)) And (Not labUnitItemIsExtraCheeses(k)) Then
																					lbHalveBaseCheese = TRUE
																				End If
																				
																				Exit For
																			End If
																		Next
																	End If
																Next
																
																For j = 0 To UBound(lanTopperIDs)
																	If lanTopperIDs(j) > 0 Then
																		If IsTopperBeforeItems(lnUnitID, lanTopperIDs(j)) Then
																			Select Case lanTopperHalfIDs(j)
																				Case 0
																					lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(1) & lasTopperShortDescriptions(j) & CHR(29) & CHR(66) & CHR(0) & CHR(10)
																					If lbAltMakeLine Then
																						lsOutput2 = lsOutput2 & CHR(29) & CHR(66) & CHR(1) & lasTopperShortDescriptions(j) & CHR(29) & CHR(66) & CHR(0) & CHR(10)
																					End If
																				Case 1
																					lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(1) & "HALF 1: " & lasTopperShortDescriptions(j) & CHR(29) & CHR(66) & CHR(0) & CHR(10)
																					If lbAltMakeLine Then
																						lsOutput2 = lsOutput2 & CHR(29) & CHR(66) & CHR(1) & "HALF 1: " & lasTopperShortDescriptions(j) & CHR(29) & CHR(66) & CHR(0) & CHR(10)
																					End If
																				Case 2
																					lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(1) & "HALF 2: " & lasTopperShortDescriptions(j) & CHR(29) & CHR(66) & CHR(0) & CHR(10)
																					If lbAltMakeLine Then
																						lsOutput2 = lsOutput2 & CHR(29) & CHR(66) & CHR(1) & "HALF 2: " & lasTopperShortDescriptions(j) & CHR(29) & CHR(66) & CHR(0) & CHR(10)
																					End If
																				Case 3
																					lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(1) & "ON SIDE: " & lasTopperShortDescriptions(j) & CHR(29) & CHR(66) & CHR(0) & CHR(10)
																					If lbAltMakeLine Then
																						lsOutput2 = lsOutput2 & CHR(29) & CHR(66) & CHR(1) & "ON SIDE: " & lasTopperShortDescriptions(j) & CHR(29) & CHR(66) & CHR(0) & CHR(10)
																					End If
																			End Select
																		End If
																	End If
																Next
																
																lsBaseCheeseAfter = ""
																For j = 0 To UBound(lanItemIDs)
																	If lanItemIDs(j) > 0 Then
																		lbIsBaseCheese = FALSE
																		For k = 0 To UBound(lanUnitItemIDs)
																			If lanItemIDs(j) = lanUnitItemIDs(k) Then
																				If labUnitItemIsBaseCheeses(k) Then
																					lbIsBaseCheese = TRUE
																				End If
																				
																				Exit For
																			End If
																		Next
																		
																		If lbIsBaseCheese And UBound(lanItemIDs) = 0 Then
																			lnExtraCheeseItemID = 0
																			For k = 0 To UBound(lanUnitItemIDs)
																				If labUnitItemIsExtraCheeses(k) Then
																					lnExtraCheeseItemID = lanUnitItemIDs(k)
																					lbIsBaseCheese = TRUE
																					
																					Exit For
																				End If
																			Next
																			
																			ldBaseCheeseWeight = 0.00
																			For k = 0 To UBound(lanWeightItemIDs)
																				If lanItemIDs(j) = lanWeightItemIDs(k) Then
																					ldBaseCheeseWeight = ladItemWeights1(k)
																					
																					Exit For
																				End If
																			Next
																			
																			If lnExtraCheeseItemID > 0 Then
																				For k = 0 To UBound(lanWeightItemIDs)
																					If lnExtraCheeseItemID = lanWeightItemIDs(k) Then
																						ldBaseCheeseWeight = ldBaseCheeseWeight + ladItemWeights1(k)
																						
																						Exit For
																					End If
																				Next
																			End If
																			
																			lsWeight = FormatCurrency(ldBaseCheeseWeight)
																			lsWeight = Right(lsWeight, (Len(lsWeight) - 1))
																			lsWeight = Right("     " & lsWeight, 5)
																		Else
																			lsWeight = "     "
																			For k = 0 To UBound(lanWeightItemIDs)
																				If lanItemIDs(j) = lanWeightItemIDs(k) Then
																					If lbIsBaseCheese Then
																						If lbHalveBaseCheese Then
																							lsWeight = FormatCurrency(ladItemWeights1(k) / 2)
																						Else
																							lsWeight = FormatCurrency(ladItemWeights1(k))
																						End If
																					Else
																						If lnItemCount <= lanItemCounts1(k) Then
																							Select Case lanHalfIDs(j)
																								Case 0
																									lsWeight = FormatCurrency(ladItemWeights1(k))
																								Case 1, 2
																									lsWeight = FormatCurrency(ladItemWeights1(k) / 2)
																							End Select
																						Else
																							If lnItemCount <= lanItemCounts2(k) Then
																								Select Case lanHalfIDs(j)
																									Case 0
																										lsWeight = FormatCurrency(ladItemWeights2(k))
																									Case 1, 2
																										lsWeight = FormatCurrency(ladItemWeights2(k) / 2)
																								End Select
																							Else
																								Select Case lanHalfIDs(j)
																									Case 0
																										lsWeight = FormatCurrency(ladItemWeights3(k))
																									Case 1, 2
																										lsWeight = FormatCurrency(ladItemWeights3(k) / 2)
																								End Select
																							End If
																						End If
																					End If
																					lsWeight = Right(lsWeight, (Len(lsWeight) - 1))
																					lsWeight = Right("     " & lsWeight, 5)
																					
																					Exit For
																				End If
																			Next
																		End If
																		
'																		If lbIsBaseCheese Then
'																			Select Case lanHalfIDs(j)
'																				Case 0
'																					lsBaseCheeseAfter = Left(lasItemShortDescriptions(j) & String(27, " "), 27) & " " & lsWeight & CHR(10)
'																				Case 1
'																					lsBaseCheeseAfter = CHR(29) & CHR(66) & CHR(1) & "HALF 1:" & CHR(29) & CHR(66) & CHR(0) & " " & Left(lasItemShortDescriptions(j) & String(19, " "), 19) & " " & lsWeight & CHR(10)
'																				Case 2
'																					lsBaseCheeseAfter = CHR(29) & CHR(66) & CHR(1) & "HALF 2:" & CHR(29) & CHR(66) & CHR(0) & " " & Left(lasItemShortDescriptions(j) & String(19, " "), 19) & " " & lsWeight & CHR(10)
'																				Case 3
'																					lsBaseCheeseAfter = CHR(29) & CHR(66) & CHR(1) & "ON SIDE:" & CHR(29) & CHR(66) & CHR(0) & " " & lasItemShortDescriptions(j) & CHR(10)
'																			End Select
'																		Else
																			Select Case lanHalfIDs(j)
																				Case 0
																					lsOutput = lsOutput & Left(lasItemShortDescriptions(j) & String(27, " "), 27) & " " & lsWeight & CHR(10)
																					If lbAltMakeLine Then
																						lsOutput2 = lsOutput2 & Left(lasItemShortDescriptions(j) & String(27, " "), 27) & " " & lsWeight & CHR(10)
																					End If
																				Case 1
																					lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(1) & "HALF 1:" & CHR(29) & CHR(66) & CHR(0) & " " & Left(lasItemShortDescriptions(j) & String(19, " "), 19) & " " & lsWeight & CHR(10)
																					If lbAltMakeLine Then
																						lsOutput2 = lsOutput2 & CHR(29) & CHR(66) & CHR(1) & "HALF 1:" & CHR(29) & CHR(66) & CHR(0) & " " & Left(lasItemShortDescriptions(j) & String(19, " "), 19) & " " & lsWeight & CHR(10)
																					End If
																				Case 2
																					lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(1) & "HALF 2:" & CHR(29) & CHR(66) & CHR(0) & " " & Left(lasItemShortDescriptions(j) & String(19, " "), 19) & " " & lsWeight & CHR(10)
																					If lbAltMakeLine Then
																						lsOutput2 = lsOutput2 & CHR(29) & CHR(66) & CHR(1) & "HALF 2:" & CHR(29) & CHR(66) & CHR(0) & " " & Left(lasItemShortDescriptions(j) & String(19, " "), 19) & " " & lsWeight & CHR(10)
																					End If
																				Case 3
																					lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(1) & "ON SIDE:" & CHR(29) & CHR(66) & CHR(0) & " " & lasItemShortDescriptions(j) & CHR(10)
																					If lbAltMakeLine Then
																						lsOutput2 = lsOutput2 & CHR(29) & CHR(66) & CHR(1) & "ON SIDE:" & CHR(29) & CHR(66) & CHR(0) & " " & lasItemShortDescriptions(j) & CHR(10)
																					End If
																			End Select
'																		End If
																	End If
																Next
																lsOutput = lsOutput & lsBaseCheeseAfter
																If lbAltMakeLine Then
																	lsOutput2 = lsOutput2 & lsBaseCheeseAfter
																End If
																
																If lnUnitID <> 1 Then
																	lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(1)
																	If lbAltMakeLine Then
																		lsOutput2 = lsOutput2 & CHR(29) & CHR(66) & CHR(1)
																	End If
																	If lnHalf1SauceID > 0 Then
																		lsWeight = "     "
																		For k = 0 To UBound(lanWeightSauceIDs)
																			If lnHalf1SauceID = lanWeightSauceIDs(k) And lnHalf1SauceModifierID = lanWeightSauceModifierIDs(k) Then
																				If (lnHalf1SauceID = lnHalf2SauceID) And (lnHalf1SauceModifierID = lnHalf2SauceModifierID) Then
																					lsWeight = FormatCurrency(ladSauceWeights(k))
																				Else
																					lsWeight = FormatCurrency(ladSauceWeights(k) / 2)
																				End If
																				lsWeight = Right(lsWeight, (Len(lsWeight) - 1))
																				lsWeight = Right("     " & lsWeight, 5)
																				
																				Exit For
																			End If
																		Next
																		
																		If (lnHalf1SauceID = lnHalf2SauceID) And (lnHalf1SauceModifierID = lnHalf2SauceModifierID) Then
																			If lnHalf1SauceModifierID > 0 Then
																				lsOutput = lsOutput & Left(GetSauceShortDescription(lnHalf1SauceID) & " " & GetSauceModifierShortDescription(lnHalf1SauceModifierID) & String(27, " "), 27) & " " & lsWeight
																				If lbAltMakeLine Then
																					lsOutput2 = lsOutput2 & Left(GetSauceShortDescription(lnHalf1SauceID) & " " & GetSauceModifierShortDescription(lnHalf1SauceModifierID) & String(27, " "), 27) & " " & lsWeight
																				End If
																			Else
																				lsOutput = lsOutput & Left(GetSauceShortDescription(lnHalf1SauceID) & String(27, " "), 27) & " " & lsWeight
																				If lbAltMakeLine Then
																					lsOutput2 = lsOutput2 & Left(GetSauceShortDescription(lnHalf1SauceID) & String(27, " "), 27) & " " & lsWeight
																				End If
																			End If
																		Else
																			If lnHalf1SauceModifierID > 0 Then
																				lsOutput = lsOutput & "HALF 1: " & Left(GetSauceShortDescription(lnHalf1SauceID) & " " & GetSauceModifierShortDescription(lnHalf1SauceModifierID) & String(19, " "), 19) & " " & lsWeight
																				If lbAltMakeLine Then
																					lsOutput2 = lsOutput2 & "HALF 1: " & Left(GetSauceShortDescription(lnHalf1SauceID) & " " & GetSauceModifierShortDescription(lnHalf1SauceModifierID) & String(19, " "), 19) & " " & lsWeight
																				End If
																			Else
																				lsOutput = lsOutput & "HALF 1: " & Left(GetSauceShortDescription(lnHalf1SauceID) & String(19, " "), 19) & " " & lsWeight
																				If lbAltMakeLine Then
																					lsOutput2 = lsOutput2 & "HALF 1: " & Left(GetSauceShortDescription(lnHalf1SauceID) & String(19, " "), 19) & " " & lsWeight
																				End If
																			End If
																		End If
																		lsOutput = lsOutput & CHR(10)
																		If lbAltMakeLine Then
																			lsOutput2 = lsOutput2 & CHR(10)
																		End If
																	End If
																	If lnHalf2SauceID > 0 And ((lnHalf1SauceID <> lnHalf2SauceID) Or (lnHalf1SauceModifierID <> lnHalf2SauceModifierID)) Then
																		lsWeight = "     "
																		For k = 0 To UBound(lanWeightSauceIDs)
																			If lnHalf2SauceID = lanWeightSauceIDs(k) And lnHalf2SauceModifierID = lanWeightSauceModifierIDs(k) Then
																				lsWeight = FormatCurrency(ladSauceWeights(k) / 2)
																				lsWeight = Right(lsWeight, (Len(lsWeight) - 1))
																				lsWeight = Right("     " & lsWeight, 5)
																				
																				Exit For
																			End If
																		Next
																		
																		If lnHalf2SauceModifierID > 0 Then
																			lsOutput = lsOutput & "HALF 2: " & Left(GetSauceShortDescription(lnHalf2SauceID) & " " & GetSauceModifierShortDescription(lnHalf2SauceModifierID) & String(19, " "), 19) & " " & lsWeight
																			If lbAltMakeLine Then
																				lsOutput2 = lsOutput2 & "HALF 2: " & Left(GetSauceShortDescription(lnHalf2SauceID) & " " & GetSauceModifierShortDescription(lnHalf2SauceModifierID) & String(19, " "), 19) & " " & lsWeight
																			End If
																		Else
																			lsOutput = lsOutput & "HALF 2: " & Left(GetSauceShortDescription(lnHalf2SauceID) & String(19, " "), 19) & " " & lsWeight
																			If lbAltMakeLine Then
																				lsOutput2 = lsOutput2 & "HALF 2: " & Left(GetSauceShortDescription(lnHalf2SauceID) & String(19, " "), 19) & " " & lsWeight
																			End If
																		End If
																		lsOutput = lsOutput & CHR(10)
																		If lbAltMakeLine Then
																			lsOutput2 = lsOutput2 & CHR(10)
																		End If
																	End If
																	lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(0)
																	If lbAltMakeLine Then
																		lsOutput2 = lsOutput2 & CHR(29) & CHR(66) & CHR(0)
																	End If
																End If
																
																For j = 0 To UBound(lanTopperIDs)
																	If lanTopperIDs(j) > 0 Then
																		If Not IsTopperBeforeItems(lnUnitID, lanTopperIDs(j)) Then
																			Select Case lanTopperHalfIDs(j)
																				Case 0
																					lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(1) & lasTopperShortDescriptions(j) & CHR(29) & CHR(66) & CHR(0) & CHR(10)
																					If lbAltMakeLine Then
																						lsOutput2 = lsOutput2 & CHR(29) & CHR(66) & CHR(1) & lasTopperShortDescriptions(j) & CHR(29) & CHR(66) & CHR(0) & CHR(10)
																					End If
																				Case 1
																					lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(1) & "HALF 1: " & lasTopperShortDescriptions(j) & CHR(29) & CHR(66) & CHR(0) & CHR(10)
																					If lbAltMakeLine Then
																						lsOutput2 = lsOutput2 & CHR(29) & CHR(66) & CHR(1) & "HALF 1: " & lasTopperShortDescriptions(j) & CHR(29) & CHR(66) & CHR(0) & CHR(10)
																					End If
																				Case 2
																					lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(1) & "HALF 2: " & lasTopperShortDescriptions(j) & CHR(29) & CHR(66) & CHR(0) & CHR(10)
																					If lbAltMakeLine Then
																						lsOutput2 = lsOutput2 & CHR(29) & CHR(66) & CHR(1) & "HALF 2: " & lasTopperShortDescriptions(j) & CHR(29) & CHR(66) & CHR(0) & CHR(10)
																					End If
																				Case 3
																					lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(1) & "ON SIDE: " & lasTopperShortDescriptions(j) & CHR(29) & CHR(66) & CHR(0) & CHR(10)
																					If lbAltMakeLine Then
																						lsOutput2 = lsOutput2 & CHR(29) & CHR(66) & CHR(1) & "ON SIDE: " & lasTopperShortDescriptions(j) & CHR(29) & CHR(66) & CHR(0) & CHR(10)
																					End If
																			End Select
																		End If
																	End If
																Next
																
																For j = 0 To UBound(lanFreeSideIDs)
																	If lanFreeSideIDs(j) > 0 Then
																		lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(1) & "    " & Left(lasFreeSideShortDescriptions(j) & String(23, " "), 23) & " (N/C)" & CHR(29) & CHR(66) & CHR(0) & CHR(10)
																		If lbAltMakeLine Then
																			lsOutput2 = lsOutput2 & CHR(29) & CHR(66) & CHR(1) & "    " & Left(lasFreeSideShortDescriptions(j) & String(23, " "), 23) & " (N/C)" & CHR(29) & CHR(66) & CHR(0) & CHR(10)
																		End If
																	End If
																Next
																For j = 0 To UBound(lanAddSideIDs)
																	If lanAddSideIDs(j) > 0 Then
																		For k = 0 To UBound(lanUnitSideIDs)
																			If lanUnitSideIDs(k) = lanAddSideIDs(j) Then
																				lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(1) & "    " & Left(lasAddSideShortDescriptions(j) & String(23, " "), 23) & " " & FormatCurrency(ladUnitSidePrices(k)) & CHR(29) & CHR(66) & CHR(0) & CHR(10)
																				If lbAltMakeLine Then
																					lsOutput2 = lsOutput2 & CHR(29) & CHR(66) & CHR(1) & "    " & Left(lasAddSideShortDescriptions(j) & String(23, " "), 23) & " " & FormatCurrency(ladUnitSidePrices(k)) & CHR(29) & CHR(66) & CHR(0) & CHR(10)
																				End If
																				
																				Exit For
																			End If
																		Next
																	End If
																Next
																
																If Len(lsOrderLineNotes) > 0 Then
																	lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(1) & "    " & lsOrderLineNotes & CHR(29) & CHR(66) & CHR(0) & CHR(10)
																	If lbAltMakeLine Then
																		lsOutput2 = lsOutput2 & CHR(29) & CHR(66) & CHR(1) & "    " & lsOrderLineNotes & CHR(29) & CHR(66) & CHR(0) & CHR(10)
																	End If
																End If
																
																lsOutput = lsOutput & CHR(10)
																If lbAltMakeLine Then
																	lsOutput2 = lsOutput2 & CHR(10)
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
						
						If lbRet Then
							' Add totals
							lsOutput = lsOutput & "This sale:            " & Replace(FormatCurrency(ldTotalCost), "$", "$" & String(11 - Len(FormatCurrency(ldTotalCost)), " ")) & CHR(10)
							lsOutput = lsOutput & "Tax 1:                " & Replace(FormatCurrency(ldTax), "$", "$" & String(11 - Len(FormatCurrency(ldTax)), " ")) & CHR(10)
							If ldTax2 <> 0 Then
								lsOutput = lsOutput & "Tax 2:                " & Replace(FormatCurrency(ldTax2), "$", "$" & String(11 - Len(FormatCurrency(ldTax2)), " ")) & CHR(10)
							End If
							If ldTotalDiscount <> 0 Then
								lsOutput = lsOutput & "Discount:             " & Replace(FormatCurrency(ldTotalDiscount), "$", "$" & String(11 - Len(FormatCurrency(ldTotalDiscount)), " ")) & CHR(10)
							End If
							If lnOrderTypeID = 1 Then
								lsOutput = lsOutput & Left(lsOrderTypeDescription & ":" & String(22, " "), 22) & Replace(FormatCurrency(ldDeliveryCharge), "$", "$" & String(11 - Len(FormatCurrency(ldDeliveryCharge)), " ")) & CHR(10)
							End If
							lsOutput = lsOutput & String(42, "-") & CHR(10)
							lsOutput = lsOutput & "Ticket Total:         " & Replace(FormatCurrency((ldTotalCost + ldTax + ldTax2 + ldDeliveryCharge) - ldTotalDiscount), "$", "$" & String(11 - Len(FormatCurrency((ldTotalCost + ldTax + ldTax2 + ldDeliveryCharge) - ldTotalDiscount)), " ")) & CHR(10)
							lsOutput = lsOutput & String(42, "-") & CHR(10)
							lsOutput2 = lsOutput2 & String(42, "-") & CHR(10)
							
							If ldTip <> 0 Then
								lsOutput = lsOutput & "Tip:                  " & Replace(FormatCurrency(ldTip), "$", "$" & String(11 - Len(FormatCurrency(ldTip)), " ")) & CHR(10)
								lsOutput = lsOutput & "Grand Total:          " & Replace(FormatCurrency((ldTotalCost + ldTax + ldTax2 + ldDeliveryCharge + ldTip) - ldTotalDiscount), "$", "$" & String(11 - Len(FormatCurrency((ldTotalCost + ldTax + ldTax2 + ldDeliveryCharge + ldTip) - ldTotalDiscount)), " ")) & CHR(10)
							End If
							
							If Len(lsCoupons) > 0 Then
								lsOutput = lsOutput & lsCoupons & CHR(10)
							End If
							lsOutput = lsOutput & CHR(10)
							
							' Indicate payment reference
							If lbIsPaid Then
								If (lnPaymentTypeID = 2 Or lnPaymentTypeID = 3) Then
									If Len(lsPaymentReference) > 0 Then
										lsOutput = lsOutput & "Payment Ref: " & lsPaymentReference & CHR(10)
										lsOutput = lsOutput & CHR(10)
									End If
								Else
									If lnPaymentTypeID = 4 Then
										lsOutput = lsOutput & "Placed On Account #" & lnAccountID & CHR(10)
										lsOutput = lsOutput & CHR(10)
									End If
								End If
							End If
							
							' Indicate pickup/delivery
							lsOutput = lsOutput & lsOrderTypeDescription & CHR(10)
							lsOutput2 = lsOutput2 & lsOrderTypeDescription & CHR(10)
							
							' Add customer and address information
							lsOutput = lsOutput & CHR(27) & CHR(69) & CHR(1)
							lsOutput2 = lsOutput2 & CHR(27) & CHR(69) & CHR(1)
							lsOutput = lsOutput & lsCustomerName & CHR(10)
							lsOutput2 = lsOutput2 & lsCustomerName & CHR(10)
							Select Case lnOrderTypeID
								Case 1
									If Len(lsAddress2) = 0 Then
										If Len(lsAddress1) > 0 Then
											lsOutput = lsOutput & lsAddress1 & CHR(10)
											lsOutput2 = lsOutput2 & lsAddress1 & CHR(10)
										End If
									Else
										lsOutput = lsOutput & lsAddress1 & " #" & lsAddress2 & CHR(10)
										lsOutput2 = lsOutput2 & lsAddress1 & " #" & lsAddress2 & CHR(10)
									End If
									
									If Len(lsCity) > 0 Or Len(lsState) > 0 Or Len(lsPostalCode) > 0 Then
										lsOutput = lsOutput & lsCity & ", " & lsState & " " & lsPostalCode & CHR(10)
										lsOutput2 = lsOutput2 & lsCity & ", " & lsState & " " & lsPostalCode & CHR(10)
									End If
									
									If Len(lsCustomerPhone) > 0 Then
										If Len(lsCustomerPhone) <= 7 Then
											lsOutput = lsOutput & Left(lsCustomerPhone, 3) & "-" & Mid(lsCustomerPhone, 4) & CHR(10)
											lsOutput2 = lsOutput2 & Left(lsCustomerPhone, 3) & "-" & Mid(lsCustomerPhone, 4) & CHR(10)
										Else
											lsOutput = lsOutput & "(" & Left(lsCustomerPhone, 3) & ") " & Mid(lsCustomerPhone, 4, 3) & "-" & Mid(lsCustomerPhone, 7) & CHR(10)
											lsOutput2 = lsOutput2 & "(" & Left(lsCustomerPhone, 3) & ") " & Mid(lsCustomerPhone, 4, 3) & "-" & Mid(lsCustomerPhone, 7) & CHR(10)
										End If
									End If
									
									If Len(lsAddressNotes) > 0 Then
										lsOutput = lsOutput & lsAddressNotes & CHR(10)
									End If
									
									If Len(lsCustomerNotes) > 0 Then
										lsOutput = lsOutput & lsCustomerNotes & CHR(10)
									End If
								Case 2
									If Len(lsCustomerPhone) > 0 Then
										If Len(lsCustomerPhone) <= 7 Then
											lsOutput = lsOutput & Left(lsCustomerPhone, 3) & "-" & Mid(lsCustomerPhone, 4) & CHR(10)
											lsOutput2 = lsOutput2 & Left(lsCustomerPhone, 3) & "-" & Mid(lsCustomerPhone, 4) & CHR(10)
										Else
											lsOutput = lsOutput & "(" & Left(lsCustomerPhone, 3) & ") " & Mid(lsCustomerPhone, 4, 3) & "-" & Mid(lsCustomerPhone, 7) & CHR(10)
											lsOutput2 = lsOutput2 & "(" & Left(lsCustomerPhone, 3) & ") " & Mid(lsCustomerPhone, 4, 3) & "-" & Mid(lsCustomerPhone, 7) & CHR(10)
										End If
									End If
							End Select
							lsOutput = lsOutput & CHR(27) & CHR(69) & CHR(0)
							lsOutput2 = lsOutput2 & CHR(27) & CHR(69) & CHR(0)
							lsOutput = lsOutput & CHR(10)
							lsOutput = lsOutput & CHR(10)
							lsOutput = lsOutput & CHR(10)
							
							If Len(lsOrderNotes) > 0 Then
								lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(1)
								lsOutput = lsOutput & lsOrderNotes & CHR(10)
								lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(0)
								lsOutput = lsOutput & CHR(10)
								lsOutput = lsOutput & CHR(10)
							End If
							
							If Len(lasUnitList(0)) > 0 Then
								lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(1)
								lsOutput = lsOutput & "Units: " & lasUnitList(0)
								For i = 1 To UBound(lasUnitList)
									lsOutput = lsOutput & ", " & lasUnitList(i)
								Next
								lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(0)
								lsOutput = lsOutput & CHR(10)
								lsOutput = lsOutput & CHR(10)
							End If
							
							If lanFreeCouponIDs(0) > 0 Then
								lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(1)
								lsOutput = lsOutput & "Free Cards: "
								For i = 0 To UBound(lanFreeCouponIDs)
									lsOutput = lsOutput & CHR(10) & lanFreeCouponCount(i) & " - " & lasFreeCouponDescriptions(i)
								Next
								lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(0)
								lsOutput = lsOutput & CHR(10)
								lsOutput = lsOutput & CHR(10)
							End If
							
							' Indicate payment status and repeat ticket number with center justification
							lsOutput = lsOutput & CHR(27) & CHR(97) & CHR(1)
							If lbIsPaid Then
								lsOutput = lsOutput & "Ticket is paid for" & CHR(10)
							Else
								lsOutput = lsOutput & "Ticket NOT paid for" & CHR(10)
							End If
							
							If Not lbIsPaid And lnOrderTypeID = 1 Then
								If (lnPaymentTypeID = 1 Or lnPaymentTypeID = 2) Then
									lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(1)
									If IsStoreChecksOK(lnStoreID) Then
										If lnCustomerID > 1 Then
											If IsCustomerCheckOK(lnCustomerID) Then
												If DateDiff("d", GetLastCustomerOrderDate(lnCustomerID), Now()) > gnCheckAcceptMaxDaysSinceOrdered Then
													lsOutput = lsOutput & "Our check acceptance policy is that the" & CHR(10)
													lsOutput = lsOutput & "address on the check must match the" & CHR(10)
													lsOutput = lsOutput & "delivery address." & CHR(10)
												End If
											Else
												lsOutput = lsOutput & " ** NO CHECKS ** " & CHR(10)
											End If
										Else
											lsOutput = lsOutput & "Our check acceptance policy is that the" & CHR(10)
											lsOutput = lsOutput & "address on the check must match the" & CHR(10)
											lsOutput = lsOutput & "delivery address." & CHR(10)
										End If
									Else
										lsOutput = lsOutput & "Sorry, we do not accept checks." & CHR(10)
									End If
									lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(0)
								End If
							End If
							
							lsOutput = lsOutput & "Ticket #" & pnOrderID & CHR(10)
							
							If ldtExpectedDate <> DateValue("1/1/1900") Then
								lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(1) & CHR(10) & "Order Expected: " & ldtExpectedDate & CHR(29) & CHR(66) & CHR(0) & CHR(10)
							End If
							
							' Finish
							lsOutput = lsOutput & CHR(27) & CHR(97) & CHR(1)
							lsOutput = lsOutput & CHR(10) & "Find Us on Facebook" & CHR(10)
							lsOutput = lsOutput & "vitos.com/facebook" & CHR(10)
							lsOutput = lsOutput & CHR(10) & CHR(10) & CHR(10) & CHR(10) & CHR(10) & CHR(10) & CHR(10) & CHR(10) & CHR(10)
							lsOutput2 = lsOutput2 & CHR(10) & CHR(10) & CHR(10) & CHR(10) & CHR(10) & CHR(10) & CHR(10) & CHR(10) & CHR(10)
							lsOutput = lsOutput & CHR(27) & CHR(97) & CHR(0)
							' Cut Paper
							lsOutput = lsOutput & CHR(29) & CHR(86) & CHR(1)
							lsOutput2 = lsOutput2 & CHR(29) & CHR(86) & CHR(1)
							' Ring bell
	'						lsOutput = lsOutput & CHR(27) & CHR(40) & CHR(65) & CHR(4) & CHR(0) & CHR(48) & CHR(56) & CHR(2) & CHR(10)
	'						lsOutput2 = lsOutput2 & CHR(27) & CHR(40) & CHR(65) & CHR(4) & CHR(0) & CHR(48) & CHR(56) & CHR(2) & CHR(10)
							' Drawer kick to trigger external bell
							lsOutput = lsOutput & CHR(27) & CHR(112) & CHR(0) & CHR(60) & CHR(60)
							lsOutput2 = lsOutput2 & CHR(27) & CHR(112) & CHR(0) & CHR(60) & CHR(60)
							
							' Print everything to make lines
							If GetStoreMakeLinePrinters(pnStoreID, lasIPAddresses) Then
								If lasIPAddresses(0) <> "" Then
									For i = 0 To UBound(lasIPAddresses)
										If Not SendToPrinter(lasIPAddresses(i), lsOutput) Then
											lbRet = FALSE
										End If
									Next
								End If
							End If
							
							' Print to alternate make lines
							If lbDoAltMakeLine Then
								If GetStoreAltMakeLinePrinters(pnStoreID, lasIPAddresses) Then
									If lasIPAddresses(0) <> "" Then
										For i = 0 To UBound(lasIPAddresses)
' 2013-10-30 - Do not fail if alt make line printing fails
'											If Not SendToPrinter(lasIPAddresses(i), lsOutput2) Then
'												lbRet = FALSE
'											End If
											SendToPrinter lasIPAddresses(i), lsOutput2
										Next
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
	End If
	
	PrintOrder = lbRet
End Function

' **************************************************************************
' Function: GetStoreMakeLinePrinters
' Purpose: Retrieves all of the printers designated as make line printers
' Parameters:	pnStoreID - StoreID to search for
'				pnIPAddresses - Array of IP addresses of make line printeres
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStoreMakeLinePrinters(ByVal pnStoreID, ByRef pasPrinterIPAddresses)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select distinct PrinterIPAddress from tblPrinters where StoreID = " & pnStoreID & " and PrinterTypeID = 1"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve pasPrinterIPAddresses(lnPos)
				
				pasPrinterIPAddresses(lnPos) = loRS("PrinterIPAddress")
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim pasPrinterIPAddresses(0)
			pasPrinterIPAddresses(0) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStoreMakeLinePrinters = lbRet
End Function

' **************************************************************************
' Function: GetStoreAltMakeLinePrinters
' Purpose: Retrieves all of the printers designated as alternate make line printers
' Parameters:	pnStoreID - StoreID to search for
'				pnIPAddresses - Array of IP addresses of make line printeres
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStoreAltMakeLinePrinters(ByVal pnStoreID, ByRef pasPrinterIPAddresses)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select distinct PrinterIPAddress from tblPrinters where StoreID = " & pnStoreID & " and PrinterTypeID = 2"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve pasPrinterIPAddresses(lnPos)
				
				pasPrinterIPAddresses(lnPos) = loRS("PrinterIPAddress")
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim pasPrinterIPAddresses(0)
			pasPrinterIPAddresses(0) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStoreAltMakeLinePrinters = lbRet
End Function

' **************************************************************************
' Function: GetPrintStoreHeader
' Purpose: Retrieves a string for printing the store header
' Parameters:	pnStoreID - StoreID to search for
' Return: True if sucessful, False if not
' **************************************************************************
Function GetPrintStoreHeader(ByVal pnStoreID)
	Dim lsRet, lsSQL, loRS
	
	lsRet = ""
	
	lsSQL = "select StoreName from tblStores where StoreID = " & pnStoreID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lsRet = loRS("StoreName") &  " #" & pnStoreID
		End If
		
		DBCloseQuery loRS
	End If
	
	GetPrintStoreHeader = lsRet
End Function

' **************************************************************************
' Function: IsUnitAltPrinter
' Purpose: Determines if a unit has an alternate printer assignment
' Parameters:	pnStoreID - StoreID to search for
'				pnUnitID - UnitID to search for
' Return: True if there is, False if not
' **************************************************************************
Function IsUnitAltPrinter(ByVal pnStoreID, ByVal pnUnitID)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	lsSQL = "select * from trelStoreUnitAltPrinters where StoreID = " & pnStoreID & " and UnitID = " & pnUnitID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lbRet = TRUE
		End If
		
		DBCloseQuery loRS
	End If
	
	IsUnitAltPrinter = lbRet
End Function

' **************************************************************************
' Function: PrintSignatureCopies
' Purpose: Prints signature copies to the appropriate printers.
' Parameters:	pnStoreUD - The StoreID
'				pnOrderID - The OrderID
'				psIPAddress - The station IP address (only if not delivery)
' Return: True if sucessful, False if not
' **************************************************************************
Function PrintSignatureCopies(ByVal pnStoreID, ByVal pnOrderID, ByVal psIPAddress)
	Dim lbRet, i, j, k, lnTmp, lasIPAddresses(), lsOutput, lsOutput2, ldTotalCost, ldTotalDiscount, lsOrderTypeDescription
	Dim lnSessionID, lsIPAddress, lnEmpID, lsRefID, ldtTransactionDate, ldtSubmitDate, ldtReleaseDate, ldtExpectedDate, lnStoreID, lnCustomerID, lsCustomerName, lsCustomerPhone, lnAddressID, lnOrderTypeID, lbIsPaid, lnPaymentTypeID, lsPaymentReference, lnAccountID, ldDeliveryCharge, ldDriverMoney, ldTax, ldTax2, ldTip, lnOrderStatusID, lsOrderNotes
	Dim lanOrderLineIDs(), lasOrderLineDescriptions(), lanQuantity(), ladCost(), ladDiscount()
	Dim lnUnitID, lnSpecialtyID, lnSizeID, lnStyleID, lnHalf1SauceID, lnHalf2SauceID, lnHalf1SauceModifierID, lnHalf2SauceModifierID, lsOrderLineNotes, lnQuantity, ldOrderLineCost, ldOrderLineDiscount, lnCouponID
	Dim lanOrderLineItemIDs(), lanItemIDs(), lanHalfIDs(), lasItemDescriptions(), lasItemShortDescriptions()
	Dim lanOrderLineTopperIDs(), lanTopperIDs(), lanTopperHalfIDs(), lasTopperDescriptions(), lasTopperShortDescriptions()
	Dim lanOrderLineFreeSideIDs(), lanFreeSideIDs(), lasFreeSideDescriptions(), lasFreeSideShortDescriptions()
	Dim lanOrderLineAddSideIDs(), lanAddSideIDs(), lasAddSideDescriptions(), lasAddSideShortDescriptions()
	Dim lanUnitSideIDs(), lasUnitSideDescriptions(), lasUnitSideShortDescriptions(), ladUnitSidePrices()
	Dim lanUnitItemIDs(), lasUnitItemDescriptions(), lasUnitItemShortDescriptions(), ladUnitItemOnSidePrice(), lanUnitItemCounts(), labUnitFreeItemFlags(), labUnitItemIsCheeses(), labUnitItemIsBaseCheeses(), labUnitItemIsExtraCheeses()
	Dim lanWeightItemIDs(), lanItemCounts1(), ladItemWeights1(), lanItemCounts2(), ladItemWeights2(), ladItemWeights3()
	Dim lanWeightSauceIDs(), lanWeightSauceModifierIDs(), ladSauceWeights()
	Dim lsAddress1, lsAddress2, lsCity, lsState, lsPostalCode, lsAddressNotes
	Dim lsAddressDescription, lsCustomerNotes
	Dim lnItemCount, lbHalveBaseCheese, lbIsBaseCheese, lnExtraCheeseItemID, ldBaseCheeseWeight, lsWeight
	Dim lsPrinterIPAddress, lsCoupons
	Dim lsAccountName
	Dim lsBaseCheeseAfter
	
	lbRet = FALSE
	
	If GetOrderDetails(pnOrderID, lnSessionID, lsIPAddress, lnEmpID, lsRefID, ldtTransactionDate, ldtSubmitDate, ldtReleaseDate, ldtExpectedDate, lnStoreID, lnCustomerID, lsCustomerName, lsCustomerPhone, lnAddressID, lnOrderTypeID, lbIsPaid, lnPaymentTypeID, lsPaymentReference, lnAccountID, ldDeliveryCharge, ldDriverMoney, ldTax, ldTax2, ldTip, lnOrderStatusID, lsOrderNotes) Then
		If GetAddressDetails(lnAddressID, lnTmp, lsAddress1, lsAddress2, lsCity, lsState, lsPostalCode, lsAddressNotes) Then
			lsAddressDescription = ""
			lsCustomerNotes = ""
			
			If lnAddressID > 1 Then
				GetCustomerAddressDetails lnCustomerID, lnAddressID, lsAddressDescription, lsCustomerNotes
			End If
			
			If GetOrderLines(pnOrderID, lanOrderLineIDs, lasOrderLineDescriptions, lanQuantity, ladCost, ladDiscount) Then
				If GetOrderCoupons(pnOrderID, lsCoupons) Then
					lbRet = TRUE
					
					' Build for customer and store copies
					ldTotalCost = 0.00
					ldTotalDiscount = 0.00
					lsOrderTypeDescription = GetOrderTypeDescription(lnOrderTypeID)
					
					' Start a new line
					lsOutput = CHR(10)
					
					' Add logo
					lsOutput = lsOutput & CHR(27) & CHR(97) & CHR(1) & CHR(28) & CHR(112) & CHR(1) & CHR(0) & CHR(27) & CHR(97) & CHR(0)
					
					' Set font size to 12
					lsOutput = lsOutput & CHR(10) & CHR(27) & CHR(33) & CHR(12)
					
					' Add header with center justification
					lsOutput = lsOutput & CHR(27) & CHR(97) & CHR(1)
					lsOutput = lsOutput & "vitos.com #" & pnStoreID & " " & GetEmployeeShortName(lnEmpID) & CHR(10)
					lsOutput = lsOutput & FormatDateTime(Now()) & " #" & pnOrderID & CHR(10)
					lsOutput = lsOutput & CHR(10)
					
					' Return to left justify and add line items to customer copy
					lsOutput = lsOutput & CHR(27) & CHR(97) & CHR(0)
					lsOutput2 = lsOutput
					For i = 0 To UBound(lanOrderLineIDs)
						If GetOrderLineDetails(lanOrderLineIDs(i), lnUnitID, lnSpecialtyID, lnSizeID, lnStyleID, lnHalf1SauceID, lnHalf2SauceID, lnHalf1SauceModifierID, lnHalf2SauceModifierID, lsOrderLineNotes, lnQuantity, ldOrderLineCost, ldOrderLineDiscount, lnCouponID) Then
							If GetOrderLineItems(lanOrderLineIDs(i), lnUnitID, lanOrderLineItemIDs, lanItemIDs, lanHalfIDs, lasItemDescriptions, lasItemShortDescriptions) Then
								If GetOrderLineToppers(lanOrderLineIDs(i), lanOrderLineTopperIDs, lanTopperIDs, lanTopperHalfIDs, lasTopperDescriptions, lasTopperShortDescriptions) Then
									If GetOrderLineFreeSides(lanOrderLineIDs(i), lanOrderLineFreeSideIDs, lanFreeSideIDs, lasFreeSideDescriptions, lasFreeSideShortDescriptions) Then
										If GetOrderLineAddSides(lanOrderLineIDs(i), lanOrderLineAddSideIDs, lanAddSideIDs, lasAddSideDescriptions, lasAddSideShortDescriptions) Then
											If GetStoreUnitSides(pnStoreID, lnUnitID, lanUnitSideIDs, lasUnitSideDescriptions, lasUnitSideShortDescriptions, ladUnitSidePrices) Then
												If GetStoreUnitItems(pnStoreID, lnUnitID, lanUnitItemIDs, lasUnitItemDescriptions, lasUnitItemShortDescriptions, ladUnitItemOnSidePrice, lanUnitItemCounts, labUnitFreeItemFlags, labUnitItemIsCheeses, labUnitItemIsBaseCheeses, labUnitItemIsExtraCheeses) Then
													If GetUnitSizeItemWeights(lnUnitID, lnSizeID, lanWeightItemIDs, lanItemCounts1, ladItemWeights1, lanItemCounts2, ladItemWeights2, ladItemWeights3) Then
														If GetUnitSizeSauceWeights(lnUnitID, lnSizeID, lanWeightSauceIDs, lanWeightSauceModifierIDs, ladSauceWeights) Then
															ldTotalCost = ldTotalCost + (lnQuantity * ldOrderLineCost)
															ldTotalDiscount = ldTotalDiscount + (lnQuantity * ldOrderLineDiscount)
															
															If lnSizeID > 0 Then
																lsOutput = lsOutput & GetSizeShortDescription(lnSizeID) & " "
															End If
															lsOutput = lsOutput & GetUnitShortDescription(lnUnitID) & " "
															If lnQuantity > 1 Then
																lsOutput = lsOutput & lnQuantity & " @ "
															End If
															lsOutput = lsOutput & FormatCurrency(ldOrderLineCost - ldOrderLineDiscount) & CHR(10)
															
															If lnSpecialtyID > 0 Then
																lsOutput = lsOutput & GetSpecialtyShortDescription(lnSpecialtyID) & CHR(10)
															End If
															If lnStyleID > 0 Then
																lsOutput = lsOutput & GetStyleShortDescription(lnStyleID) & CHR(10)
															End If
															
															If lnUnitID = 1 Then
																If lnHalf1SauceID > 0 Then
																	lsWeight = "     "
																	For k = 0 To UBound(lanWeightSauceIDs)
																		If lnHalf1SauceID = lanWeightSauceIDs(k) And lnHalf1SauceModifierID = lanWeightSauceModifierIDs(k) Then
																			If (lnHalf1SauceID = lnHalf2SauceID) And (lnHalf1SauceModifierID = lnHalf2SauceModifierID) Then
																				lsWeight = FormatCurrency(ladSauceWeights(k))
																			Else
																				lsWeight = FormatCurrency(ladSauceWeights(k) / 2)
																			End If
																			lsWeight = Right(lsWeight, (Len(lsWeight) - 1))
																			lsWeight = Right("     " & lsWeight, 5)
																			
																			Exit For
																		End If
																	Next
																	
																	If (lnHalf1SauceID = lnHalf2SauceID) And (lnHalf1SauceModifierID = lnHalf2SauceModifierID) Then
																		If lnHalf1SauceModifierID > 0 Then
																			lsOutput = lsOutput & Left(GetSauceShortDescription(lnHalf1SauceID) & " " & GetSauceModifierShortDescription(lnHalf1SauceModifierID) & String(27, " "), 27)
																		Else
																			lsOutput = lsOutput & Left(GetSauceShortDescription(lnHalf1SauceID) & String(27, " "), 27)
																		End If
																	Else
																		If lnHalf1SauceModifierID > 0 Then
																			lsOutput = lsOutput & "HALF 1: " & Left(GetSauceShortDescription(lnHalf1SauceID) & " " & GetSauceModifierShortDescription(lnHalf1SauceModifierID) & String(19, " "), 19)
																		Else
																			lsOutput = lsOutput & "HALF 1: " & Left(GetSauceShortDescription(lnHalf1SauceID) & String(19, " "), 19)
																		End If
																	End If
																	lsOutput = lsOutput & CHR(10)
																End If
																If lnHalf2SauceID > 0 And ((lnHalf1SauceID <> lnHalf2SauceID) Or (lnHalf1SauceModifierID <> lnHalf2SauceModifierID)) Then
																	lsWeight = "     "
																	For k = 0 To UBound(lanWeightSauceIDs)
																		If lnHalf2SauceID = lanWeightSauceIDs(k) And lnHalf2SauceModifierID = lanWeightSauceModifierIDs(k) Then
																			lsWeight = FormatCurrency(ladSauceWeights(k) / 2)
																			lsWeight = Right(lsWeight, (Len(lsWeight) - 1))
																			lsWeight = Right("     " & lsWeight, 5)
																			
																			Exit For
																		End If
																	Next
																	
																	If lnHalf2SauceModifierID > 0 Then
																		lsOutput = lsOutput & "HALF 2: " & Left(GetSauceShortDescription(lnHalf2SauceID) & " " & GetSauceModifierShortDescription(lnHalf2SauceModifierID) & String(19, " "), 19)
																	Else
																		lsOutput = lsOutput & "HALF 2: " & Left(GetSauceShortDescription(lnHalf2SauceID) & String(19, " "), 19)
																	End If
																	lsOutput = lsOutput & CHR(10)
																End If
															End If
															
															lnItemCount = 0
															lbHalveBaseCheese = FALSE
															For j = 0 To UBound(lanItemIDs)
																If lanItemIDs(j) > 0 Then
																	For k = 0 To UBound(lanUnitItemIDs)
																		If lanItemIDs(j) = lanUnitItemIDs(k) Then
																			If (Not labUnitFreeItemFlags(k)) And (Not labUnitItemIsCheeses(k)) Then
																				Select Case lanHalfIDs(j)
																					Case 0:
																						lnItemCount = lnItemCount + lanUnitItemCounts(k)
																					Case 1:
																						lnItemCount = lnItemCount + (lanUnitItemCounts(k) / 2)
																					Case 2:
																						lnItemCount = lnItemCount + (lanUnitItemCounts(k) / 2)
																				End Select
																			End If
																			
																			If labUnitItemIsCheeses(k) And (Not labUnitItemIsBaseCheeses(k)) And (Not labUnitItemIsExtraCheeses(k)) Then
																				lbHalveBaseCheese = TRUE
																			End If
																			
																			Exit For
																		End If
																	Next
																End If
															Next
															
															For j = 0 To UBound(lanTopperIDs)
																If lanTopperIDs(j) > 0 Then
																	If IsTopperBeforeItems(lnUnitID, lanTopperIDs(j)) Then
																		Select Case lanTopperHalfIDs(j)
																			Case 0
																				lsOutput = lsOutput & lasTopperShortDescriptions(j) & CHR(10)
																			Case 1
																				lsOutput = lsOutput & "HALF 1: " & lasTopperShortDescriptions(j) & CHR(10)
																			Case 2
																				lsOutput = lsOutput & "HALF 2: " & lasTopperShortDescriptions(j) & CHR(10)
																			Case 3
																				lsOutput = lsOutput & "ON SIDE: " & lasTopperShortDescriptions(j) & CHR(10)
																		End Select
																	End If
																End If
															Next
															
															lsBaseCheeseAfter = ""
															For j = 0 To UBound(lanItemIDs)
																If lanItemIDs(j) > 0 Then
																	lbIsBaseCheese = FALSE
																	For k = 0 To UBound(lanUnitItemIDs)
																		If lanItemIDs(j) = lanUnitItemIDs(k) Then
																			If labUnitItemIsBaseCheeses(k) Then
																				lbIsBaseCheese = TRUE
																			End If
																			
																			Exit For
																		End If
																	Next
																	
																	If lbIsBaseCheese And UBound(lanItemIDs) = 0 Then
																		lnExtraCheeseItemID = 0
																		For k = 0 To UBound(lanUnitItemIDs)
																			If labUnitItemIsExtraCheeses(k) Then
																				lnExtraCheeseItemID = lanUnitItemIDs(k)
																				lbIsBaseCheese = TRUE
																				
																				Exit For
																			End If
																		Next
																		
																		ldBaseCheeseWeight = 0.00
																		For k = 0 To UBound(lanWeightItemIDs)
																			If lanItemIDs(j) = lanWeightItemIDs(k) Then
																				ldBaseCheeseWeight = ladItemWeights1(k)
																				
																				Exit For
																			End If
																		Next
																		
																		If lnExtraCheeseItemID > 0 Then
																			For k = 0 To UBound(lanWeightItemIDs)
																				If lnExtraCheeseItemID = lanWeightItemIDs(k) Then
																					ldBaseCheeseWeight = ldBaseCheeseWeight + ladItemWeights1(k)
																					
																					Exit For
																				End If
																			Next
																		End If
																		
																		lsWeight = FormatCurrency(ldBaseCheeseWeight)
																		lsWeight = Right(lsWeight, (Len(lsWeight) - 1))
																		lsWeight = Right("     " & lsWeight, 5)
																	Else
																		lsWeight = "     "
																		For k = 0 To UBound(lanWeightItemIDs)
																			If lanItemIDs(j) = lanWeightItemIDs(k) Then
																				If lbIsBaseCheese Then
																					If lbHalveBaseCheese Then
																						lsWeight = FormatCurrency(ladItemWeights1(k) / 2)
																					Else
																						lsWeight = FormatCurrency(ladItemWeights1(k))
																					End If
																				Else
																					If lnItemCount <= lanItemCounts1(k) Then
																						Select Case lanHalfIDs(j)
																							Case 0
																								lsWeight = FormatCurrency(ladItemWeights1(k))
																							Case 1, 2
																								lsWeight = FormatCurrency(ladItemWeights1(k) / 2)
																						End Select
																					Else
																						If lnItemCount <= lanItemCounts2(k) Then
																							Select Case lanHalfIDs(j)
																								Case 0
																									lsWeight = FormatCurrency(ladItemWeights2(k))
																								Case 1, 2
																									lsWeight = FormatCurrency(ladItemWeights2(k) / 2)
																							End Select
																						Else
																							Select Case lanHalfIDs(j)
																								Case 0
																									lsWeight = FormatCurrency(ladItemWeights3(k))
																								Case 1, 2
																									lsWeight = FormatCurrency(ladItemWeights3(k) / 2)
																							End Select
																						End If
																					End If
																				End If
																				lsWeight = Right(lsWeight, (Len(lsWeight) - 1))
																				lsWeight = Right("     " & lsWeight, 5)
																				
																				Exit For
																			End If
																		Next
																	End If
																	
'																	If lbIsBaseCheese Then
'																		Select Case lanHalfIDs(j)
'																			Case 0
'																				lsBaseCheeseAfter = Left(lasItemShortDescriptions(j) & String(27, " "), 27) & CHR(10)
'																			Case 1
'																				lsBaseCheeseAfter = "HALF 1:" & Left(lasItemShortDescriptions(j) & String(19, " "), 19) & CHR(10)
'																			Case 2
'																				lsBaseCheeseAfter = "HALF 2:" & Left(lasItemShortDescriptions(j) & String(19, " "), 19) & CHR(10)
'																			Case 3
'																				lsBaseCheeseAfter = "ON SIDE:" & lasItemShortDescriptions(j) & CHR(10)
'																		End Select
'																	Else
																		Select Case lanHalfIDs(j)
																			Case 0
																				lsOutput = lsOutput & Left(lasItemShortDescriptions(j) & String(27, " "), 27) & CHR(10)
																			Case 1
																				lsOutput = lsOutput & "HALF 1: " & Left(lasItemShortDescriptions(j) & String(19, " "), 19) & CHR(10)
																			Case 2
																				lsOutput = lsOutput & "HALF 2: " & Left(lasItemShortDescriptions(j) & String(19, " "), 19) & CHR(10)
																			Case 3
																				lsOutput = lsOutput & "ON SIDE: " & lasItemShortDescriptions(j) & CHR(10)
																		End Select
'																	End If
																End If
															Next
															lsOutput = lsOutput & lsBaseCheeseAfter
															
															If lnUnitID <> 1 Then
																If lnHalf1SauceID > 0 Then
																	lsWeight = "     "
																	For k = 0 To UBound(lanWeightSauceIDs)
																		If lnHalf1SauceID = lanWeightSauceIDs(k) And lnHalf1SauceModifierID = lanWeightSauceModifierIDs(k) Then
																			If (lnHalf1SauceID = lnHalf2SauceID) And (lnHalf1SauceModifierID = lnHalf2SauceModifierID) Then
																				lsWeight = FormatCurrency(ladSauceWeights(k))
																			Else
																				lsWeight = FormatCurrency(ladSauceWeights(k) / 2)
																			End If
																			lsWeight = Right(lsWeight, (Len(lsWeight) - 1))
																			lsWeight = Right("     " & lsWeight, 5)
																			
																			Exit For
																		End If
																	Next
																	
																	If (lnHalf1SauceID = lnHalf2SauceID) And (lnHalf1SauceModifierID = lnHalf2SauceModifierID) Then
																		If lnHalf1SauceModifierID > 0 Then
																			lsOutput = lsOutput & Left(GetSauceShortDescription(lnHalf1SauceID) & " " & GetSauceModifierShortDescription(lnHalf1SauceModifierID) & String(27, " "), 27)
																		Else
																			lsOutput = lsOutput & Left(GetSauceShortDescription(lnHalf1SauceID) & String(27, " "), 27)
																		End If
																	Else
																		If lnHalf1SauceModifierID > 0 Then
																			lsOutput = lsOutput & "HALF 1: " & Left(GetSauceShortDescription(lnHalf1SauceID) & " " & GetSauceModifierShortDescription(lnHalf1SauceModifierID) & String(19, " "), 19)
																		Else
																			lsOutput = lsOutput & "HALF 1: " & Left(GetSauceShortDescription(lnHalf1SauceID) & String(19, " "), 19)
																		End If
																	End If
																	lsOutput = lsOutput & CHR(10)
																End If
																If lnHalf2SauceID > 0 And ((lnHalf1SauceID <> lnHalf2SauceID) Or (lnHalf1SauceModifierID <> lnHalf2SauceModifierID)) Then
																	lsWeight = "     "
																	For k = 0 To UBound(lanWeightSauceIDs)
																		If lnHalf2SauceID = lanWeightSauceIDs(k) And lnHalf2SauceModifierID = lanWeightSauceModifierIDs(k) Then
																			lsWeight = FormatCurrency(ladSauceWeights(k) / 2)
																			lsWeight = Right(lsWeight, (Len(lsWeight) - 1))
																			lsWeight = Right("     " & lsWeight, 5)
																			
																			Exit For
																		End If
																	Next
																	
																	If lnHalf2SauceModifierID > 0 Then
																		lsOutput = lsOutput & "HALF 2: " & Left(GetSauceShortDescription(lnHalf2SauceID) & " " & GetSauceModifierShortDescription(lnHalf2SauceModifierID) & String(19, " "), 19)
																	Else
																		lsOutput = lsOutput & "HALF 2: " & Left(GetSauceShortDescription(lnHalf2SauceID) & String(19, " "), 19)
																	End If
																	lsOutput = lsOutput & CHR(10)
																End If
															End If
															
															For j = 0 To UBound(lanTopperIDs)
																If lanTopperIDs(j) > 0 Then
																	If Not IsTopperBeforeItems(lnUnitID, lanTopperIDs(j)) Then
																		Select Case lanTopperHalfIDs(j)
																			Case 0
																				lsOutput = lsOutput & lasTopperShortDescriptions(j) & CHR(10)
																			Case 1
																				lsOutput = lsOutput & "HALF 1: " & lasTopperShortDescriptions(j) & CHR(10)
																			Case 2
																				lsOutput = lsOutput & "HALF 2: " & lasTopperShortDescriptions(j) & CHR(10)
																			Case 3
																				lsOutput = lsOutput & "ON SIDE: " & lasTopperShortDescriptions(j) & CHR(10)
																		End Select
																	End If
																End If
															Next
															
															For j = 0 To UBound(lanFreeSideIDs)
																If lanFreeSideIDs(j) > 0 Then
																	lsOutput = lsOutput & "    " & Left(lasFreeSideShortDescriptions(j) & String(23, " "), 23) & " (N/C)" & CHR(10)
																End If
															Next
															For j = 0 To UBound(lanAddSideIDs)
																If lanAddSideIDs(j) > 0 Then
																	For k = 0 To UBound(lanUnitSideIDs)
																		If lanUnitSideIDs(k) = lanAddSideIDs(j) Then
																			lsOutput = lsOutput & "    " & Left(lasAddSideShortDescriptions(j) & String(23, " "), 23) & " " & FormatCurrency(ladUnitSidePrices(k)) & CHR(10)
																			
																			Exit For
																		End If
																	Next
																End If
															Next
															
															If Len(lsOrderLineNotes) > 0 Then
																lsOutput = lsOutput & "    " & lsOrderLineNotes & CHR(10)
															End If
															
															lsOutput = lsOutput & CHR(10)
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
					
					If lbRet Then
						' Add totals
						lsOutput = lsOutput & "This sale:            " & Replace(FormatCurrency(ldTotalCost), "$", "$" & String(11 - Len(FormatCurrency(ldTotalCost)), " ")) & CHR(10)
						lsOutput2 = lsOutput2 & "This sale:            " & Replace(FormatCurrency(ldTotalCost), "$", "$" & String(11 - Len(FormatCurrency(ldTotalCost)), " ")) & CHR(10)
						lsOutput = lsOutput & "Tax 1:                " & Replace(FormatCurrency(ldTax), "$", "$" & String(11 - Len(FormatCurrency(ldTax)), " ")) & CHR(10)
						lsOutput2 = lsOutput2 & "Tax 1:                " & Replace(FormatCurrency(ldTax), "$", "$" & String(11 - Len(FormatCurrency(ldTax)), " ")) & CHR(10)
						If ldTax2 <> 0 Then
							lsOutput = lsOutput & "Tax 2:                " & Replace(FormatCurrency(ldTax2), "$", "$" & String(11 - Len(FormatCurrency(ldTax2)), " ")) & CHR(10)
							lsOutput2 = lsOutput2 & "Tax 2:                " & Replace(FormatCurrency(ldTax2), "$", "$" & String(11 - Len(FormatCurrency(ldTax2)), " ")) & CHR(10)
						End If
						If ldTotalDiscount <> 0 Then
							lsOutput = lsOutput & "Discount:             " & Replace(FormatCurrency(ldTotalDiscount), "$", "$" & String(11 - Len(FormatCurrency(ldTotalDiscount)), " ")) & CHR(10)
							lsOutput2 = lsOutput2 & "Discount:             " & Replace(FormatCurrency(ldTotalDiscount), "$", "$" & String(11 - Len(FormatCurrency(ldTotalDiscount)), " ")) & CHR(10)
						End If
						If lnOrderTypeID = 1 Then
							lsOutput = lsOutput & Left(lsOrderTypeDescription & ":" & String(22, " "), 22) & Replace(FormatCurrency(ldDeliveryCharge), "$", "$" & String(11 - Len(FormatCurrency(ldDeliveryCharge)), " ")) & CHR(10)
							lsOutput2 = lsOutput2 & Left(lsOrderTypeDescription & ":" & String(22, " "), 22) & Replace(FormatCurrency(ldDeliveryCharge), "$", "$" & String(11 - Len(FormatCurrency(ldDeliveryCharge)), " ")) & CHR(10)
						End If
						lsOutput = lsOutput & String(42, "-") & CHR(10)
						lsOutput2 = lsOutput2 & String(42, "-") & CHR(10)
						lsOutput = lsOutput & "Ticket Total:         " & Replace(FormatCurrency((ldTotalCost + ldTax + ldTax2 + ldDeliveryCharge) - ldTotalDiscount), "$", "$" & String(11 - Len(FormatCurrency((ldTotalCost + ldTax + ldTax2 + ldDeliveryCharge) - ldTotalDiscount)), " ")) & CHR(10)
						lsOutput2 = lsOutput2 & "Ticket Total:         " & Replace(FormatCurrency((ldTotalCost + ldTax + ldTax2 + ldDeliveryCharge) - ldTotalDiscount), "$", "$" & String(11 - Len(FormatCurrency((ldTotalCost + ldTax + ldTax2 + ldDeliveryCharge) - ldTotalDiscount)), " ")) & CHR(10)
						lsOutput = lsOutput & String(42, "-") & CHR(10)
						lsOutput2 = lsOutput2 & String(42, "-") & CHR(10)
						
						If ldTip <> 0 Then
							lsOutput = lsOutput & "Tip:                  " & Replace(FormatCurrency(ldTip), "$", "$" & String(11 - Len(FormatCurrency(ldTip)), " ")) & CHR(10)
							lsOutput2 = lsOutput2 & "Tip:                  " & Replace(FormatCurrency(ldTip), "$", "$" & String(11 - Len(FormatCurrency(ldTip)), " ")) & CHR(10)
							lsOutput = lsOutput & "Grand Total:          " & Replace(FormatCurrency((ldTotalCost + ldTax + ldTax2 + ldDeliveryCharge + ldTip) - ldTotalDiscount), "$", "$" & String(11 - Len(FormatCurrency((ldTotalCost + ldTax + ldTax2 + ldDeliveryCharge + ldTip) - ldTotalDiscount)), " ")) & CHR(10)
							lsOutput2 = lsOutput2 & "Grand Total:          " & Replace(FormatCurrency((ldTotalCost + ldTax + ldTax2 + ldDeliveryCharge + ldTip) - ldTotalDiscount), "$", "$" & String(11 - Len(FormatCurrency((ldTotalCost + ldTax + ldTax2 + ldDeliveryCharge + ldTip) - ldTotalDiscount)), " ")) & CHR(10)
						End If
						
						If Len(lsCoupons) > 0 Then
							lsOutput = lsOutput & lsCoupons & CHR(10)
							lsOutput = lsOutput & CHR(10)
						End If
						
						If ldTip = 0 Then
							lsOutput = lsOutput & "Tip:                 ____________" & CHR(10) & CHR(10)
							lsOutput2 = lsOutput2 & "Tip:                 ____________" & CHR(10) & CHR(10)
							lsOutput = lsOutput & "Grand Total:         ____________" & CHR(10) & CHR(10)
							lsOutput2 = lsOutput2 & "Grand Total:         ____________" & CHR(10) & CHR(10)
						End If
						lsOutput = lsOutput & CHR(10)
						lsOutput2 = lsOutput2 & CHR(10)
						
						' Signature Line
						lsOutput = lsOutput & "Signature: ________________________" & CHR(10)
						lsOutput2 = lsOutput2 & "Signature: ________________________" & CHR(10)
						If lnPaymentTypeID = 3 Then
							lsOutput = lsOutput & "      I agree to above total amount" & CHR(10)
							lsOutput2 = lsOutput2 & "      I agree to above total amount" & CHR(10)
							lsOutput = lsOutput & "      as per card issuer agreement." & CHR(10)
							lsOutput2 = lsOutput2 & "      as per card issuer agreement." & CHR(10)
						End If
						
						' Indicate payment reference
						Select Case lnPaymentTypeID
							Case 3
								lsOutput = lsOutput & "Credit Card Auth #: " & lsPaymentReference & CHR(10)
								lsOutput2 = lsOutput2 & "Credit Card Auth #: " & lsPaymentReference & CHR(10)
							Case 4
								lsAccountName = GetAccountName(lnAccountID)
								lsOutput = lsOutput & "Placed On Account #" & lnAccountID & " " & lsAccountName & CHR(10)
								lsOutput2 = lsOutput2 & "Placed On Account #" & lnAccountID & " " & lsAccountName & CHR(10)
						End Select
						lsOutput = lsOutput & CHR(10)
						lsOutput2 = lsOutput2 & CHR(10)
						
						' Indicate pickup/delivery
						lsOutput = lsOutput & lsOrderTypeDescription & CHR(10)
						lsOutput2 = lsOutput2 & lsOrderTypeDescription & CHR(10)
						
						' Add customer and address information
						lsOutput = lsOutput & lsCustomerName & CHR(10)
						lsOutput2 = lsOutput2 & lsCustomerName & CHR(10)
						Select Case lnOrderTypeID
							Case 1
								If Len(lsAddress2) = 0 Then
									If Len(lsAddress1) > 0 Then
										lsOutput = lsOutput & lsAddress1 & CHR(10)
										lsOutput2 = lsOutput2 & lsAddress1 & CHR(10)
									End If
								Else
									lsOutput = lsOutput & lsAddress1 & " #" & lsAddress2 & CHR(10)
									lsOutput2 = lsOutput2 & lsAddress1 & " #" & lsAddress2 & CHR(10)
								End If
								
								If Len(lsCity) > 0 Or Len(lsState) > 0 Or Len(lsPostalCode) > 0 Then
									lsOutput = lsOutput & lsCity & ", " & lsState & " " & lsPostalCode & CHR(10)
									lsOutput2 = lsOutput2 & lsCity & ", " & lsState & " " & lsPostalCode & CHR(10)
								End If
								
								If Len(lsCustomerPhone) > 0 Then
									If Len(lsCustomerPhone) <= 7 Then
										lsOutput = lsOutput & Left(lsCustomerPhone, 3) & "-" & Mid(lsCustomerPhone, 4) & CHR(10)
										lsOutput2 = lsOutput2 & Left(lsCustomerPhone, 3) & "-" & Mid(lsCustomerPhone, 4) & CHR(10)
									Else
										lsOutput = lsOutput & "(" & Left(lsCustomerPhone, 3) & ") " & Mid(lsCustomerPhone, 4, 3) & "-" & Mid(lsCustomerPhone, 7) & CHR(10)
										lsOutput2 = lsOutput2 & "(" & Left(lsCustomerPhone, 3) & ") " & Mid(lsCustomerPhone, 4, 3) & "-" & Mid(lsCustomerPhone, 7) & CHR(10)
									End If
								End If
								
								If Len(lsAddressNotes) > 0 Then
									lsOutput = lsOutput & lsAddressNotes & CHR(10)
								End If
								
								If Len(lsCustomerNotes) > 0 Then
									lsOutput = lsOutput & lsCustomerNotes & CHR(10)
								End If
							Case 2
								If Len(lsCustomerPhone) > 0 Then
									If Len(lsCustomerPhone) <= 7 Then
										lsOutput = lsOutput & Left(lsCustomerPhone, 3) & "-" & Mid(lsCustomerPhone, 4) & CHR(10)
										lsOutput2 = lsOutput2 & Left(lsCustomerPhone, 3) & "-" & Mid(lsCustomerPhone, 4) & CHR(10)
									Else
										lsOutput = lsOutput & "(" & Left(lsCustomerPhone, 3) & ") " & Mid(lsCustomerPhone, 4, 3) & "-" & Mid(lsCustomerPhone, 7) & CHR(10)
										lsOutput2 = lsOutput2 & "(" & Left(lsCustomerPhone, 3) & ") " & Mid(lsCustomerPhone, 4, 3) & "-" & Mid(lsCustomerPhone, 7) & CHR(10)
									End If
								End If
						End Select
						lsOutput = lsOutput & CHR(10)
						lsOutput2 = lsOutput2 & CHR(10)
						
						If Len(lsOrderNotes) > 0 Then
							lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(1)
							lsOutput = lsOutput & lsOrderNotes & CHR(10)
							lsOutput = lsOutput & CHR(29) & CHR(66) & CHR(0)
							lsOutput = lsOutput & CHR(10)
						End If
						
						' Finish
						lsOutput = lsOutput & CHR(27) & CHR(97) & CHR(1)
						lsOutput2 = lsOutput2 & CHR(27) & CHR(97) & CHR(1)
						lsOutput = lsOutput & CHR(10) & "Find Us on Facebook" & CHR(10)
						lsOutput = lsOutput & "vitos.com/facebook" & CHR(10)
						lsOutput = lsOutput & CHR(10) & "*** CUSTOMER COPY ***" & CHR(10)
						lsOutput2 = lsOutput2 & CHR(10) & "*** STORE COPY ***" & CHR(10)
						lsOutput = lsOutput & CHR(10) & CHR(10) & CHR(10) & CHR(10) & CHR(10) & CHR(10) & CHR(10) & CHR(10) & CHR(10)
						lsOutput2 = lsOutput2 & CHR(10) & CHR(10) & CHR(10) & CHR(10) & CHR(10) & CHR(10) & CHR(10) & CHR(10) & CHR(10)
						lsOutput = lsOutput & CHR(27) & CHR(97) & CHR(0)
						lsOutput2 = lsOutput2 & CHR(27) & CHR(97) & CHR(0)
						' Cut Paper
						lsOutput = lsOutput & CHR(29) & CHR(86) & CHR(1)
						lsOutput2 = lsOutput2 & CHR(29) & CHR(86) & CHR(1)
						' Ring bell
'						lsOutput = lsOutput & CHR(27) & CHR(40) & CHR(65) & CHR(4) & CHR(0) & CHR(48) & CHR(56) & CHR(2) & CHR(10)
'						lsOutput2 = lsOutput2 & CHR(27) & CHR(40) & CHR(65) & CHR(4) & CHR(0) & CHR(48) & CHR(56) & CHR(2) & CHR(10)
' 2011-10-18 TAM: Do not send this for signature copies
'							' Drawer kick to trigger external bell
'							lsOutput = lsOutput & CHR(27) & CHR(112) & CHR(0) & CHR(60) & CHR(60)
'							lsOutput2 = lsOutput2 & CHR(27) & CHR(112) & CHR(0) & CHR(60) & CHR(60)
						
						If lnOrderTypeID = 1 Then
							If GetStoreCCPrinters(pnStoreID, lasIPAddresses) Then
								If lasIPAddresses(0) <> "" Then
									For i = 0 To UBound(lasIPAddresses)
										' Print customer copy
										If Not SendToPrinter(lasIPAddresses(i), lsOutput) Then
											lbRet = FALSE
										End If
										
										If lbRet Then
											' Print store copy
											If Not SendToPrinter(lasIPAddresses(i), lsOutput2) Then
												lbRet = FALSE
											End If
										End If
										
										If Not lbRet Then
											Exit For
										End If
									Next
								Else
									lbRet = FALSE
								End If
							Else
								If GetStoreMakeLinePrinters(pnStoreID, lasIPAddresses) Then
									If lasIPAddresses(0) <> "" Then
										' Print customer copy
										If Not SendToPrinter(lasIPAddresses(0), lsOutput) Then
											lbRet = FALSE
										End If
										
										If lbRet Then
											' Print store copy
											If Not SendToPrinter(lasIPAddresses(0), lsOutput2) Then
												lbRet = FALSE
											End If
										End If
									Else
										lbRet = FALSE
									End If
								Else
									lbRet = FALSE
								End If
							End If
						Else
							If GetStoreStationPrinter(pnStoreID, psIPAddress, lsPrinterIPAddress) Then
								' Print customer copy
								If Not SendToPrinter(lsPrinterIPAddress, lsOutput) Then
									lbRet = FALSE
								End If
								
								If lbRet Then
									' Print store copy
									If Not SendToPrinter(lsPrinterIPAddress, lsOutput2) Then
										lbRet = FALSE
									End If
								End If
							Else
								If GetStoreMakeLinePrinters(pnStoreID, lasIPAddresses) Then
									If lasIPAddresses(0) <> "" Then
										' Print customer copy
										If Not SendToPrinter(lasIPAddresses(0), lsOutput) Then
											lbRet = FALSE
										End If
										
										If lbRet Then
											' Print store copy
											If Not SendToPrinter(lasIPAddresses(0), lsOutput2) Then
												lbRet = FALSE
											End If
										End If
									Else
										lbRet = FALSE
									End If
								Else
									lbRet = FALSE
								End If
							End If
						End If
					End If
				End If
			End If
		End If
	End If
	
	PrintSignatureCopies = lbRet
End Function

' **************************************************************************
' Function: GetStoreCCPrinters
' Purpose: Retrieves all of the printers designated as credit card printers
' Parameters:	pnStoreID - StoreID to search for
'				pnIPAddresses - Array of IP addresses of credit card printeres
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStoreCCPrinters(ByVal pnStoreID, ByRef pasPrinterIPAddresses)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select distinct PrinterIPAddress from tblPrinters where StoreID = " & pnStoreID & " and PrinterTypeID = 3"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve pasPrinterIPAddresses(lnPos)
				
				pasPrinterIPAddresses(lnPos) = loRS("PrinterIPAddress")
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim pasPrinterIPAddresses(0)
			pasPrinterIPAddresses(0) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStoreCCPrinters = lbRet
End Function

' **************************************************************************
' Function: GetStoreStationPrinter
' Purpose: Retrieves the IP address of the printer for a store/station.
' Parameters:	pnStoreID - The StoreID to search for
'				psStationIPAddress - The station IP address to look for
'				psPrinterIPAddress - The Printer IP address
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStoreStationPrinter(ByVal pnStoreID, ByVal psStationIPAddress, ByRef psPrinterIPAddress)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	lsSQL = "select PrinterIPAddress from trelStorePrinters inner join tblPrinters on trelStorePrinters.StoreID = tblPrinters.StoreID and trelStorePrinters.PrinterID = tblPrinters.PrinterID where trelStorePrinters.StoreID = " & pnStoreID & " and StationIPAddress = '" & DBCleanLiteral(psStationIPAddress) & "' and PrinterTypeID = 4"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lbRet = TRUE
			
			psPrinterIPAddress = Trim(loRS("PrinterIPAddress"))
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStoreStationPrinter = lbRet
End Function

' **************************************************************************
' Function: GetStorePrinters
' Purpose: Retrieves the IP addresses of all printers for a store.
' Parameters:	pnStoreID - The StoreID to search for
'				pasPrinterIPAddresses - Array of Printer IP addresses found
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStorePrinters(ByVal pnStoreID, ByRef pasPrinterIPAddresses)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select distinct PrinterIPAddress from tblPrinters where StoreID = " & pnStoreID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve pasPrinterIPAddresses(lnPos)
				
				pasPrinterIPAddresses(lnPos) = loRS("PrinterIPAddress")
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim pasPrinterIPAddresses(0)
			pasPrinterIPAddresses(0) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStorePrinters = lbRet
End Function
%>