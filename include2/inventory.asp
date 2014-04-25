<%
' **************************************************************************
' File: inventory.asp
' Purpose: Functions for inventory related activities.
' Created: 7/19/2011 - TAM
' Description:
'	Include this file on any page where inventory data is manipulated.
'	This file includes the following functions: GetUnitStandards, GetUnitSizeItemWeights,
'		GetUnitSizeSauceWeights, GetUnitSizeStyleWeights, GetUnitSizeTopperWeights,
'		GetUnitSizeSideWeights, GetStoreInventory, GetStoreOnhandInventory, GetStoreOnhandInventoryVariance,
'		InitializeOnhandInventory, SetOnhandInventory, SetOnhandInventoryRecount, LockOnhandInventory, GetStoreVendors,
'		GetVendorOrderInventory, CreateVendorOrder, CountVendorWeekOrders,
'		GetStoreStockLevels, SetInventoryStockLevels, CheckExistingOrder, AutoReceiveVendorOrders,
'		GetVendorOrderDetails, GetVendorInfo, GetVendorOrderLines,
'		AutoReceiveVendorOrder, DeleteOnhandInventory, IsOnhandInventoryInitialized,
'		GetActualFoodCost, GetStoreOnhandDates, GetVendorInventory,
'		ReceiveVendorOrder, GetStoreOrders, GetVendorReceivedOrderLines,
'		AdjustReceivedInventory, GetVendorOrderAdditionalCharge,
'		SetVendorOrderAdditionalCharge, GetStoreVendorInventory,
'		GetFoodCostByOrder, GetItemBaseCost, GetSauceBaseCost, GetStyleBaseCost,
'		GetTopperBaseCost, GetSideBaseCost, GetUnitStandardCost
'
' Revision History:
' 9/6/201 - Created
' **************************************************************************

' **************************************************************************
' Function: GetUnitStandards
' Purpose: Returns items and weights for a given unit/size.
' Parameters:	pnUnitID - The UnitID to search for
'				pnSizeID - The SizeID to search for
'				pnOrderTypeID - The OrderTypeID to search for
'				panInventoryIDs - Array of InventoryIDs found
'				panInventoryQty - Array of inventory quantities
' Return: True if sucessful, False if not
' **************************************************************************
Function GetUnitStandards(ByVal pnUnitID, ByVal pnSizeID, ByVal pnOrderTypeID, ByRef panInventoryIDs, ByRef panInventoryQty)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	GetUnitStandards = lbRet
End Function

' **************************************************************************
' Function: GetUnitSizeItemWeights
' Purpose: Returns items and weights for a given unit/size.
' Parameters:	pnUnitID - The UnitID to search for
'				pnSizeID - The SizeID to search for
'				panItemIDs - Array of ItemIDs found
'				panItemCounts1 - Array of first item count
'				padItemWeights1 - Array of first item weight
'				panItemCounts2 - Array of second item count
'				padItemWeights2 - Array of second item weight
'				padItemWeights3 - Array of third item weight
' Return: True if sucessful, False if not
' **************************************************************************
Function GetUnitSizeItemWeights(ByVal pnUnitID, ByVal pnSizeID, ByRef panItemIDs, ByRef panItemCounts1, ByRef padItemWeights1, ByRef panItemCounts2, ByRef padItemWeights2, ByRef padItemWeights3)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	GetUnitSizeItemWeights = lbRet
End Function

' **************************************************************************
' Function: GetUnitSizeSauceWeights
' Purpose: Returns sauces and weights for a given unit/size.
' Parameters:	pnUnitID - The UnitID to search for
'				pnSizeID - The SizeID to search for
'				panSauceIDs - Array of SauceIDs found
'				panSauceModifierIDs - Array of SauceModifierIDs found
'				padSauceWeights - Array of sauce weights
' Return: True if sucessful, False if not
' **************************************************************************
Function GetUnitSizeSauceWeights(ByVal pnUnitID, ByVal pnSizeID, ByRef panSauceIDs, ByRef panSauceModifierIDs, ByRef padSauceWeights)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	GetUnitSizeSauceWeights = lbRet
End Function

' **************************************************************************
' Function: GetUnitSizeStyleWeights
' Purpose: Returns styles and weights for a given unit/size.
' Parameters:	pnUnitID - The UnitID to search for
'				pnSizeID - The SizeID to search for
'				panStyleIDs - Array of StyleIDs found
'				padStyleRecipeIDs - Array of RecipeIDs found
'				padStyleWeights - Array of style weights
' Return: True if sucessful, False if not
' **************************************************************************
Function GetUnitSizeStyleWeights(ByVal pnUnitID, ByVal pnSizeID, ByRef panStyleIDs, ByRef padStyleRecipeIDs, ByRef padStyleWeights)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	GetUnitSizeStyleWeights = lbRet
End Function

' **************************************************************************
' Function: GetUnitSizeTopperWeights
' Purpose: Returns toppers and weights for a given unit/size.
' Parameters:	pnUnitID - The UnitID to search for
'				pnSizeID - The SizeID to search for
'				panTopperIDs - Array of TopperIDs found
'				padTopperWeights - Array of topper weights
' Return: True if sucessful, False if not
' **************************************************************************
Function GetUnitSizeTopperWeights(ByVal pnUnitID, ByVal pnSizeID, ByRef panTopperIDs, ByRef padTopperWeights)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	GetUnitSizeTopperWeights = lbRet
End Function

' **************************************************************************
' Function: GetUnitSizeSideWeights
' Purpose: Returns sides and weights for a given unit/size.
' Parameters:	pnUnitID - The UnitID to search for
'				pnSizeID - The SizeID to search for
'				panSideIDs - Array of SideIDs found
'				padSideWeights - Array of side weights
' Return: True if sucessful, False if not
' **************************************************************************
Function GetUnitSizeSideWeights(ByVal pnUnitID, ByVal pnSizeID, ByRef panSideIDs, ByRef padSideWeights)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	GetUnitSizeSideWeights = lbRet
End Function

' **************************************************************************
' Function: GetStoreInventory
' Purpose: Retrieves the full inventory listing for a store
' Parameters:	pnStoreID - The StoreID to search for
'				panInventoryID - Array of InventoryIDs found
'				pasInventoryDescription - Array of inventory descriptions found
'				pasCaseDescription - Array of case descriptions found
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStoreInventory(ByVal pnStoreID, ByRef panInventoryID, ByRef pasInventoryDescription, ByRef pasCaseDescription)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	GetStoreInventory = lbRet
End Function

' **************************************************************************
' Function: GetStoreOnhandInventory
' Purpose: Retrieves the onhand inventory listing for a store
' Parameters:	pnStoreID - The StoreID to search for
'				pnOnhandInventoryID - The OnhandInventoryID to search for
'				panInventoryID - Array of InventoryIDs found
'				pasInventoryDescription - Array of inventory descriptions found
'				pasCaseDescription - Array of case descriptions found
'				panOnhandQuantity - Array of onhand quantities found
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStoreOnhandInventory(ByVal pnStoreID, ByVal pnOnhandInventoryID, ByRef panInventoryID, ByRef pasInventoryDescription, ByRef pasCaseDescription, ByRef panOnhandQuantity)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	GetStoreOnhandInventory = lbRet
End Function


' **************************************************************************
' Function: GetStoreOnhandInventoryVariance
' Purpose: Retrieves the onhand inventory variance listing for a store
' Parameters:	pnStoreID - The StoreID to search for
'				pnOnhandInventoryID - The OnhandInventoryID to search for
'				panInventoryID - Array of InventoryIDs found
'				pasInventoryDescription - Array of inventory descriptions found
'				pasCaseDescription - Array of case descriptions found
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStoreOnhandInventoryVariance(ByVal pnStoreID, ByVal pnOnhandInventoryID, ByRef panInventoryID, ByRef pasInventoryDescription, ByRef pasCaseDescription)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	GetStoreOnhandInventoryVariance = lbRet
End Function

' **************************************************************************
' Function: IsOnhandInventoryInitialized
' Purpose: Determines if a store is set up for taking inventory
' Parameters:	pnStoreID - The StoreID to search for
'				pnOnhandInventoryID - The OnhandInventoryID found
'				pbIsCyclic - The flag for cyclic inventory found
'				pdtInventoryDate - The inventory date found
' Return: True if it is, False if not
' **************************************************************************
Function IsOnhandInventoryInitialized(ByVal pnStoreID, ByRef pnOnhandInventoryID, ByRef pbIsCyclic, ByRef pdtInventoryDate)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	pnOnhandInventoryID = 0
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	IsOnhandInventoryInitialized = lbRet
End Function

' **************************************************************************
' Function: InitializeOnhandInventory
' Purpose: Sets a store up for taking inventory
' Parameters:	pnStoreID - The StoreID
'				pnEmpID - The EmpID
'				psTransactionDate - The transaction date
'				pbIsCyclic - The cyclic flag
'				psInventoryDate - The inventory date
' Return: The OnhandInventoryID
' **************************************************************************
Function InitializeOnhandInventory(ByVal pnStoreID, ByVal pnEmpID, ByVal psTransactionDate, ByVal pbIsCyclic, ByVal psInventoryDate)
	Dim lnRet, lsSQL, loRS
	
	lnRet = 0
	
	' REMOVED - NOT NECESSARY FOR DEV
	
	InitializeOnhandInventory = lnRet
End Function

' **************************************************************************
' Function: SetOnhandInventory
' Purpose: Sets onhand inventory quantities
' Parameters:	pnStoreID - The StoreID
'				pnOnhandInventoryID - The OnhandInventoryID to set
'				pnInventoryID - The InventoryID to set
'				pdOnhandQuantity - The onhand quantity
' Return: True if sucessful, False if not
' **************************************************************************
Function SetOnhandInventory(ByVal pnStoreID, ByVal pnOnhandInventoryID, ByVal pnInventoryID, ByVal pdOnhandQuantity)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	SetOnhandInventory = lbRet
End Function

' **************************************************************************
' Function: SetOnhandInventoryRecount
' Purpose: Sets recounted onhand inventory quantities
' Parameters:	pnStoreID - The StoreID
'				pnOnhandInventoryID - The OnhandInventoryID to set
'				pnInventoryID - The InventoryID to set
'				pdOnhandQuantity - The onhand quantity
' Return: True if sucessful, False if not
' **************************************************************************
Function SetOnhandInventoryRecount(ByVal pnStoreID, ByVal pnOnhandInventoryID, ByVal pnInventoryID, ByVal pdOnhandQuantity)
	Dim lbRet, lsSQL, loRS
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	SetOnhandInventoryRecount = lbRet
End Function

' **************************************************************************
' Function: LockOnhandInventory
' Purpose: Locks an onhand inventory
' Parameters:	pnStoreID - The StoreID
'				pnOnhandInventoryID - The OnhandInventoryID to lock
'				pnEmpID - The EmpID that is locking inventory
'				pbIsCyclic - The cyclic flag
'				pbIsRecount - Is this lock due to recount
' Return: True if sucessful, False if not
' **************************************************************************
Function LockOnhandInventory(ByVal pnStoreID, ByVal pnOnhandInventoryID, ByVal pnEmpID, ByVal pbIsCyclic, ByVal pbIsRecount)
	Dim lbRet, lsSQL, loRS, loRS2, lsCurrentInventoryDate, lsLastInventoryDate, ldOnhandQuantity, ldIdealUsedQuantity, lnCurRec, i, lnLastOnhandInventoryID
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	LockOnhandInventory = lbRet
End Function

' **************************************************************************
' Function: GetStoreVendors
' Purpose: Retrieves the list of vendors for a store
' Parameters:	pnStoreID - The StoreID to search for
'				panVendorID - Array of VendorIDs found
'				pasVendorDescription - Array of vendor descriptions found
'				pasOrderEMail - Array of ordering e-mail addresses
'				pasAccountNumber - Array of account numbers found
'				panFirstDeliveryDay - Array of first order delivery days found
'				panSecondDeliveryDay - Array of second order delivery days found
'				panThirdDeliveryDay - Array of third order delivery days found
'				panFourthDeliveryDay - Array of fourth order delivery days found
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStoreVendors(ByVal pnStoreID, ByRef panVendorID, ByRef pasVendorDescription, pasOrderEMail, ByRef pasAccountNumber, ByRef panFirstDeliveryDay, ByRef panSecondDeliveryDay, ByRef panThirdDeliveryDay, ByRef panFourthDeliveryDay)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	GetStoreVendors = lbRet
End Function

' **************************************************************************
' Function: GetVendorOrderInventory
' Purpose: Retrieves the onhand inventory listing for a store
' Parameters:	pnStoreID - The StoreID to search for
'				pnVendorID - The VendorID to search for
'				pnOnhandInventoryID - The OnhandInventoryID
'				pnWhichOrder - Which weekly order
'				panInventoryID - Array of InventoryIDs found
'				pasInventoryDescription - Array of inventory descriptions found
'				pasUOMDescription - Array of UOM descriptions found
'				panOnhandQuantity - Array of onhand quantities found
'				panStockQuantity - Array of stock quantities found
'				pasVendorItemNumber - Arrry of vendor item numbers
'				padVendorPrice - Array of vendor prices
' Return: True if sucessful, False if not
' **************************************************************************
Function GetVendorOrderInventory(ByVal pnStoreID, ByVal pnVendorID, ByVal pnOnhandInventoryID, ByVal pnWhichOrder, ByRef panInventoryID, ByRef pasInventoryDescription, ByRef pasUOMDescription, ByRef panOnhandQuantity, ByRef panStockQuantity, ByRef pasVendorItemNumber, ByRef padVendorPrice)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	GetVendorOrderInventory = lbRet
End Function

' **************************************************************************
' Function: CreateVendorOrder
' Purpose: Creates a vendor order.
' Parameters:	pnEmpID - The EmpID
'				psTransactionDate - The transaction date
'				psExpected - The expected date
'				pnStoreID - The StoreID
'				pnVendorID - The VendorID
'				psAccountNumber - The vendor account number
'				psEMail - The vendor email address
' Return: The VendorOrderID
' **************************************************************************
Function CreateVendorOrder(ByVal pnEmpID, ByVal psTransactionDate, ByVal psExpectedDate, ByVal pnStoreID, ByVal pnVendorID, ByVal psAccountNumber, ByVal psEMail)
	Dim lnRet, lsSQL, loRS
	
	lnRet = 0
	
	' REMOVED - NOT NECESSARY FOR DEV
	
	CreateVendorOrder = lnRet
End Function

' **************************************************************************
' Function: CreateVendorOrderLine
' Purpose: Creates a vendor order line.
' Parameters:	pnVendorOrderID - The VendorOrderID
'				pnInventoryID - The InventoryID
'				pdVendorPrice - The vendor price
'				psVendorItemNumber - The vendor item number
'				pdOrderQuantity - The quantity being ordered
'				psInventoryDescription - The inventory item description
'				psUOMDescription - The UOM description
' Return: The VendorOrderID
' **************************************************************************
Function CreateVendorOrderLine(ByVal pnVendorOrderID, ByVal pnInventoryID, ByVal pdVendorPrice, ByVal psVendorItemNumber, ByVal pdOrderQuantity, ByVal psInventoryDescription, ByVal psUOMDescription)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	CreateVendorOrderLine = lbRet
End Function

' **************************************************************************
' Function: CountVendorWeekOrders
' Purpose: Determines how many orders have been placed with a vendor this week
' Parameters:	pnStoreID - The StoreID
'				pnVendorID - The VendorID
' Return: The count of orders
' **************************************************************************
Function CountVendorWeekOrders(ByVal pnStoreID, ByVal pnVendorID)
	Dim lnRet, ldtSunday, lsSQL, loRS
	
	lnRet = 0
	
	' REMOVED - NOT NECESSARY FOR DEV
	
	CountVendorWeekOrders = lnRet
End Function

' **************************************************************************
' Function: GetStoreStockLevels
' Purpose: Retrieves the full inventory stock levels for a store
' Parameters:	pnStoreID - The StoreID to search for
'				panInventoryID - Array of InventoryIDs found
'				pasInventoryDescription - Array of inventory descriptions found
'				pasCaseDescription - Array of case descriptions found
'				panDisplayOrder - Array of display orders
'				panInvoiceOrder - Array of invoice orders
'				padFirstOrder - Array of first order stock levels
'				padSecondOrder - Array of second order stock levels
'				padThirdOrder - Array of third order stock levels
'				padFourthOrder - Array of fourth order stock levels
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStoreStockLevels(ByVal pnStoreID, ByRef panInventoryID, ByRef pasInventoryDescription, ByRef pasCaseDescription, ByRef panDisplayOrder, ByRef panInvoiceOrder, ByRef padFirstOrder, ByRef padSecondOrder, ByRef padThirdOrder, ByRef padFourthOrder)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	GetStoreStockLevels = lbRet
End Function

' **************************************************************************
' Function: SetInventoryStockLevels
' Purpose: Sets a store's stock levels for an inventory item
' Parameters:	pnStoreID - The StoreID to set
'				pnInventoryID - The InventoryID to set
'				pnDisplayOrder - The display order
'				pnInvoiceOrder - The invoice order
'				pdFirstOrder - The first order stock level
'				pdSecondOrder - The second order stock level
'				pdThirdOrder - The third order stock level
'				pdFourthOrder - The fourth order stock level
' Return: True if sucessful, False if not
' **************************************************************************
Function SetInventoryStockLevels(ByVal pnStoreID, ByVal pnInventoryID, ByVal pnDisplayOrder, ByVal pnInvoiceOrder, ByVal pdFirstOrder, ByVal pdSecondOrder, ByVal pdThirdOrder, ByVal pdFourthOrder)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	SetInventoryStockLevels = lbRet
End Function

' **************************************************************************
' Function: CheckExistingOrder
' Purpose: Determines if an order has already been placed
' Parameters:	pnStoreID - The StoreID
'				pnVendorID - The VendorID
'				psExpectedDate - The expected delivery day
' Return: TRUE if order has been placed otherwise FALSE
' **************************************************************************
Function CheckExistingOrder(ByVal pnStoreID, ByVal pnVendorID, ByVal psExpectedDate)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	CheckExistingOrder = lbRet
End Function

' **************************************************************************
' Function: GetVendorOrderDetails
' Purpose: Retrieves vendor order details.
' Parameters:	pnVendorOrderID - The VendorOrderID to search for
'				pnEmpID - The employee ID
'				pdtTransaction - The transaction date
'				pdtSubmitDate - The submit date
'				pdtExpectedDate - The expected date
'				pnStoreID - The store ID
'				pnVendorID - The vendor ID
'				psAccountNumber - The account number
'				psEMail - The e-mail address
' Return: True if sucessful, False if not
' **************************************************************************
Function GetVendorOrderDetails(ByVal pnVendorOrderID, ByRef pnEmpID, ByRef pdtTransaction, ByRef pdtSubmitDate, ByRef pdtExpectedDate, ByRef pnStoreID, ByRef pnVendorID, ByRef psAccountNumber, ByRef psEMail)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	GetVendorOrderDetails = lbRet
End Function

' **************************************************************************
' Function: GetVendorInfo
' Purpose: Retrieves information about a vendor
' Parameters:	pnVendorID - The VendorIDs to search for
'				psVendorDescription - The vendor description
' Return: True if sucessful, False if not
' **************************************************************************
Function GetVendorInfo(ByVal pnVendorID, ByRef psVendorDescription)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	GetVendorInfo = lbRet
End Function

' **************************************************************************
' Function: GetVendorOrderLines
' Purpose: Retrieves the list of vendor order lines
' Parameters:	pnVendorOrderID - The VendorOrderID to search for
'				panInventoryID - Array of InventoryIDs found
'				padOrderQuantity - Array of onhand quantities found
'				pasVendorItemNumber - Array of vendor item numbers found
'				padVendorPrice - Array of vendor prices
'				pasInventoryDescription - Array of inventory descriptions found
'				pasInventoryUOM - Array of UOM descriptions found
' Return: True if sucessful, False if not
' **************************************************************************
Function GetVendorOrderLines(ByVal pnVendorOrderID, ByRef panInventoryID, ByRef padOrderQuantity, ByRef pasVendorItemNumber, ByRef padVendorPrice, ByRef pasInventoryDescription, ByRef pasInventoryUOM)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	GetVendorOrderLines = lbRet
End Function

' **************************************************************************
' Function: AutoReceiveVendorOrders
' Purpose: Automatically receives vendor orders for a store
' Parameters:	pnStoreID - The StoreID to search for
'				psExpectedDate - The expected delivery date to search for
' Return: True if sucessful, False if not
' **************************************************************************
Function AutoReceiveVendorOrders(ByVal pnStoreID, ByVal psExpectedDate)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	AutoReceiveVendorOrders = lbRet
End Function

' **************************************************************************
' Function: DeleteOnhandInventory
' Purpose: Clears a store's onhand inventory
' Parameters:	pnStoreID - The StoreID to search for
'				pnOnhandInventoryID - The OnhandInventoryID to delete
' Return: True if sucessful, False if not
' **************************************************************************
Function DeleteOnhandInventory(ByVal pnStoreID, ByVal pnOnhandInventoryID)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	DeleteOnhandInventory = lbRet
End Function

' **************************************************************************
' Function: GetActualFoodCost
' Purpose: Retrieves the actual food costs
' Parameters:	pnStoreID - The StoreID to search for
'				pdDate1 - The starting date (inclusive)
'				pdDate2 - The ending date (not inclusive)
'				pasInventoryDescription - Array of inventory descriptions found
'				pasInventoryUOM - Array of inventory UOMs found
'				padBeginningOnhand - Array of beginning onhand inventory found
'				padBeginningValue - Array of beginning values
'				padEndingOnhand - Array of ending onhand inventory found
'				padEndingValue - Array of ending values
'				padPurchased - Array of purchased inventory
'				padPurchasedCost - Array of purchased inventory total costs
'				padActualUsed - Array of actual used
'				padActualCost - Array of actual cost
'				padIdealUsed - Array of ideal used
'				padIdealCost - Array of ideal cost
' Return: True if sucessful, False if not
' **************************************************************************
Function GetActualFoodCost(ByVal pnStoreID, ByVal pdDate1, ByVal pdDate2, ByRef pasInventoryDescription, ByRef pasInventoryUOM, ByRef padBeginningOnhand, ByRef padBeginningValue, ByRef padEndingOnhand, ByRef padEndingValue, ByRef padPurchased, ByRef padPurchasedCost, ByRef padActualUsed, ByRef padActualCost, ByRef padIdealUsed, ByRef padIdealCost)
	Dim lbRet, lsSQL, loRS, lnPos, loRS2
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	GetActualFoodCost = lbRet
End Function

' **************************************************************************
' Function: GetStoreOnhandDates
' Purpose: Retrieves the list of dates onhand inventory has been taken
' Parameters:	pnStoreID - The StoreID to search for
'				pdDate - The starting date
'				padInventoryDate - Array of onhand inventory dates found
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStoreOnhandDates(ByVal pnStoreID, ByVal pdDate, ByRef padInventoryDate)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	GetStoreOnhandDates = lbRet
End Function

' **************************************************************************
' Function: GetVendorInventory
' Purpose: Retrieves the inventory listing for a store's vendor
' Parameters:	pnStoreID - The StoreID to search for
'				pnVendorID - The VendorID to search for
'				panInventoryID - Array of InventoryIDs found
'				pasInventoryDescription - Array of inventory descriptions found
'				pasUOMDescription - Array of UOM descriptions found
'				pasVendorItemNumber - Arrry of vendor item numbers
'				padVendorPrice - Array of vendor prices
' Return: True if sucessful, False if not
' **************************************************************************
Function GetVendorInventory(ByVal pnStoreID, ByVal pnVendorID, ByRef panInventoryID, ByRef pasInventoryDescription, ByRef pasUOMDescription, ByRef pasVendorItemNumber, ByRef padVendorPrice)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	GetVendorInventory = lbRet
End Function

' **************************************************************************
' Function: ReceiveVendorOrder
' Purpose: Receives a vendor order
' Parameters:	pnOrderID - The OrderID to search for
'				psReceiveDate - The date the order was received
' Return: True if sucessful, False if not
' **************************************************************************
Function ReceiveVendorOrder(ByVal pnOrderID, ByVal psReceiveDate)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	ReceiveVendorOrder = lbRet
End Function

' **************************************************************************
' Function: GetStoreOrders
' Purpose: Retrieves the list of dates onhand inventory has been taken
' Parameters:	pnStoreID - The StoreID to search for
'				pdDate - The starting date
'				panVendorOrderID - Array of VendorOrderIDs found
'				panVendorID - Array of VendorIDs found
'				pasVendorDescription - Array of VendorDescriptions found
'				padDate - Array of dates found
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStoreOrders(ByVal pnStoreID, ByVal pdDate, ByRef panVendorOrderID, ByRef panVendorID, ByRef pasVendorDescription, ByRef padExpectedDate)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	GetStoreOrders = lbRet
End Function

' **************************************************************************
' Function: GetVendorReceivedOrderLines
' Purpose: Retrieves the list of vendor received order lines
' Parameters:	pnVendorOrderID - The VendorOrderID to search for
'				panInventoryID - Array of InventoryIDs found
'				padOrderQuantity - Array of onhand quantities found
'				pasVendorItemNumber - Array of vendor item numbers found
'				padVendorPrice - Array of vendor prices
'				pasInventoryDescription - Array of inventory descriptions found
'				pasInventoryUOM - Array of UOM descriptions found
'				padReceivedQuantity - Array of received quantities found
' Return: True if sucessful, False if not
' **************************************************************************
Function GetVendorReceivedOrderLines(ByVal pnVendorOrderID, ByRef panInventoryID, ByRef padOrderQuantity, ByRef pasVendorItemNumber, ByRef padVendorPrice, ByRef pasInventoryDescription, ByRef pasInventoryUOM, ByRef padReceivedQuantity)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	GetVendorReceivedOrderLines = lbRet
End Function

' **************************************************************************
' Function: AdjustReceivedInventory
' Purpose: Adjusts the quantity and price received of an item on an order
' Parameters:	pnOrderID - The OrderID to search for
'				pnInventoryID - The InventoryID to search for
'				pdReceivedQuantity - The quantity received
'				pdVendorPrice - The vendor price
'				pdReceivedQuantityOrig - The original quantity received
'				pdVendorPriceOrig - The original vendor price
' Return: True if sucessful, False if not
' **************************************************************************
Function AdjustReceivedInventory(ByVal pnOrderID, ByVal pnInventoryID, ByVal pdReceivedQuantity, ByVal pdVendorPrice, ByVal pdReceivedQuantityOrig, ByVal pdVendorPriceOrig)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	AdjustReceivedInventory = lbRet
End Function

' **************************************************************************
' Function: GetVendorOrderAdditionalCharge
' Purpose: Retrieves the additional charge on a vendor order
' Parameters:	pnVendorOrderID - The VendorOrderIDs to search for
' Return: The additional charge
' **************************************************************************
Function GetVendorOrderAdditionalCharge(ByVal pnVendorOrderID)
	Dim ldRet, lsSQL, loRS
	
	ldRet = 0.00
	
	' REMOVED - NOT NECESSARY FOR DEV
	
	GetVendorOrderAdditionalCharge = ldRet
End Function

' **************************************************************************
' Function: SetVendorOrderAdditionalCharge
' Purpose: Sets the additional charge on a vendor order
' Parameters:	pnVendorOrderID - The VendorOrderIDs to search for
'				pdAdditionalCharge - The additional charge
' Return: True if sucessful, False if not
' **************************************************************************
Function SetVendorOrderAdditionalCharge(ByVal pnVendorOrderID, pdAdditionalCharge)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	SetVendorOrderAdditionalCharge = lbRet
End Function

' **************************************************************************
' Function: GetStoreVendorInventory
' Purpose: Retrieves the vendor inventory listing for a store
' Parameters:	pnStoreID - The StoreID to search for
'				pnVendorID - The VendorID to search for
'				panInventoryID - Array of InventoryIDs found
'				pasInventoryDescription - Array of inventory descriptions found
'				pasUOMDescription - Array of UOM descriptions found
'				pasVendorItemNumber - Arrry of vendor item numbers
'				padVendorPrice - Array of vendor prices
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStoreVendorInventory(ByVal pnStoreID, ByVal pnVendorID, ByRef panInventoryID, ByRef pasInventoryDescription, ByRef pasUOMDescription, ByRef pasVendorItemNumber, ByRef padVendorPrice)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	GetStoreVendorInventory = lbRet
End Function

' **************************************************************************
' Function: GetFoodCostByOrder
' Purpose: Retrieves the list of orders received and their costs
' Parameters:	pnStoreID - The StoreID to search for
'				pdDate1 - The starting date (inclusive)
'				pdDate2 - The ending date (not inclusive)
'				panVendorOrderID - Array of vendor order IDs
'				pasVendorDescription - Array of vendor descriptions found
'				padAdditionalCharge - Array of additional charges found
'				padTotalFoodCost - Array of total food costs
'				padtDateReceived - Array of received dates
' Return: True if sucessful, False if not
' **************************************************************************
Function GetFoodCostByOrder(ByVal pnStoreID, ByVal pdDate1, ByVal pdDate2, ByRef panVendorOrderID, ByRef pasVendorDescription, ByRef padAdditionalCharge, ByRef padTotalFoodCost, ByRef padtDateReceived)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	' REMOVED - NOT NECESSARY FOR DEV
	lbRet = TRUE
	
	GetFoodCostByOrder = lbRet
End Function

' **************************************************************************
' Function: GetItemBaseCost
' Purpose: Returns the base cost for an item/unit/size.
' Parameters:	pnStoreID = The StoreID to search for
'				pnOrderLineID - The associated OrderLineID
'				pdWeight - The associated weight
'				pnUnitID - The UnitID to search for
'				pnSizeID - The SizeID to search for
'				pnItemID - The ItemID to search for
' Return: The item's base cost
' **************************************************************************
Function GetItemBaseCost(ByVal pnStoreID, ByVal pnOrderLineID, ByVal pdWeight, ByVal pnUnitID, ByVal pnSizeID, ByVal pnItemID)
	Dim ldRet, lsSQL, loRS, ldTotalRecipeQty, ldIdealWeight, ldIdealCost
	
	ldRet = 0.0000
	ldTotalRecipeQty = 0.0000
	
	' REMOVED - NOT NECESSARY FOR DEV
	
	GetItemBaseCost = ldRet
End Function

' **************************************************************************
' Function: GetSauceBaseCost
' Purpose: Returns the base cost for an sauce/unit/size.
' Parameters:	pnStoreID = The StoreID to search for
'				pnOrderLineID - The associated OrderLineID
'				pdWeight - The associated weight
'				pnUnitID - The UnitID to search for
'				pnSizeID - The SizeID to search for
'				pnSauceID - The SauceID to search for
'				pnSauceModifierID - The SauceModifierID to search for
' Return: The sauce's base cost
' **************************************************************************
Function GetSauceBaseCost(ByVal pnStoreID, ByVal pnOrderLineID, ByVal pdWeight, ByVal pnUnitID, ByVal pnSizeID, ByVal pnSauceID, ByVal pnSauceModifierID)
	Dim ldRet, lsSQL, loRS, ldTotalRecipeQty, ldIdealWeight, ldIdealCost
	
	ldRet = 0.0000
	ldTotalRecipeQty = 0.0000
	
	' REMOVED - NOT NECESSARY FOR DEV
	
	GetSauceBaseCost = ldRet
End Function

' **************************************************************************
' Function: GetStyleBaseCost
' Purpose: Returns the base cost for an style/unit/size.
' Parameters:	pnStoreID = The StoreID to search for
'				pnOrderLineID - The associated OrderLineID
'				pdWeight - The associated weight
'				pnRecipeID - The RecipeID to search for
'				pnUnitID - The UnitID to search for
'				pnSizeID - The SizeID to search for
'				pnStyleID - The StyleID to search for
' Return: The style's base cost
' **************************************************************************
Function GetStyleBaseCost(ByVal pnStoreID, ByVal pnOrderLineID, ByVal pdWeight, ByVal pnRecipeID, ByVal pnUnitID, ByVal pnSizeID, ByVal pnStyleID)
	Dim ldRet, lsSQL, loRS, ldTotalRecipeQty, ldIdealWeight, ldIdealCost, lsRecipeDescription
	
	ldRet = 0.0000
	ldTotalRecipeQty = 0.0000
	lsRecipeDescription = ""
	
	' REMOVED - NOT NECESSARY FOR DEV
	
	GetStyleBaseCost = ldRet
End Function

' **************************************************************************
' Function: GetTopperBaseCost
' Purpose: Returns the base cost for an topper/unit/size.
' Parameters:	pnStoreID = The StoreID to search for
'				pnOrderLineID - The associated OrderLineID
'				pdWeight - The associated weight
'				pnUnitID - The UnitID to search for
'				pnSizeID - The SizeID to search for
'				pnTopperID - The TopperID to search for
' Return: The style's base cost
' **************************************************************************
Function GetTopperBaseCost(ByVal pnStoreID, ByVal pnOrderLineID, ByVal pdWeight, ByVal pnUnitID, ByVal pnSizeID, ByVal pnTopperID)
	Dim ldRet, lsSQL, loRS, ldTotalRecipeQty, ldIdealWeight, ldIdealCost
	
	ldRet = 0.0000
	ldTotalRecipeQty = 0.0000
	
	' REMOVED - NOT NECESSARY FOR DEV
	
	GetTopperBaseCost = ldRet
End Function

' **************************************************************************
' Function: GetSideBaseCost
' Purpose: Returns the base cost for an side/unit/size.
' Parameters:	pnStoreID = The StoreID to search for
'				pnOrderLineID - The associated OrderLineID
'				pdWeight - The associated weight
'				pnUnitID - The UnitID to search for
'				pnSizeID - The SizeID to search for
'				pnSideID - The SideID to search for
' Return: The style's base cost
' **************************************************************************
Function GetSideBaseCost(ByVal pnStoreID, ByVal pnOrderLineID, ByVal pdWeight, ByVal pnUnitID, ByVal pnSizeID, ByVal pnSideID)
	Dim ldRet, lsSQL, loRS, ldTotalRecipeQty, ldIdealWeight, ldIdealCost
	
	ldRet = 0.0000
	ldTotalRecipeQty = 0.0000
	
	' REMOVED - NOT NECESSARY FOR DEV
	
	GetSideBaseCost = ldRet
End Function

' **************************************************************************
' Function: GetUnitStandardCost
' Purpose: Returns the cost for an order type/unit/size.
' Parameters:	pnStoreID = The StoreID to search for
'				pnOrderLineID - The associated OrderLineID
'				pnUnitID - The UnitID to search for
'				pnSizeID - The SizeID to search for
'				pnOrderTypeID - The OrderTypeID to search for
' Return: The unit standard cost
' **************************************************************************
Function GetUnitStandardCost(ByVal pnStoreID, ByVal pnOrderLineID, ByVal pnUnitID, ByVal pnSizeID, ByVal pnOrderTypeID)
	Dim ldRet, lsSQL, loRS, ldIdealWeight, ldIdealCost
	
	ldRet = 0.0000
	
	' REMOVED - NOT NECESSARY FOR DEV
	
	GetUnitStandardCost = ldRet
End Function
%>