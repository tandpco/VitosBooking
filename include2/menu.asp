<%
' **************************************************************************
' File: menu.asp
' Purpose: Functions for menu related activities.
' Created: 7/19/2011 - TAM
' Description:
'	Include this file on any page where menu data is manipulated.
'	This file includes the following functions: GetUnits, GetSizeDescription,
'		GetStyleDescription, GetSpecialtyDescription, GetUnitDescription,
'		GetSauceDescription, GetSauceModifierDescription, GetItemDescription,
'		GetTopperDescription, GetSideDescription, GetSizeShortDescription,
'		GetStyleShortDescription, GetSpecialtyShortDescription, GetUnitShortDescription,
'		GetSauceShortDescription, GetSauceModifierShortDescription, GetItemShortDescription,
'		GetTopperShortDescription, GetSideShortDescription, GetStoreUnitSizes,
'		GetStoreSizeStyles, GetStoreUnitSauces, GetSauceModifiers, GetStoreUnitItems,
'		GetStoreUnitToppers, GetStoreUnitSides, GetStoreUnitSpecialties,
'		GetUpchargeItems, GetSpecialtySizeSideGroups, GetUnitSizeSideGroups,
'		GetSideGroups, IsTopperBeforeItems
'
' Revision History:
' 7/19/201 - Created
' **************************************************************************

' **************************************************************************
' Function: GetUnits
' Purpose: Finds units for a store.
' Parameters:	pnStoreID - The StoreID to search for
'				panUnitDs - Array of UnitIDs found
'				pasDescriptions - Array of description found
'				pasShortDescriptions - Array of short description found
' Return: True if sucessful, False if not
' **************************************************************************
Function GetUnits(ByVal pnStoreID, ByRef panUnitIDs, ByRef pasDescriptions, ByRef pasShortDescriptions)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select distinct tblUnit.UnitID, UnitDescription, UnitShortDescription, UnitMenuSortOrder from tblUnit inner join trelUnitSize on tblUnit.UnitID = trelUnitSize.UnitID inner join trelStoreUnitSize on trelUnitSize.UnitID = trelStoreUnitSize.UnitID and trelUnitSize.SizeID = trelStoreUnitSize.SizeID where StoreID = " & pnStoreID & " and tblUnit.IsActive <> 0 order by UnitMenuSortOrder"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve panUnitIDs(lnPos), pasDescriptions(lnPos), pasShortDescriptions(lnPos)
				
				panUnitIDs(lnPos) = loRS("UnitID")
				If IsNull(loRS("UnitDescription")) Then
					pasDescriptions(lnPos) = ""
				Else
					pasDescriptions(lnPos) = Trim(loRS("UnitDescription"))
				End If
				If IsNull(loRS("UnitShortDescription")) Then
					pasShortDescriptions(lnPos) = ""
				Else
					pasShortDescriptions(lnPos) = Trim(loRS("UnitShortDescription"))
				End If
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panUnitIDs(0), pasDescriptions(0), pasShortDescriptions(0)
			panUnitIDs(0) = 0
			pasDescriptions(0) = ""
			pasShortDescriptions(0) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetUnits = lbRet
End Function

' **************************************************************************
' Function: GetSizeDescription
' Purpose: Returns the description from a SizeID.
' Parameters:	pnSizeID - The SizeID to search for
' Return: Description as a string
' **************************************************************************
Function GetSizeDescription(ByVal pnSizeID)
	Dim lsRet, lsSQL, loRS
	
	lsRet = ""
	
	lsSQL = "select SizeDescription from tblSizes where SizeID = " & pnSizeID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lsRet = Trim(loRS("SizeDescription"))
		End If
		
		DBCloseQuery loRS
	End If
	
	GetSizeDescription = lsRet
End Function

' **************************************************************************
' Function: GetStyleDescription
' Purpose: Returns the description from a StyleID.
' Parameters:	pnStyleID - The StyleID to search for
' Return: Description as a string
' **************************************************************************
Function GetStyleDescription(ByVal pnStyleID)
	Dim lsRet, lsSQL, loRS
	
	lsRet = ""
	
	lsSQL = "select StyleDescription from tblStyles where StyleID = " & pnStyleID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lsRet = Trim(loRS("StyleDescription"))
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStyleDescription = lsRet
End Function

' **************************************************************************
' Function: GetSpecialtyDescription
' Purpose: Returns the description from a SpecialtyID.
' Parameters:	pnSpecialtyID - The SpecialtyID to search for
' Return: Description as a string
' **************************************************************************
Function GetSpecialtyDescription(ByVal pnSpecialtyID)
	Dim lsRet, lsSQL, loRS
	
	lsRet = ""
	
	lsSQL = "select SpecialtyDescription from tblSpecialty where SpecialtyID = " & pnSpecialtyID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lsRet = Trim(loRS("SpecialtyDescription"))
		End If
		
		DBCloseQuery loRS
	End If
	
	GetSpecialtyDescription = lsRet
End Function

' **************************************************************************
' Function: GetUnitDescription
' Purpose: Returns the description from a UnitID.
' Parameters:	pnUnitID - The UnitID to search for
' Return: Description as a string
' **************************************************************************
Function GetUnitDescription(ByVal pnUnitID)
	Dim lsRet, lsSQL, loRS
	
	lsRet = ""
	
	lsSQL = "select UnitDescription from tblUnit where UnitID = " & pnUnitID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lsRet = Trim(loRS("UnitDescription"))
		End If
		
		DBCloseQuery loRS
	End If
	
	GetUnitDescription = lsRet
End Function

' **************************************************************************
' Function: GetSauceDescription
' Purpose: Returns the description from a SauceID.
' Parameters:	pnSauceID - The SauceID to search for
' Return: Description as a string
' **************************************************************************
Function GetSauceDescription(ByVal pnSauceID)
	Dim lsRet, lsSQL, loRS
	
	lsRet = ""
	
	lsSQL = "select SauceDescription from tblSauce where SauceID = " & pnSauceID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lsRet = Trim(loRS("SauceDescription"))
		End If
		
		DBCloseQuery loRS
	End If
	
	GetSauceDescription = lsRet
End Function

' **************************************************************************
' Function: GetSauceModifierDescription
' Purpose: Returns the description from a SauceModifierID.
' Parameters:	pnSauceID - The SauceID to search for
' Return: Description as a string
' **************************************************************************
Function GetSauceModifierDescription(ByVal pnSauceModifierID)
	Dim lsRet, lsSQL, loRS
	
	lsRet = ""
	
	lsSQL = "select SauceModifierDescription from tblSauceModifier where SauceModifierID = " & pnSauceModifierID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lsRet = Trim(loRS("SauceModifierDescription"))
		End If
		
		DBCloseQuery loRS
	End If
	
	GetSauceModifierDescription = lsRet
End Function

' **************************************************************************
' Function: GetItemDescription
' Purpose: Returns the description from a ItemID.
' Parameters:	pnSauceID - The SauceID to search for
' Return: Description as a string
' **************************************************************************
Function GetItemDescription(ByVal pnItemID)
	Dim lsRet, lsSQL, loRS
	
	lsRet = ""
	
	lsSQL = "select ItemDescription from tblItems where ItemID = " & pnItemID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lsRet = Trim(loRS("ItemDescription"))
		End If
		
		DBCloseQuery loRS
	End If
	
	GetItemDescription = lsRet
End Function

' **************************************************************************
' Function: GetTopperDescription
' Purpose: Returns the description from a TopperID.
' Parameters:	pnTopperID - The TopperID to search for
' Return: Description as a string
' **************************************************************************
Function GetTopperDescription(ByVal pnTopperID)
	Dim lsRet, lsSQL, loRS
	
	lsRet = ""
	
	lsSQL = "select TopperDescription from tblTopper where TopperID = " & pnTopperID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lsRet = Trim(loRS("TopperDescription"))
		End If
		
		DBCloseQuery loRS
	End If
	
	GetTopperDescription = lsRet
End Function

' **************************************************************************
' Function: GetSideDescription
' Purpose: Returns the description from a SideID.
' Parameters:	pnSideID - The TopperID to search for
' Return: Description as a string
' **************************************************************************
Function GetSideDescription(ByVal pnSideID)
	Dim lsRet, lsSQL, loRS
	
	lsRet = ""
	
	lsSQL = "select SideDescription from tblSides where SideID = " & pnSideID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lsRet = Trim(loRS("SideDescription"))
		End If
		
		DBCloseQuery loRS
	End If
	
	GetSideDescription = lsRet
End Function

' **************************************************************************
' Function: GetSizeShortDescription
' Purpose: Returns the Short description from a SizeID.
' Parameters:	pnSizeID - The SizeID to search for
' Return: Description as a string
' **************************************************************************
Function GetSizeShortDescription(ByVal pnSizeID)
	Dim lsRet, lsSQL, loRS
	
	lsRet = ""
	
	lsSQL = "select SizeShortDescription from tblSizes where SizeID = " & pnSizeID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lsRet = Trim(loRS("SizeShortDescription"))
		End If
		
		DBCloseQuery loRS
	End If
	
	GetSizeShortDescription = lsRet
End Function

' **************************************************************************
' Function: GetStyleShortDescription
' Purpose: Returns the Short description from a StyleID.
' Parameters:	pnStyleID - The StyleID to search for
' Return: Description as a string
' **************************************************************************
Function GetStyleShortDescription(ByVal pnStyleID)
	Dim lsRet, lsSQL, loRS
	
	lsRet = ""
	
	lsSQL = "select StyleShortDescription from tblStyles where StyleID = " & pnStyleID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lsRet = Trim(loRS("StyleShortDescription"))
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStyleShortDescription = lsRet
End Function

' **************************************************************************
' Function: GetSpecialtyShortDescription
' Purpose: Returns the Short description from a SpecialtyID.
' Parameters:	pnSpecialtyID - The SpecialtyID to search for
' Return: Description as a string
' **************************************************************************
Function GetSpecialtyShortDescription(ByVal pnSpecialtyID)
	Dim lsRet, lsSQL, loRS
	
	lsRet = ""
	
	lsSQL = "select SpecialtyShortDescription from tblSpecialty where SpecialtyID = " & pnSpecialtyID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lsRet = Trim(loRS("SpecialtyShortDescription"))
		End If
		
		DBCloseQuery loRS
	End If
	
	GetSpecialtyShortDescription = lsRet
End Function

' **************************************************************************
' Function: GetUnitShortDescription
' Purpose: Returns the Short description from a UnitID.
' Parameters:	pnUnitID - The UnitID to search for
' Return: Description as a string
' **************************************************************************
Function GetUnitShortDescription(ByVal pnUnitID)
	Dim lsRet, lsSQL, loRS
	
	lsRet = ""
	
	lsSQL = "select UnitShortDescription from tblUnit where UnitID = " & pnUnitID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lsRet = Trim(loRS("UnitShortDescription"))
		End If
		
		DBCloseQuery loRS
	End If
	
	GetUnitShortDescription = lsRet
End Function

' **************************************************************************
' Function: GetSauceShortDescription
' Purpose: Returns the Short description from a SauceID.
' Parameters:	pnSauceID - The SauceID to search for
' Return: Description as a string
' **************************************************************************
Function GetSauceShortDescription(ByVal pnSauceID)
	Dim lsRet, lsSQL, loRS
	
	lsRet = ""
	
	lsSQL = "select SauceShortDescription from tblSauce where SauceID = " & pnSauceID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lsRet = Trim(loRS("SauceShortDescription"))
		End If
		
		DBCloseQuery loRS
	End If
	
	GetSauceShortDescription = lsRet
End Function

' **************************************************************************
' Function: GetSauceModifierShortDescription
' Purpose: Returns the Short description from a SauceModifierID.
' Parameters:	pnSauceID - The SauceID to search for
' Return: Description as a string
' **************************************************************************
Function GetSauceModifierShortDescription(ByVal pnSauceModifierID)
	Dim lsRet, lsSQL, loRS
	
	lsRet = ""
	
	lsSQL = "select SauceModifierShortDescription from tblSauceModifier where SauceModifierID = " & pnSauceModifierID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lsRet = Trim(loRS("SauceModifierShortDescription"))
		End If
		
		DBCloseQuery loRS
	End If
	
	GetSauceModifierShortDescription = lsRet
End Function

' **************************************************************************
' Function: GetItemShortDescription
' Purpose: Returns the Short description from a ItemID.
' Parameters:	pnSauceID - The SauceID to search for
' Return: Description as a string
' **************************************************************************
Function GetItemShortDescription(ByVal pnItemID)
	Dim lsRet, lsSQL, loRS
	
	lsRet = ""
	
	lsSQL = "select ItemShortDescription from tblItems where ItemID = " & pnItemID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lsRet = Trim(loRS("ItemShortDescription"))
		End If
		
		DBCloseQuery loRS
	End If
	
	GetItemShortDescription = lsRet
End Function

' **************************************************************************
' Function: GetTopperShortDescription
' Purpose: Returns the Short description from a TopperID.
' Parameters:	pnTopperID - The TopperID to search for
' Return: Description as a string
' **************************************************************************
Function GetTopperShortDescription(ByVal pnTopperID)
	Dim lsRet, lsSQL, loRS
	
	lsRet = ""
	
	lsSQL = "select TopperShortDescription from tblTopper where TopperID = " & pnTopperID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lsRet = Trim(loRS("TopperShortDescription"))
		End If
		
		DBCloseQuery loRS
	End If
	
	GetTopperShortDescription = lsRet
End Function

' **************************************************************************
' Function: GetSideShortDescription
' Purpose: Returns the Short description from a SideID.
' Parameters:	pnSideID - The TopperID to search for
' Return: Description as a string
' **************************************************************************
Function GetSideShortDescription(ByVal pnSideID)
	Dim lsRet, lsSQL, loRS
	
	lsRet = ""
	
	lsSQL = "select SideShortDescription from tblSides where SideID = " & pnSideID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lsRet = Trim(loRS("SideShortDescription"))
		End If
		
		DBCloseQuery loRS
	End If
	
	GetSideShortDescription = lsRet
End Function

' **************************************************************************
' Function: GetStoreUnitSizes
' Purpose: Finds sizes for a store/unit.
' Parameters:	pnStoreID - The StoreID to search for
'				pnUnitID - The UnitID to search for
'				panUnitSizeIDs - Array of SizeIDs found
'				pasUnitSizeDescriptions - Array of description found
'				pasUnitSizeShortDescriptions - Array of short description found
'				padUnitSizeStandardBasePrice - Array of standard base prices
'				panUnitSizeStandardNumberIncludedItems - Array of standard included items
'				padUnitSizeSpecialtyBasePrice - Array of specialty base prices
'				panUnitSizeSpecialtyNumberIncludedItems - Array of specialty included items
'				panUnitSizePercentSpecialtyItemVariances - Array of specialty item variance allowed
'				padPerAdditionalItemPrices - Array of per additional item prices
'				pabIsTaxable - Array of taxable flags
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStoreUnitSizes(ByVal pnStoreID, ByVal pnUnitID, ByRef panUnitSizeIDs, ByRef pasUnitSizeDescriptions, ByRef pasUnitSizeShortDescriptions, ByRef padUnitSizeStandardBasePrice, ByRef panUnitSizeStandardNumberIncludedItems, ByRef padUnitSizeSpecialtyBasePrice, ByRef panUnitSizeSpecialtyNumberIncludedItems, ByRef panUnitSizePercentSpecialtyItemVariances, ByRef padPerAdditionalItemPrices, ByRef pabIsTaxable)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select tblSizes.SizeID, SizeDescription, SizeShortDescription, StandardBasePrice, StandardNumberIncludedItems, SpecialtyBasePrice, SpecialtyNumberIncludedItems, PercentSpecialtyItemVariance, PerAdditionalItemPrice, IsTaxable from trelStoreUnitSize inner join tblSizes on trelStoreUnitSize.SizeID = tblSizes.SizeID where StoreID = " & pnStoreID & " and UnitID = " & pnUnitID & " and IsActive <> 0 order by tblSizes.SizeID"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve panUnitSizeIDs(lnPos), pasUnitSizeDescriptions(lnPos), pasUnitSizeShortDescriptions(lnPos), padUnitSizeStandardBasePrice(lnPos), panUnitSizeStandardNumberIncludedItems(lnPos), padUnitSizeSpecialtyBasePrice(lnPos), panUnitSizeSpecialtyNumberIncludedItems(lnPos), panUnitSizePercentSpecialtyItemVariances(lnPos), padPerAdditionalItemPrices(lnPos), pabIsTaxable(lnPos)
				
				panUnitSizeIDs(lnPos) = loRS("SizeID")
				If IsNull(loRS("SizeDescription")) Then
					pasUnitSizeDescriptions(lnPos) = ""
				Else
					pasUnitSizeDescriptions(lnPos) = Trim(loRS("SizeDescription"))
				End If
				If IsNull(loRS("SizeShortDescription")) Then
					pasUnitSizeShortDescriptions(lnPos) = ""
				Else
					pasUnitSizeShortDescriptions(lnPos) = Trim(loRS("SizeShortDescription"))
				End If
				If IsNull(loRS("StandardBasePrice")) Then
					padUnitSizeStandardBasePrice(lnPos) = 0.00
				Else
					padUnitSizeStandardBasePrice(lnPos) = loRS("StandardBasePrice")
				End If
				If IsNull(loRS("StandardNumberIncludedItems")) Then
					panUnitSizeStandardNumberIncludedItems(lnPos) = 0
				Else
					panUnitSizeStandardNumberIncludedItems(lnPos) = loRS("StandardNumberIncludedItems")
				End If
				If IsNull(loRS("SpecialtyBasePrice")) Then
					padUnitSizeSpecialtyBasePrice(lnPos) = 0.00
				Else
					padUnitSizeSpecialtyBasePrice(lnPos) = loRS("SpecialtyBasePrice")
				End If
				If IsNull(loRS("SpecialtyNumberIncludedItems")) Then
					panUnitSizeSpecialtyNumberIncludedItems(lnPos) = 0
				Else
					panUnitSizeSpecialtyNumberIncludedItems(lnPos) = loRS("SpecialtyNumberIncludedItems")
				End If
				If IsNull(loRS("PercentSpecialtyItemVariance")) Then
					panUnitSizePercentSpecialtyItemVariances(lnPos) = 0
				Else
					panUnitSizePercentSpecialtyItemVariances(lnPos) = loRS("PercentSpecialtyItemVariance")
				End If
				If IsNull(loRS("PerAdditionalItemPrice")) Then
					padPerAdditionalItemPrices(lnPos) = 0.00
				Else
					padPerAdditionalItemPrices(lnPos) = loRS("PerAdditionalItemPrice")
				End If
				If loRS("IsTaxable") = 0 Then
					pabIsTaxable(lnPos) = FALSE
				Else
					pabIsTaxable(lnPos) = TRUE
				End If
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panUnitSizeIDs(0), pasUnitSizeDescriptions(0), pasUnitSizeShortDescriptions(0), padUnitSizeStandardBasePrice(0), panUnitSizeStandardNumberIncludedItems(0), padUnitSizeSpecialtyBasePrice(0), panUnitSizeSpecialtyNumberIncludedItems(0), panUnitSizePercentSpecialtyItemVariances(0), padPerAdditionalItemPrices(0), pabIsTaxable(0)
			panUnitSizeIDs(0) = 0
			pasUnitSizeDescriptions(0) = ""
			pasUnitSizeShortDescriptions(0) = ""
			padUnitSizeStandardBasePrice(0) = 0.00
			panUnitSizeStandardNumberIncludedItems(0) = 0
			padUnitSizeSpecialtyBasePrice(0) = 0.00
			panUnitSizeSpecialtyNumberIncludedItems(0) = 0
			panUnitSizePercentSpecialtyItemVariances(0) = 0
			padPerAdditionalItemPrices(0) = 0.00
			pabIsTaxable(0) = FALSE
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStoreUnitSizes = lbRet
End Function

' **************************************************************************
' Function: GetStoreSizeStyles
' Purpose: Finds sizes and styles for a store/unit.
' Parameters:	pnStoreID - The StoreID to search for
'				pnUnitID - The UnitID to search for
'				panSizeStyleIDs - Array of StyleIDs found
'				pasSizeStyleDescriptions - Array of description found
'				pasSizeStyleShortDescriptions - Array of short description found
'				pasSizeStyleSpecialMessage - Array of special message found
'				panSizeStyleSizeIDs - 2 dimensional array of SizeIDs
'				padSizeStyleSurcharges - 2 dimensional array of style surcharges
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStoreSizeStyles(ByVal pnStoreID, ByVal pnUnitID, ByRef panUnitSizeIDs, ByRef panSizeStyleIDs, ByRef pasSizeStyleDescriptions, ByRef pasSizeStyleShortDescriptions, ByRef pasSizeStyleSpecialMessage, ByRef panSizeStyleSizeIDs, ByRef padSizeStyleSurcharges)
	Dim lbRet, lsSQL, loRS, lnPos, lnStyleID, lnPos2, lnPos2Max, i
	
	lbRet = FALSE
	
	lsSQL = "select tblStyles.StyleID, StyleDescription, StyleShortDescription, StyleSpecialMessage, trelSizeStyle.SizeID, StyleSurcharge from trelStoreSizeStyle inner join trelSizeStyle on trelStoreSizeStyle.StyleID = trelSizeStyle.StyleID and trelStoreSizeStyle.SizeID = trelSizeStyle.SizeID inner join tblStyles on trelSizeStyle.StyleID = tblStyles.StyleID inner join trelUnitStyles on tblStyles.StyleID = trelUnitStyles.StyleID  where StoreID = " & pnStoreID & " and UnitID = " & pnUnitID & " and IsActive <> 0 order by tblStyles.StyleID"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = -1
			lnPos2 = 0
			lnPos2Max = -1
			lnStyleID = 0
			
			Do While Not loRS.eof
				If loRS("StyleID") <> lnStyleID Then
					lnPos = lnPos + 1
					lnStyleID = loRS("StyleID")
					lnPos2 = 0
				End If
				
				If lnPos2 > lnPos2Max Then
					lnPos2Max = lnPos2
				End If
				
				lnPos2 = lnPos2 + 1
				loRS.MoveNext
			Loop
			
			ReDim Preserve panSizeStyleIDs(lnPos), pasSizeStyleDescriptions(lnPos), pasSizeStyleShortDescriptions(lnPos), pasSizeStyleSpecialMessage(lnPos), panSizeStyleSizeIDs(lnPos, lnPos2Max), padSizeStyleSurcharges(lnPos, lnPos2Max)
			
			lnPos = -1
			lnPos2 = 0
			lnStyleID = 0
			
			loRS.Requery
			Do While Not loRS.eof
				If loRS("StyleID") <> lnStyleID Then
					If lnPos <> -1 And lnPos2 <= lnPos2Max Then
						For i = lnPos2 To lnPos2Max
							panSizeStyleSizeIDs(lnPos, i) = 0
							padSizeStyleSurcharges(lnPos, i) = 0.00
						Next
					End If
					
					lnPos = lnPos + 1
					lnStyleID = loRS("StyleID")
					lnPos2 = 0
					
					panSizeStyleIDs(lnPos) = loRS("StyleID")
					If IsNull(loRS("StyleDescription")) Then
						pasSizeStyleDescriptions(lnPos) = ""
					Else
						pasSizeStyleDescriptions(lnPos) = Trim(loRS("StyleDescription"))
					End If
					If IsNull(loRS("StyleShortDescription")) Then
						pasSizeStyleShortDescriptions(lnPos) = ""
					Else
						pasSizeStyleShortDescriptions(lnPos) = Trim(loRS("StyleShortDescription"))
					End If
					If IsNull(loRS("StyleSpecialMessage")) Then
						pasSizeStyleSpecialMessage(lnPos) = ""
					Else
						pasSizeStyleSpecialMessage(lnPos) = Trim(loRS("StyleSpecialMessage"))
					End If
				End If
				
				panSizeStyleSizeIDs(lnPos, lnPos2) = loRS("SizeID")
				If IsNull(loRS("StyleSurcharge")) Then
					padSizeStyleSurcharges(lnPos, lnPos2) = 0.00
				Else
					padSizeStyleSurcharges(lnPos, lnPos2) = loRS("StyleSurcharge")
				End If
				
				lnPos2 = lnPos2 + 1
				loRS.MoveNext
			Loop
			
			For i = lnPos2 To lnPos2Max
				panSizeStyleSizeIDs(lnPos, i) = 0
				padSizeStyleSurcharges(lnPos, i) = 0.00
			Next
		Else
			ReDim panSizeStyleIDs(0), pasSizeStyleDescriptions(0), pasSizeStyleShortDescriptions(0), pasSizeStyleSpecialMessage(0), panSizeStyleSizeIDs(0, 0), padSizeStyleSurcharges(0, 0)
			panSizeStyleIDs(0) = 0
			pasSizeStyleDescriptions(0) = ""
			pasSizeStyleShortDescriptions(0) = ""
			pasSizeStyleSpecialMessage(0) = ""
			panSizeStyleSizeIDs(0, 0) = 0
			padSizeStyleSurcharges(0, 0) = 0.00
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStoreSizeStyles = lbRet
End Function

' **************************************************************************
' Function: GetStoreUnitSauces
' Purpose: Finds sauces for a unit.
' Parameters:	pnStoreID - The StoreID to search for
'				pnUnitID - The UnitID to search for
'				panSauceIDs - Array of SauceIDs found
'				pasSauceDescriptions - Array of description found
'				pasSauceShortDescriptions - Array of short description found
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStoreUnitSauces(ByVal pnStoreID, ByVal pnUnitID, ByRef panSauceIDs, ByRef pasSauceDescriptions, ByRef pasSauceShortDescriptions)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select distinct tblSauce.SauceID, SauceDescription, SauceShortDescription from trelStoreUnitSize inner join trelUnitSauce on trelStoreUnitSize.UnitID = trelUnitSauce.UnitID inner join tblSauce on trelUnitSauce.SauceID = tblSauce.SauceID where StoreID = " & pnStoreID & " and trelStoreUnitSize.UnitID = " & pnUnitID & " and IsActive <> 0 order by tblSauce.SauceID"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve panSauceIDs(lnPos), pasSauceDescriptions(lnPos), pasSauceShortDescriptions(lnPos)
				
				panSauceIDs(lnPos) = loRS("SauceID")
				If IsNull(loRS("SauceDescription")) Then
					pasSauceDescriptions(lnPos) = ""
				Else
					pasSauceDescriptions(lnPos) = Trim(loRS("SauceDescription"))
				End If
				If IsNull(loRS("SauceShortDescription")) Then
					pasSauceShortDescriptions(lnPos) = ""
				Else
					pasSauceShortDescriptions(lnPos) = Trim(loRS("SauceShortDescription"))
				End If
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panSauceIDs(0), pasSauceDescriptions(0), pasSauceShortDescriptions(0)
			panSauceIDs(0) = 0
			pasSauceDescriptions(0) = ""
			pasSauceShortDescriptions(0) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStoreUnitSauces = lbRet
End Function

' **************************************************************************
' Function: GetSauceModifiers
' Purpose: Finds sauce modifiers.
' Parameters:	panSauceModifierIDs - Array of SauceModifierIDs found
'				pasSauceModifierDescriptions - Array of description found
'				pasSauceModifierShortDescriptions - Array of short description found
' Return: True if sucessful, False if not
' **************************************************************************
Function GetSauceModifiers(ByRef panSauceModifierIDs, ByRef pasSauceModifierDescriptions, ByRef pasSauceModifierShortDescriptions)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select SauceModifierID, SauceModifierDescription, SauceModifierShortDescription from tblSauceModifier where IsActive <> 0 order by SauceModifierID"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve panSauceModifierIDs(lnPos), pasSauceModifierDescriptions(lnPos), pasSauceModifierShortDescriptions(lnPos)
				
				panSauceModifierIDs(lnPos) = loRS("SauceModifierID")
				If IsNull(loRS("SauceModifierDescription")) Then
					pasSauceModifierDescriptions(lnPos) = ""
				Else
					pasSauceModifierDescriptions(lnPos) = Trim(loRS("SauceModifierDescription"))
				End If
				If IsNull(loRS("SauceModifierShortDescription")) Then
					pasSauceModifierShortDescriptions(lnPos) = ""
				Else
					pasSauceModifierShortDescriptions(lnPos) = Trim(loRS("SauceModifierShortDescription"))
				End If
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panSauceModifierIDs(0), pasSauceModifierDescriptions(0), pasSauceModifierShortDescriptions(0)
			panSauceModifierIDs(0) = 0
			pasSauceModifierDescriptions(0) = ""
			pasSauceModifierShortDescriptions(0) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetSauceModifiers = lbRet
End Function

' **************************************************************************
' Function: GetStoreUnitItems
' Purpose: Finds items for a store/unit.
' Parameters:	pnStoreID - The StoreID to search for
'				pnUnitID - The UnitID to search for
'				panItemIDs - Array of ItemIDs found
'				pasItemDescriptions - Array of description found
'				pasItemShortDescriptions - Array of short description found
'				padItemOnSidePrice - Array on the side prices
'				panItemCounts - Array of item counts
'				pabFreeItemFlags - Array of free item flags
'				pabIsCheeses - Array of is cheese flags
'				pabIsBaseCheeses - Array of is base cheese flags
'				pabIsExtraCheeses - Array of is extra cheese flags
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStoreUnitItems(ByVal pnStoreID, ByVal pnUnitID, ByRef panItemIDs, ByRef pasItemDescriptions, ByRef pasItemShortDescriptions, ByRef padItemOnSidePrice, ByRef panItemCounts, ByRef pabFreeItemFlags, ByRef pabIsCheeses, ByRef pabIsBaseCheeses, ByRef pabIsExtraCheeses)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select tblItems.ItemID, ItemDescription, ItemShortDescription, OnSidePrice, ItemCount, FreeItemFlag, IsCheese, IsBaseCheese, IsExtraCheese from trelStoreItem inner join tblItems on trelStoreItem.ItemID = tblItems.ItemID inner join trelUnitItems on tblItems.ItemID = trelUnitItems.ItemID where StoreID = " & pnStoreID & " and UnitID = " & pnUnitID & " and IsActive <> 0 order by ItemSortOrder, tblItems.ItemID"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve panItemIDs(lnPos), pasItemDescriptions(lnPos), pasItemShortDescriptions(lnPos), padItemOnSidePrice(lnPos), panItemCounts(lnPos), pabFreeItemFlags(lnPos), pabIsCheeses(lnPos), pabIsBaseCheeses(lnPos), pabIsExtraCheeses(lnPos)
				
				panItemIDs(lnPos) = loRS("ItemID")
				If IsNull(loRS("ItemDescription")) Then
					pasItemDescriptions(lnPos) = ""
				Else
					pasItemDescriptions(lnPos) = Trim(loRS("ItemDescription"))
				End If
				If IsNull(loRS("ItemShortDescription")) Then
					pasItemShortDescriptions(lnPos) = ""
				Else
					pasItemShortDescriptions(lnPos) = Trim(loRS("ItemShortDescription"))
				End If
				If IsNull(loRS("OnSidePrice")) Then
					padItemOnSidePrice(lnPos) = 0.00
				Else
					padItemOnSidePrice(lnPos) = loRS("OnSidePrice")
				End If
				If IsNull(loRS("ItemCount")) Then
					panItemCounts(lnPos) = 1
				Else
					panItemCounts(lnPos) = loRS("ItemCount")
				End If
				If IsNull(loRS("FreeItemFlag")) Then
					pabFreeItemFlags(lnPos) = FALSE
				Else
					If loRS("FreeItemFlag") <> 0 Then
						pabFreeItemFlags(lnPos) = TRUE
					Else
						pabFreeItemFlags(lnPos) = FALSE
					End If
				End If
				If IsNull(loRS("IsCheese")) Then
					pabIsCheeses(lnPos) = FALSE
				Else
					If loRS("IsCheese") <> 0 Then
						pabIsCheeses(lnPos) = TRUE
					Else
						pabIsCheeses(lnPos) = FALSE
					End If
				End If
				If IsNull(loRS("IsBaseCheese")) Then
					pabIsBaseCheeses(lnPos) = FALSE
				Else
					If loRS("IsBaseCheese") <> 0 Then
						pabIsBaseCheeses(lnPos) = TRUE
					Else
						pabIsBaseCheeses(lnPos) = FALSE
					End If
				End If
				If IsNull(loRS("IsExtraCheese")) Then
					pabIsExtraCheeses(lnPos) = FALSE
				Else
					If loRS("IsExtraCheese") <> 0 Then
						pabIsExtraCheeses(lnPos) = TRUE
					Else
						pabIsExtraCheeses(lnPos) = FALSE
					End If
				End If
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panItemIDs(0), pasItemDescriptions(0), pasItemShortDescriptions(0), padItemOnSidePrice(0), panItemCounts(0), pabFreeItemFlags(0), pabIsCheeses(0), pabIsBaseCheeses(0), pabIsExtraCheeses(0)
			panItemIDs(0) = 0
			pasItemDescriptions(0) = ""
			pasItemShortDescriptions(0) = ""
			padItemOnSidePrice(0) = 0.00
			panItemCounts(0) = 1
			pabFreeItemFlags(0) = FALSE
			pabIsCheeses(0) = FALSE
			pabIsBaseCheeses(0) = FALSE
			pabIsExtraCheeses(0) = FALSE
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStoreUnitItems = lbRet
End Function

' **************************************************************************
' Function: GetStoreUnitToppers
' Purpose: Finds toppers for a unit.
' Parameters:	pnStoreID - The StoreID to search for
'				pnUnitID - The UnitID to search for
'				panTopperIDs - Array of TopperIDs found
'				pasTopperDescriptions - Array of description found
'				pasTopperShortDescriptions - Array of short description found
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStoreUnitToppers(ByVal pnStoreID, ByVal pnUnitID, ByRef panTopperIDs, ByRef pasTopperDescriptions, ByRef pasTopperShortDescriptions)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select distinct tblTopper.TopperID, TopperDescription, TopperShortDescription from trelStoreUnitSize inner join trelUnitTopper on trelStoreUnitSize.UnitID = trelUnitTopper.UnitID inner join tblTopper on trelUnitTopper.TopperID = tblTopper.TopperID where StoreID = " & pnStoreID & " and trelStoreUnitSize.UnitID = " & pnUnitID & " and IsActive <> 0 order by tblTopper.TopperID"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve panTopperIDs(lnPos), pasTopperDescriptions(lnPos), pasTopperShortDescriptions(lnPos)
				
				panTopperIDs(lnPos) = loRS("TopperID")
				If IsNull(loRS("TopperDescription")) Then
					pasTopperDescriptions(lnPos) = ""
				Else
					pasTopperDescriptions(lnPos) = Trim(loRS("TopperDescription"))
				End If
				If IsNull(loRS("TopperShortDescription")) Then
					pasTopperShortDescriptions(lnPos) = ""
				Else
					pasTopperShortDescriptions(lnPos) = Trim(loRS("TopperShortDescription"))
				End If
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panTopperIDs(0), pasTopperDescriptions(0), pasTopperShortDescriptions(0)
			panTopperIDs(0) = 0
			pasTopperDescriptions(0) = ""
			pasTopperShortDescriptions(0) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStoreUnitToppers = lbRet
End Function

' **************************************************************************
' Function: GetStoreUnitSides
' Purpose: Finds sides for a unit.
' Parameters:	pnStoreID - The StoreID to search for
'				pnUnitID - The UnitID to search for
'				panSideIDs - Array of SideIDs found
'				pasSideDescriptions - Array of description found
'				pasSideShortDescriptions - Array of short description found
'				padSidePrices - Array of prices found
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStoreUnitSides(ByVal pnStoreID, ByVal pnUnitID, ByRef panSideIDs, ByRef pasSideDescriptions, ByRef pasSideShortDescriptions, ByRef padSidePrices)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
'	lsSQL = "select distinct tblSides.SideID, SideDescription, SideShortDescription, SidePrice from trelStoreUnitSize inner join trelUnitSides on trelStoreUnitSize.UnitID = trelUnitSides.UnitID inner join tblSides on trelUnitSides.SideID = tblSides.SideID where StoreID = " & pnStoreID & " and trelStoreUnitSize.UnitID = " & pnUnitID & " and IsActive <> 0 order by tblSides.SideID"
	lsSQL = "select tblSides.SideID, SideDescription, SideShortDescription, SidePrice from trelStoreSides inner join trelUnitSides on trelStoreSides.SideID = trelUnitSides.SideID and trelUnitSides.UnitID = " & pnUnitID & " inner join tblSides on trelUnitSides.SideID = tblSides.SideID where StoreID = " & pnStoreID & " and IsActive <> 0 order by tblSides.SideID"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve panSideIDs(lnPos), pasSideDescriptions(lnPos), pasSideShortDescriptions(lnPos), padSidePrices(lnPos)
				
				panSideIDs(lnPos) = loRS("SideID")
				If IsNull(loRS("SideDescription")) Then
					pasSideDescriptions(lnPos) = ""
				Else
					pasSideDescriptions(lnPos) = Trim(loRS("SideDescription"))
				End If
				If IsNull(loRS("SideShortDescription")) Then
					pasSideShortDescriptions(lnPos) = ""
				Else
					pasSideShortDescriptions(lnPos) = Trim(loRS("SideShortDescription"))
				End If
				If IsNull(loRS("SidePrice")) Then
					padSidePrices(lnPos) = 0.00
				Else
					padSidePrices(lnPos) = loRS("SidePrice")
				End If
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panSideIDs(0), pasSideDescriptions(0), pasSideShortDescriptions(0), padSidePrices(0)
			panSideIDs(0) = 0
			pasSideDescriptions(0) = ""
			pasSideShortDescriptions(0) = ""
			padSidePrices(0) = 0.00
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStoreUnitSides = lbRet
End Function

' **************************************************************************
' Function: GetStoreUnitSpecialties
' Purpose: Finds specialties for a unit.
' Parameters:	pnStoreID - The StoreID to search for
'				pnUnitID - The UnitID to search for
'				panSpecialtyIDs - Array of SpecialtyIDs found
'				pasSpecialtyDescriptions - Array of description found
'				pasSpecialtyShortDescriptions - Array of short description found
'				panSpecialtySauceID - Array of SauceIDs found
'				panSpecialtyStyleID - Array of StyleIDs found
'				pabSpecialtyNoBaseCheese - Array of no base cheese flags found
'				panSpecialtyItemIDs - 2 dimensional array of ItemIDs found
'				panSpecialtyItemQuantity - 2 dimensional array of quantities found
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStoreUnitSpecialties(ByVal pnStoreID, ByVal pnUnitID, ByRef panSpecialtyIDs, ByRef pasSpecialtyDescriptions, ByRef pasSpecialtyShortDescriptions, ByRef panSpecialtySauceID, ByRef panSpecialtyStyleID, ByRef pabSpecialtyNoBaseCheese, ByRef panSpecialtyItemIDs, ByRef panSpecialtyItemQuantity)
	Dim lbRet, lsSQL, loRS, lnPos, lnSpecialtyID, lnPos2, lnPos2Max, i
	
	lbRet = FALSE
	
	lsSQL = "select tblSpecialty.SpecialtyID, SpecialtyDescription, SpecialtyShortDescription, SauceID, StyleID, NoBaseCheese, ItemID, SpecialtyItemQuantity from trelStoreSpecialty inner join tblSpecialty on trelStoreSpecialty.SpecialtyID = tblSpecialty.SpecialtyID inner join trelSpecialtyItem on tblSpecialty.SpecialtyID = trelSpecialtyItem.SpecialtyID where StoreID = " & pnStoreID & " and UnitID = " & pnUnitID & " and IsActive <> 0 order by tblSpecialty.SpecialtyID"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = -1
			lnPos2 = 0
			lnPos2Max = -1
			lnSpecialtyID = 0
			
			Do While Not loRS.eof
				If loRS("SpecialtyID") <> lnSpecialtyID Then
					lnPos = lnPos + 1
					lnSpecialtyID = loRS("SpecialtyID")
					lnPos2 = 0
				End If
				
				If lnPos2 > lnPos2Max Then
					lnPos2Max = lnPos2
				End If
				
				lnPos2 = lnPos2 + 1
				loRS.MoveNext
			Loop
			
			ReDim Preserve panSpecialtyIDs(lnPos), pasSpecialtyDescriptions(lnPos), pasSpecialtyShortDescriptions(lnPos), panSpecialtySauceID(lnPos), panSpecialtyStyleID(lnPos), pabSpecialtyNoBaseCheese(lnPos), panSpecialtyItemIDs(lnPos, lnPos2Max), panSpecialtyItemQuantity(lnPos, lnPos2Max)
			lnPos = -1
			lnPos2 = 0
			lnSpecialtyID = 0
			
			loRS.Requery
			Do While Not loRS.eof
				If loRS("SpecialtyID") <> lnSpecialtyID Then
					If lnPos <> -1 And lnPos2 <= lnPos2Max Then
						For i = lnPos2 To lnPos2Max
							panSpecialtyItemIDs(lnPos, i) = 0
							panSpecialtyItemQuantity(lnPos, i) = 0.00
						Next
					End If
					
					lnPos = lnPos + 1
					lnSpecialtyID = loRS("SpecialtyID")
					lnPos2 = 0
					
					panSpecialtyIDs(lnPos) = loRS("SpecialtyID")
					If IsNull(loRS("SpecialtyDescription")) Then
						pasSpecialtyDescriptions(lnPos) = ""
					Else
						pasSpecialtyDescriptions(lnPos) = Trim(loRS("SpecialtyDescription"))
					End If
					If IsNull(loRS("SpecialtyShortDescription")) Then
						pasSpecialtyShortDescriptions(lnPos) = ""
					Else
						pasSpecialtyShortDescriptions(lnPos) = Trim(loRS("SpecialtyShortDescription"))
					End If
					If IsNull(loRS("SauceID")) Then
						panSpecialtySauceID(lnPos) = 0
					Else
						panSpecialtySauceID(lnPos) = loRS("SauceID")
					End If
					If IsNull(loRS("StyleID")) Then
						panSpecialtyStyleID(lnPos) = 0
					Else
						panSpecialtyStyleID(lnPos) = loRS("StyleID")
					End If
					If loRS("NoBaseCheese") <> 0 Then
						pabSpecialtyNoBaseCheese(lnPos) = TRUE
					Else
						pabSpecialtyNoBaseCheese(lnPos) = FALSE
					End If
				End If
				
				panSpecialtyItemIDs(lnPos, lnPos2) = loRS("ItemID")
				If IsNull(loRS("SpecialtyItemQuantity")) Then
					panSpecialtyItemQuantity(lnPos, lnPos2) = 0.00
				Else
					panSpecialtyItemQuantity(lnPos, lnPos2) = loRS("SpecialtyItemQuantity")
				End If
				
				lnPos2 = lnPos2 + 1
				loRS.MoveNext
			Loop
			
			For i = lnPos2 To lnPos2Max
				panSpecialtyItemIDs(lnPos, i) = 0
				panSpecialtyItemQuantity(lnPos, i) = 0.00
			Next
		Else
			ReDim panSpecialtyIDs(0), pasSpecialtyDescriptions(0), pasSpecialtyShortDescriptions(0), panSpecialtySauceID(0), panSpecialtyStyleID(0), pabSpecialtyNoBaseCheese(0), panSpecialtyItemIDs(0, 0), panSpecialtyItemQuantity(0, 0)
			panSpecialtyIDs(0) = 0
			pasSpecialtyDescriptions(0) = ""
			pasSpecialtyShortDescriptions(0) = ""
			panSpecialtySauceID(0) = 0
			panSpecialtyStyleID(0) = 0
			pabSpecialtyNoBaseCheese(0) = FALSE
			panSpecialtyItemIDs(0, 0) = 0
			panSpecialtyItemQuantity(0, 0) = 0
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStoreUnitSpecialties = lbRet
End Function

' **************************************************************************
' Function: GetUpchargeItems
' Purpose: Finds upcharge items for a unit.
' Parameters:	pnStoreID - The StoreID to search for
'				pnUnitID - The UnitID to search for
'				panUpchargeSizeIDs - Array of SizeIDs found
'				panUpchargeItemIDs - 2 dimensional array of ItemIDs found
'				padUpchargePrice - 2 dimensional array of prices found
' Return: True if sucessful, False if not
' **************************************************************************
Function GetUpchargeItems(ByVal pnStoreID, ByVal pnUnitID, ByRef panUpchargeSizeIDs, ByRef panUpchargeItemIDs, ByRef padUpchargePrice)
	Dim lbRet, lsSQL, loRS, lnPos, lnSizeID, lnPos2, lnPos2Max, i
	
	lbRet = FALSE
	
	lsSQL = "select SizeID, ItemID, PremiumSurcharge from trelUpchargeItems where StoreID = " & pnStoreID & " and UnitID = " & pnUnitID & " order by SizeID, ItemID"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = -1
			lnPos2 = 0
			lnPos2Max = -1
			lnSizeID = 0
			
			Do While Not loRS.eof
				If loRS("SizeID") <> lnSizeID Then
					lnPos = lnPos + 1
					lnSizeID = loRS("SizeID")
					lnPos2 = 0
				End If
				
				If lnPos2 > lnPos2Max Then
					lnPos2Max = lnPos2
				End If
				
				lnPos2 = lnPos2 + 1
				loRS.MoveNext
			Loop
			
			ReDim Preserve panUpchargeSizeIDs(lnPos), panUpchargeItemIDs(lnPos, lnPos2Max), padUpchargePrice(lnPos, lnPos2Max)
			lnPos = -1
			lnPos2 = 0
			lnSizeID = 0
			
			loRS.Requery
			Do While Not loRS.eof
				If loRS("SizeID") <> lnSizeID Then
					If lnPos <> -1 And lnPos2 <= lnPos2Max Then
						For i = lnPos2 To lnPos2Max
							panUpchargeItemIDs(lnPos, i) = 0
							padUpchargePrice(lnPos, i) = 0.00
						Next
					End If
					
					lnPos = lnPos + 1
					lnSizeID = loRS("SizeID")
					lnPos2 = 0
					
					panUpchargeSizeIDs(lnPos) = loRS("SizeID")
				End If
				
				panUpchargeItemIDs(lnPos, lnPos2) = loRS("ItemID")
				If IsNull(loRS("PremiumSurcharge")) Then
					padUpchargePrice(lnPos, lnPos2) = 0.00
				Else
					padUpchargePrice(lnPos, lnPos2) = loRS("PremiumSurcharge")
				End If
				
				lnPos2 = lnPos2 + 1
				loRS.MoveNext
			Loop
			
			For i = lnPos2 To lnPos2Max
				panUpchargeItemIDs(lnPos, i) = 0
				padUpchargePrice(lnPos, i) = 0.00
			Next
		Else
			ReDim panUpchargeSizeIDs(0), panUpchargeItemIDs(0, 0), padUpchargePrice(0, 0)
			panUpchargeSizeIDs(0) = 0
			panUpchargeItemIDs(0, 0) = 0
			padUpchargePrice(0, 0) = 0
		End If
		
		DBCloseQuery loRS
	End If
	
	GetUpchargeItems = lbRet
End Function

' **************************************************************************
' Function: GetSpecialtySizeSideGroups
' Purpose: Finds sizes and side groups for a unit's specialties.
' Parameters:	pnStoreID - The StoreID to search for
'				pnUnitID - The UnitID to search for
'				panSideGroupSpecialtyIDs - Array of SpecialtyIDs found
'				panSideGroupSizeIDs - 2 dimensional Array of SizeIDs found
'				panSideGroupSideGroupIDs - 3 dimensional array of SideGroupID found
'				padSideGroupQuantity - 3 dimensional array of quantities found
' Return: True if sucessful, False if not
' **************************************************************************
Function GetSpecialtySizeSideGroups(ByVal pnStoreID, ByVal pnUnitID, ByRef panSideGroupSpecialtyIDs, ByRef panSideGroupSizeIDs, ByRef panSideGroupSideGroupIDs, ByRef padSideGroupQuantity)
	Dim lbRet, lsSQL, loRS, lnPos, lnSpecialtyID, lnSizeID, lnPos2, lnPos2Max, lnPos3, lnPos3Max, i, j
	
	lbRet = FALSE
	
	lsSQL = "select trelSpecialtySizeSideGroup.SpecialtyID, SizeID, SideGroupID, Quantity from trelStoreSpecialty inner join tblSpecialty on trelStoreSpecialty.SpecialtyID = tblSpecialty.SpecialtyID inner join trelSpecialtySizeSideGroup on tblSpecialty.SpecialtyID = trelSpecialtySizeSideGroup.SpecialtyID where StoreID = " & pnStoreID & " and UnitID = " & pnUnitID & " and IsActive <> 0 order by trelSpecialtySizeSideGroup.SpecialtyID, trelSpecialtySizeSideGroup.SizeID, trelSpecialtySizeSideGroup.SideGroupID"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = -1
			lnPos2 = 0
			lnPos2Max = -1
			lnPos3 = 0
			lnPos3Max = -1
			lnSpecialtyID = 0
			lnSizeID = 0
			
			Do While Not loRS.eof
				If loRS("SpecialtyID") <> lnSpecialtyID Then
					lnPos = lnPos + 1
					lnSpecialtyID = loRS("SpecialtyID")
					lnPos2 = 0
					lnSizeID = loRS("SizeID")
					lnPos3 = 0
				Else
					If loRS("SizeID") <> lnSizeID Then
						lnPos2 = lnPos2 + 1
						lnSizeID = loRS("SizeID")
						lnPos3 = 0
					End If
				End If
				
				If lnPos2 > lnPos2Max Then
					lnPos2Max = lnPos2
				End If
				
				If lnPos3 > lnPos3Max Then
					lnPos3Max = lnPos3
				End If
				
				lnPos3 = lnPos3 + 1
				loRS.MoveNext
			Loop
			
			ReDim Preserve panSideGroupSpecialtyIDs(lnPos), panSideGroupSizeIDs(lnPos, lnPos2Max), panSideGroupSideGroupIDs(lnPos, lnPos2Max, lnPos3Max), padSideGroupQuantity(lnPos, lnPos2Max, lnPos3Max)
			lnPos = -1
			lnPos2 = 0
			lnPos3 = 0
			lnSpecialtyID = 0
			lnSizeID = 0
			
			loRS.Requery
			Do While Not loRS.eof
				If loRS("SpecialtyID") <> lnSpecialtyID Then
					If lnPos <> -1 And lnPos2 <= lnPos2Max Then
						For j = lnPos3 To lnPos3Max
							panSideGroupSideGroupIDs(lnPos, lnPos2, j) = 0
							padSideGroupQuantity(lnPos, lnPos2, j) = 0.00
						Next
						
						lnPos2 = lnPos2 + 1
						For i = lnPos2 To lnPos2Max
							panSideGroupSizeIDs(lnPos, i) = 0
							lnPos3 = 0
							For j = lnPos3 To lnPos3Max
								panSideGroupSideGroupIDs(lnPos, i, j) = 0
								padSideGroupQuantity(lnPos, i, j) = 0.00
							Next
						Next
					End If
					
					lnPos = lnPos + 1
					lnSpecialtyID = loRS("SpecialtyID")
					lnPos2 = 0
					lnSizeID = loRS("SizeID")
					lnPos3 = 0
					
					panSideGroupSpecialtyIDs(lnPos) = loRS("SpecialtyID")
					panSideGroupSizeIDs(lnPos, lnPos2) = loRS("SizeID")
				Else
					If loRS("SizeID") <> lnSizeID Then
						For j = lnPos3 To lnPos3Max
							panSideGroupSideGroupIDs(lnPos, lnPos2, j) = 0
							padSideGroupQuantity(lnPos, lnPos2, j) = 0.00
						Next
						
						lnPos2 = lnPos2 + 1
						lnSizeID = loRS("SizeID")
						lnPos3 = 0
						
						panSideGroupSizeIDs(lnPos, lnPos2) = loRS("SizeID")
					End If
				End If
				
				panSideGroupSideGroupIDs(lnPos, lnPos2, lnPos3) = loRS("SideGroupID")
				padSideGroupQuantity(lnPos, lnPos2, lnPos3) = loRS("Quantity")
				
				lnPos3 = lnPos3 + 1
				loRS.MoveNext
			Loop
			
			For j = lnPos3 To lnPos3Max
				panSideGroupSideGroupIDs(lnPos, lnPos2, j) = 0
				padSideGroupQuantity(lnPos, lnPos2, j) = 0.00
			Next
			
			lnPos2 = lnPos2 + 1
			For i = lnPos2 To lnPos2Max
				panSideGroupSizeIDs(lnPos, i) = 0
				lnPos3 = 0
				For j = lnPos3 To lnPos3Max
					panSideGroupSideGroupIDs(lnPos, i, j) = 0
					padSideGroupQuantity(lnPos, i, j) = 0.00
				Next
			Next
		Else
			ReDim panSideGroupSpecialtyIDs(0), panSideGroupSizeIDs(0, 0), panSideGroupSideGroupIDs(0, 0, 0), padSideGroupQuantity(0, 0, 0)
			panSideGroupSpecialtyIDs(0) = 0
			panSideGroupSizeIDs(0, 0) = 0
			panSideGroupSideGroupIDs(0, 0, 0) = 0
			padSideGroupQuantity(0, 0, 0) = 0
		End If
		
		DBCloseQuery loRS
	End If
	
	GetSpecialtySizeSideGroups = lbRet
End Function

' **************************************************************************
' Function: GetUnitSizeSideGroups
' Purpose: Finds upcharge items for a unit.
' Parameters:	pnStoreID - The StoreID to search for
'				pnUnitID - The UnitID to search for
'				panUnitGroupSizeIDs - Array of SizeIDs found
'				panUnitGroupSideGroupIDs - 2 dimensional array of SideGroupID found
'				padUnitGroupQuantity - 2 dimensional array of quantities found
' Return: True if sucessful, False if not
' **************************************************************************
Function GetUnitSizeSideGroups(ByVal pnStoreID, ByVal pnUnitID, ByRef panUnitGroupSizeIDs, ByRef panUnitGroupSideGroupIDs, ByRef padUnitGroupQuantity)
	Dim lbRet, lsSQL, loRS, lnPos, lnSizeID, lnPos2, lnPos2Max, i
	
	lbRet = FALSE
	
	lsSQL = "select trelUnitSizeSideGroup.SizeID, SideGroupID, Quantity from trelStoreUnitSize inner join trelUnitSizeSideGroup on trelStoreUnitSize.SizeID = trelUnitSizeSideGroup.SizeID and trelStoreUnitSize.UnitID = trelUnitSizeSideGroup.UnitID where StoreID = " & pnStoreID & " and trelStoreUnitSize.UnitID = " & pnUnitID & " order by trelUnitSizeSideGroup.SizeID, trelUnitSizeSideGroup.SideGroupID"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = -1
			lnPos2 = 0
			lnPos2Max = -1
			lnSizeID = 0
			
			Do While Not loRS.eof
				If loRS("SizeID") <> lnSizeID Then
					lnPos = lnPos + 1
					lnSizeID = loRS("SizeID")
					lnPos2 = 0
				End If
				
				If lnPos2 > lnPos2Max Then
					lnPos2Max = lnPos2
				End If
				
				lnPos2 = lnPos2 + 1
				loRS.MoveNext
			Loop
			
			ReDim Preserve panUnitGroupSizeIDs(lnPos), panUnitGroupSideGroupIDs(lnPos, lnPos2Max), padUnitGroupQuantity(lnPos, lnPos2Max)
			lnPos = -1
			lnPos2 = 0
			lnSizeID = 0
			
			loRS.Requery
			Do While Not loRS.eof
				If loRS("SizeID") <> lnSizeID Then
					If lnPos <> -1 And lnPos2 <= lnPos2Max Then
						For i = lnPos2 To lnPos2Max
							panUnitGroupSideGroupIDs(lnPos, i) = 0
							padUnitGroupQuantity(lnPos, i) = 0.00
						Next
					End If
					
					lnPos = lnPos + 1
					lnSizeID = loRS("SizeID")
					lnPos2 = 0
					
					panUnitGroupSizeIDs(lnPos) = loRS("SizeID")
				End If
				
				panUnitGroupSideGroupIDs(lnPos, lnPos2) = loRS("SideGroupID")
				If IsNull(loRS("Quantity")) Then
					padUnitGroupQuantity(lnPos, lnPos2) = 0.00
				Else
					padUnitGroupQuantity(lnPos, lnPos2) = loRS("Quantity")
				End If
				
				lnPos2 = lnPos2 + 1
				loRS.MoveNext
			Loop
			
			For i = lnPos2 To lnPos2Max
				panUnitGroupSideGroupIDs(lnPos, i) = 0
				padUnitGroupQuantity(lnPos, i) = 0.00
			Next
		Else
			ReDim panUnitGroupSizeIDs(0), panUnitGroupSideGroupIDs(0, 0), padUnitGroupQuantity(0, 0)
			panUnitGroupSizeIDs(0) = 0
			panUnitGroupSideGroupIDs(0, 0) = 0
			padUnitGroupQuantity(0, 0) = 0
		End If
		
		DBCloseQuery loRS
	End If
	
	GetUnitSizeSideGroups = lbRet
End Function

' **************************************************************************
' Function: GetSideGroups
' Purpose: Finds side groups.
' Parameters:	panSideGroupIDs - Array of SpecialtyIDs found
'				pasSideGroupDescriptions - Array of description found
'				pasSideGroupShortDescriptions - Array of short description found
'				panSideGroupSideIDs - 2 dimensional array of SideIDs found
'				pasSideGroupSideDescriptions - 2 dimensional array of description found
'				pasSideGroupSideShortDescriptions - 2 dimensional array of short description found
'				pabSideGroupSideIsDefault - 2 dimensional array of IsDefault flags
' Return: True if sucessful, False if not
' **************************************************************************
Function GetSideGroups(ByRef panSideGroupIDs, ByRef pasSideGroupDescriptions, ByRef pasSideGroupShortDescriptions, ByRef panSideGroupSideIDs, ByRef pasSideGroupSideDescriptions, ByRef pasSideGroupSideShortDescriptions, ByRef pabSideGroupSideIsDefault)
	Dim lbRet, lsSQL, loRS, lnPos, lnSideGroupID, lnPos2, lnPos2Max, i
	
	lbRet = FALSE
	
	lsSQL = "select distinct tblSideGroup.SideGroupID, SideGroupDescription, SideGroupShortDescription, tblSides.SideID, SideDescription, SideShortDescription, IsDefault from tblSideGroup inner join trelSideGroupSides on tblSideGroup.SideGroupID = trelSideGroupSides.SideGroupID inner join tblSides on trelSideGroupSides.SidesID = tblSides.SideID and tblSideGroup.IsActive <> 0 and tblSides.IsActive <> 0 order by tblSideGroup.SideGroupID, tblSides.SideID"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = -1
			lnPos2 = 0
			lnPos2Max = -1
			lnSideGroupID = 0
			
			Do While Not loRS.eof
				If loRS("SideGroupID") <> lnSideGroupID Then
					lnPos = lnPos + 1
					lnSideGroupID = loRS("SideGroupID")
					lnPos2 = 0
				End If
				
				If lnPos2 > lnPos2Max Then
					lnPos2Max = lnPos2
				End If
				
				lnPos2 = lnPos2 + 1
				loRS.MoveNext
			Loop
			
			ReDim Preserve panSideGroupIDs(lnPos), pasSideGroupDescriptions(lnPos), pasSideGroupShortDescriptions(lnPos), panSideGroupSideIDs(lnPos, lnPos2Max), pasSideGroupSideDescriptions(lnPos, lnPos2Max), pasSideGroupSideShortDescriptions(lnPos, lnPos2Max), pabSideGroupSideIsDefault(lnPos, lnPos2Max)
			lnPos = -1
			lnPos2 = 0
			lnSideGroupID = 0
			
			loRS.Requery
			Do While Not loRS.eof
				If loRS("SideGroupID") <> lnSideGroupID Then
					If lnPos <> -1 And lnPos2 <= lnPos2Max Then
						For i = lnPos2 To lnPos2Max
							panSideGroupSideIDs(lnPos, i) = 0
							pasSideGroupSideDescriptions(lnPos, i) = ""
							pasSideGroupSideShortDescriptions(lnPos, i) = ""
							pabSideGroupSideIsDefault(lnPos, i) = FALSE
						Next
					End If
					
					lnPos = lnPos + 1
					lnSideGroupID = loRS("SideGroupID")
					lnPos2 = 0
					
					panSideGroupIDs(lnPos) = loRS("SideGroupID")
					If IsNull(loRS("SideGroupDescription")) Then
						pasSideGroupDescriptions(lnPos) = ""
					Else
						pasSideGroupDescriptions(lnPos) = Trim(loRS("SideGroupDescription"))
					End If
					If IsNull(loRS("SideGroupShortDescription")) Then
						pasSideGroupShortDescriptions(lnPos) = ""
					Else
						pasSideGroupShortDescriptions(lnPos) = Trim(loRS("SideGroupShortDescription"))
					End If
				End If
				
				panSideGroupSideIDs(lnPos, lnPos2) = loRS("SideID")
				If IsNull(loRS("SideDescription")) Then
					pasSideGroupSideDescriptions(lnPos, lnPos2) = ""
				Else
					pasSideGroupSideDescriptions(lnPos, lnPos2) = Trim(loRS("SideDescription"))
				End If
				If IsNull(loRS("SideShortDescription")) Then
					pasSideGroupSideShortDescriptions(lnPos, lnPos2) = ""
				Else
					pasSideGroupSideShortDescriptions(lnPos, lnPos2) = Trim(loRS("SideShortDescription"))
				End If
				If loRS("IsDefault") <> 0 Then
					pabSideGroupSideIsDefault(lnPos, lnPos2) = TRUE
				Else
					pabSideGroupSideIsDefault(lnPos, lnPos2) = FALSE
				End If
				
				lnPos2 = lnPos2 + 1
				loRS.MoveNext
			Loop
			
			For i = lnPos2 To lnPos2Max
				panSideGroupSideIDs(lnPos, i) = 0
				pasSideGroupSideDescriptions(lnPos, i) = ""
				pasSideGroupSideShortDescriptions(lnPos, i) = ""
				pabSideGroupSideIsDefault(lnPos, i) = FALSE
			Next
		Else
			ReDim panSideGroupIDs(0), pasSideGroupDescriptions(0), pasSideGroupShortDescriptions(0), panSideGroupSideIDs(0, 0), pasSideGroupSideDescriptions(0, 0), pasSideGroupSideShortDescriptions(0, 0), pabSideGroupSideIsDefault(0, 0)
			panSideGroupIDs(0) = 0
			pasSideGroupDescriptions(0) = ""
			pasSideGroupShortDescriptions(0) = ""
			panSideGroupSideIDs(0) = 0
			pasSideGroupSideDescriptions(0, 0) = ""
			pasSideGroupSideShortDescriptions(0, 0) = ""
			pabSideGroupSideIsDefault(0, 0) = FALSE
		End If
		
		DBCloseQuery loRS
	End If
	
	GetSideGroups = lbRet
End Function

' **************************************************************************
' Function: IsTopperBeforeItems
' Purpose: Determines if a topper goes on before items.
' Parameters:	pnUnitID - The UnitID to find
'				pnTopperID - The TopperID to find
' Return: True or false
' **************************************************************************
Function IsTopperBeforeItems(ByVal pnUnitID, ByVal pnTopperID)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	lsSQL = "select IsBeforeItems from trelUnitTopper where UnitID = " & pnUnitID & " and TopperID = " & pnTopperID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			If loRS("IsBeforeItems") <> 0 Then
				lbRet = TRUE
			End If
		End If
		
		DBCloseQuery loRS
	End If
	
	IsTopperBeforeItems = lbRet
End Function
%>
