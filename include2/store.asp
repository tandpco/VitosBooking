<%
' **************************************************************************
' File: store.asp
' Purpose: Functions for store related activities.
' Created: 7/19/2011 - TAM
' Description:
'	Include this file on any page where store data is manipulated.
'	This file includes the following functions: GetStoreDetails, GetStoreTaxRate
'		IsStoreChecksOK, GetStoreStationCashDrawer, GetStoreByNetwork,
'		GetStoreTransactionDate, GetStoreHours, GetCallerID, GetMarqueeText,
'		GetStoresByPostalCode, IsStoreEnabled, GetStorePostalCodes,
'		CreatePayIn, GetStores, GetDefaultDeliveryCharge, GetDefaultDriverMoney,
'		CheckExtraDelivery, GetStoreAreaCodes, GetStoreTaxRate2,
'		IsDeliveryTaxable, GetTotalSales, GetTotalPayIns, GetTotalPayOuts
'		GetTotalLabor, GetTotalDriveMoney, IsDriverSwipeRequired, 
'		GetStoreOpenCloseTime, GetPayInOutCategory
'
' Revision History:
' 7/19/201 - Created
' **************************************************************************

' **************************************************************************
' Function: GetStoreDetails
' Purpose: Retrieves store details.
' Parameters:	pnStoreID - The StoreID to search for
'				psStoreName - The store name
'				psAddress1 - The address first line
'				psAddress2 - The address second line
'				psCity - The city
'				psState - The state
'				psPostalCode - The postal code
'				psPhone - The phone number
'				psFAX - The FAX number
'				psHours - The store's hours
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStoreDetails(ByVal pnStoreID, ByRef psStoreName, ByRef psAddress1, ByRef psAddress2, ByRef psCity, ByRef psState, ByRef psPostalCode, ByRef psPhone, ByRef psFAX, ByRef psHours)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	lsSQL = "select StoreName, Address1, Address2, City, State, PostalCode, Phone, Fax, Hours from tblStores where StoreID = " & pnStoreID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lbRet = TRUE
			
			If IsNull(loRS("StoreName")) Then
				psStoreName = ""
			Else
				psStoreName = Trim(loRS("StoreName"))
			End If
			If IsNull(loRS("Address1")) Then
				psAddress1 = ""
			Else
				psAddress1 = Trim(loRS("Address1"))
			End If
			If IsNull(loRS("Address2")) Then
				psAddress2 = ""
			Else
				psAddress2 = Trim(loRS("Address2"))
			End If
			If IsNull(loRS("City")) Then
				psCity = ""
			Else
				psCity = Trim(loRS("City"))
			End If
			If IsNull(loRS("State")) Then
				psState = ""
			Else
				psState = Trim(loRS("State"))
			End If
			If IsNull(loRS("PostalCode")) Then
				psPostalCode = ""
			Else
				psPostalCode = Trim(loRS("PostalCode"))
			End If
			If IsNull(loRS("Phone")) Then
				psPhone = ""
			Else
				psPhone = Trim(loRS("Phone"))
			End If
			If IsNull(loRS("Fax")) Then
				psFAX = ""
			Else
				psFAX = Trim(loRS("Fax"))
			End If
			If IsNull(loRS("Hours")) Then
				psHours = ""
			Else
				psHours = Trim(loRS("Hours"))
			End If
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStoreDetails = lbRet
End Function

' **************************************************************************
' Function: GetStoreTaxRate
' Purpose: Retrieves the tax rate for a store.
' Parameters:	pnStoreID - The StoreID to search for
' Return: The tax rate
' **************************************************************************
Function GetStoreTaxRate(ByVal pnStoreID)
	Dim ldRet, lsSQL, loRS
	
	ldRet = -1
	
	lsSQL = "select TaxRate from tblStores where StoreID = " & pnStoreID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			ldRet = (loRS("TaxRate") / 100)
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStoreTaxRate = ldRet
End Function

' **************************************************************************
' Function: IsStoreChecksOK
' Purpose: Determines if a store accepts checks.
' Parameters:	pnStoreID - The StoreID to search for
' Return: True or false
' **************************************************************************
Function IsStoreChecksOK(ByVal pnStoreID)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	lsSQL = "select CheckOK from tblStores where StoreID = " & pnStoreID & " and CheckOK <> 0"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lbRet = TRUE
		End If
		
		DBCloseQuery loRS
	End If
	
	IsStoreChecksOK = lbRet
End Function

' **************************************************************************
' Function: GetStoreStationCashDrawer
' Purpose: Retrieves the IP address of the cash drawer for a store/station.
' Parameters:	pnStoreID - The StoreID to search for
'				psStationIPAddress - The station IP address to look for
'				psPrinterIPAddress - The Printer IP address
'				pbIsCashDrawer2 - Flag for second cash drawer
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStoreStationCashDrawer(ByVal pnStoreID, ByVal psStationIPAddress, ByRef psPrinterIPAddress, ByRef pbIsCashDrawer2)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	lsSQL = "select PrinterIPAddress, IsCashDrawer2 from trelStorePrinters inner join tblPrinters on trelStorePrinters.StoreID = tblPrinters.StoreID and trelStorePrinters.PrinterID = tblPrinters.PrinterID where trelStorePrinters.StoreID = " & pnStoreID & " and StationIPAddress = '" & DBCleanLiteral(psStationIPAddress) & "' and PrinterTypeID = 4"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lbRet = TRUE
			
			psPrinterIPAddress = Trim(loRS("PrinterIPAddress"))
			If loRS("IsCashDrawer2") <> 0 Then
				pbIsCashDrawer2 = TRUE
			Else
				pbIsCashDrawer2 = FALSE
			End If
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStoreStationCashDrawer = lbRet
End Function

' **************************************************************************
' Function: GetStoreByNetwork
' Purpose: Retrieves the store based on the network.
' Parameters:	psIPAddress - The network to search for
' Return: The StoreID
' **************************************************************************
Function GetStoreByNetwork(ByVal psIPAddress)
	Dim lnRet, lsSQL, loRS
	
	lnRet = -1
	
	If Left(psIPAddress, 9) = "10.0.254." Or Left(psIPAddress, 7) = "10.0.0." Or Left(psIPAddress, 10) = "192.168.1." Or Left(psIPAddress, 10) = "192.168.2." Then
		lnRet = 1
	Else
		If Left(psIPAddress, 7) = "10.0.1." Or Left(psIPAddress, 7) = "10.0.2." Then
			lnRet = 9
		Else
			lsSQL = "select StoreID from tblStores where networkip = '" & Left(psIPAddress, InStrRev(psIPAddress, ".")) & "0'"
			
			If DBOpenQuery(lsSQL, FALSE, loRS) Then
				If Not loRS.bof And Not loRS.eof Then
					lnRet = loRS("StoreID")
				End If
				
				DBCloseQuery loRS
			End If
		End If
	End If
	
	GetStoreByNetwork = lnRet
End Function

' **************************************************************************
' Function: GetStoreTransactionDate
' Purpose: Retrieves the current store transaction date.
' Parameters:	pnStoreID - The StoreID to search for
' Return: The current transaction date if the store is open, otherwise 1/1/1900.
' **************************************************************************
Function GetStoreTransactionDate(ByVal pnStoreID)
	Dim ldtRet, lsSQL, loRS
	
	ldtRet = DateValue("1/1/1900")
	
	lsSQL = "select top 1 CurrentStatus, ReportDate from tblStoreReportDate where StoreID = " & pnStoreID & " order by RADRAT desc"
	
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			If loRS("CurrentStatus") = "Open" Then
				ldtRet = loRS("ReportDate")
			End If
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStoreTransactionDate = ldtRet
End Function

' **************************************************************************
' Function: GetStoreHours
' Purpose: Retrieves the operating hours of a store.
' Parameters:	pnStoreID - The StoreID to search for
'				pnOpenMon - The Monday opening time
'				pnCloseMon - The Monday closing time
'				pnOpenTue - The Tuesday opening time
'				pnCloseTue - The Tuesday closing time
'				pnOpenWed - The Wednesday opening time
'				pnCloseWed - The Wednesday closing time
'				pnOpenThu - The Thursday opening time
'				pnCloseThu - The Thursday closing time
'				pnOpenFri - The Friday opening time
'				pnCloseFri - The Friday closing time
'				pnOpenSat - The Saturday opening time
'				pnCloseSat - The Saturday closing time
'				pnOpenSun - The Sunday opening time
'				pnCloseSun - The Sunday closing time
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStoreHours(ByVal pnStoreID, ByRef pnOpenMon, ByRef pnCloseMon, ByRef pnOpenTue, ByRef pnCloseTue, ByRef pnOpenWed, ByRef pnCloseWed, ByRef pnOpenThu, ByRef pnCloseThu, ByRef pnOpenFri, ByRef pnCloseFri, ByRef pnOpenSat, ByRef pnCloseSat, ByRef pnOpenSun, ByRef pnCloseSun)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	lsSQL = "select OpenMon, OpenTue, OpenWed, OpenThu, OpenFri, OpenSat, OpenSun, CloseMon, CloseTue, CloseWed, CloseThu, CloseFri, CloseSat, CloseSun from tblStores where StoreID = " & pnStoreID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			pnOpenMon = loRS("OpenMon")
			pnCloseMon = loRS("CloseMon")
			pnOpenTue = loRS("OpenTue")
			pnCloseTue = loRS("CloseTue")
			pnOpenWed = loRS("OpenWed")
			pnCloseWed = loRS("CloseWed")
			pnOpenThu = loRS("OpenThu")
			pnCloseThu = loRS("CloseThu")
			pnOpenFri = loRS("OpenFri")
			pnCloseFri = loRS("CloseFri")
			pnOpenSat = loRS("OpenSat")
			pnCloseSat = loRS("CloseSat")
			pnOpenSun = loRS("OpenSun")
			pnCloseSun = loRS("CloseSun")
		Else
			pnOpenMon = 0
			pnCloseMon = 0
			pnOpenTue = 0
			pnCloseTue = 0
			pnOpenWed = 0
			pnCloseWed = 0
			pnOpenThu = 0
			pnCloseThu = 0
			pnOpenFri = 0
			pnCloseFri = 0
			pnOpenSat = 0
			pnCloseSat = 0
			pnOpenSun = 0
			pnCloseSun = 0
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStoreHours = lbRet
End Function

' **************************************************************************
' Function: GetCallerID
' Purpose: Retrieves caller ID info for a store.
' Parameters:	pnStoreID - The StoreID to search for
'				panLineIDs - Array of line IDs
'				pasPhoneNumbers - Array of phone numbers
'				pasNames - Array of names
' Return: True if sucessful, False if not
' **************************************************************************
Function GetCallerID(ByVal pnStoreID, ByRef panLineIDs, ByRef pasPhoneNumbers, ByRef pasNames)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select LineID, CIDPhoneNumber, CIDName from tblCallerID where StoreID = " & pnStoreID & " order by LineID"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve panLineIDs(lnPos), pasPhoneNumbers(lnPos), pasNames(lnPos)
				
				panLineIDs(lnPos) = loRS("LineID")
				If IsNull(loRS("CIDPhoneNumber")) Then
					pasPhoneNumbers(lnPos) = ""
				Else
					pasPhoneNumbers(lnPos) = loRS("CIDPhoneNumber")
				End If
				If IsNull(loRS("CIDName")) Then
					pasNames(lnPos) = ""
				Else
					pasNames(lnPos) = loRS("CIDName")
				End If
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panLineIDs(0), pasPhoneNumbers(0), pasNames(0)
			panLineIDs(0) = 0
			pasPhoneNumbers(0) = ""
			pasNames(0) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetCallerID = lbRet
End Function

' **************************************************************************
' Function: GetMarqueeText
' Purpose: Retrieves the current marquee text.
' Parameters:	pnStoreID - The StoreID to search for
'				NOTE: Currently pnStoreID is unused
' Return: The current marquee text.
' **************************************************************************
Function GetMarqueeText(ByVal pnStoreID)
	Dim lsRet, lsSQL, loRS
	
	lsRet = ""
	
	lsSQL = "select MarqueeMain, MarqueeSub from tblMarquee where getdate() between StartDate and EndDate"
	
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			If Not IsNull(loRS("MarqueeMain")) Then
				lsRet = loRS("MarqueeMain")
				
				If Not IsNull(loRS("MarqueeSub")) Then
					lsRet = lsRet & " (Promo Code: " & loRS("MarqueeSub") & ")"
				End If
			End If
		End If
		
		DBCloseQuery loRS
	End If
	
	GetMarqueeText = lsRet
End Function

' **************************************************************************
' Function: GetStoresByPostalCode
' Purpose: Gets a list of stores covering a postal code.
' Parameters:	psPostalCode - The postal code to search for
'				paStoreID - Array of StoreIDs
'				paName - Array of store names
'				paAddress1 - Array of address line 1s
'				paAddress2 - Array of address line 2s
'				paCity - Array of cities
'				paStates - Array of states
'				paPostalCode - Array of postal codes
'				paPhone - Array of phone number
'				paFax - Array of FAX numbers
'				paHours - Array of hours
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStoresByPostalCode(ByVal psPostalCode, ByRef paStoreID, ByRef paName, ByRef paAddress1, ByRef paAddress2, ByRef paCity, ByRef paState, ByRef paPostalCode, ByRef paPhone, ByRef paFax, ByRef paHours)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	' For now show all
	lsSQL = "select StoreID, StoreName, Address1, Address2, City, State, PostalCode, Phone, Fax, Hours from tblStores where tblStores.StoreID > 0"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve paStoreID(lnPos), paName(lnPos), paAddress1(lnPos), paAddress2(lnPos), paCity(lnPos), paState(lnPos), paPostalCode(lnPos), paPhone(lnPos), paFax(lnPos), paHours(lnPos)
				
				paStoreID(lnPos) = loRS("StoreID")
				If IsNull(loRS("StoreName")) Then
					paName(lnPos) = ""
				Else
					paName(lnPos) = loRS("StoreName")
				End If
				If IsNull(loRS("Address1")) Then
					paAddress1(lnPos) = ""
				Else
					paAddress1(lnPos) = loRS("Address1")
				End If
				If IsNull(loRS("Address2")) Then
					paAddress2(lnPos) = ""
				Else
					paAddress2(lnPos) = loRS("Address2")
				End If
				If IsNull(loRS("City")) Then
					paCity(lnPos) = ""
				Else
					paCity(lnPos) = loRS("City")
				End If
				If IsNull(loRS("State")) Then
					paState(lnPos) = ""
				Else
					paState(lnPos) = loRS("State")
				End If
				If IsNull(loRS("PostalCode")) Then
					paPostalCode(lnPos) = ""
				Else
					paPostalCode(lnPos) = loRS("PostalCode")
				End If
				If IsNull(loRS("Phone")) Then
					paPhone(lnPos) = ""
				Else
					paPhone(lnPos) = loRS("Phone")
				End If
				If IsNull(loRS("Fax")) Then
					paFax(lnPos) = ""
				Else
					paFax(lnPos) = loRS("Fax")
				End If
				If IsNull(loRS("Hours")) Then
					paHours(lnPos) = ""
				Else
					paHours(lnPos) = loRS("Hours")
				End If
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim paStoreID(0), paName(0), paAddress1(0), paAddress2(0), paCity(0), paState(0), paPostalCode(0), paPhone(0), paFax(0), paHours(0)
			paStoreID(lnPos) = 0
			paName(lnPos) = ""
			paAddress1(lnPos) = ""
			paAddress2(lnPos) = ""
			paCity(lnPos) = ""
			paState(lnPos) = ""
			paPostalCode(lnPos) = ""
			paPhone(lnPos) = ""
			paFax(lnPos) = ""
			paHours(lnPos) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStoresByPostalCode = lbRet
End Function

' **************************************************************************
' Function: IsStoreEnabled
' Purpose: Determines if a store is enabled.
' Parameters:	pnStoreID - The StoreID to search for
' Return: TRUE or FALSE.
' **************************************************************************
Function IsStoreEnabled(ByVal pnStoreID)
	Dim lbRet, lsSQL, loRS, ldtNow, lnTime
	
	lbRet = FALSE
	
	lsSQL = "select isactive, openmon, closemon, opentue, closetue, openwed, closewed, openthu, closethu, openfri, closefri, opensat, closesat, opensun, closesun from tblStores where StoreID = " & pnStoreID
	
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
				If Not IsNull(loRS("isactive")) Then
					If loRS("isactive") Then
						ldtNow = Now
						lnTime = Hour(ldtNow) * 100 + Minute(ldtNow)
						
						Select Case Weekday(ldtNow)
							Case 1
								If loRS("closesun") < loRS("opensun") Then
									If (lnTime >= loRS("opensun") And lnTime <= 2400) or (lnTime >= 0 And lnTime <= loRS("closesun")) Then
										lbRet = TRUE
									End If
								Else
									If lnTime >= loRS("opensun") And lnTime <= loRS("closesun") Then
										lbRet = TRUE
									End If
								End If
							Case 2
								If loRS("closemon") < loRS("openmon") Then
									If (lnTime >= loRS("openmon") And lnTime <= 2400) or (lnTime >= 0 And lnTime <= loRS("closemon")) Then
										lbRet = TRUE
									End If
								Else
									If lnTime >= loRS("openmon") And lnTime <= loRS("closemon") Then
										lbRet = TRUE
									End If
								End If
							Case 3
								If loRS("closetue") < loRS("opentue") Then
									If (lnTime >= loRS("opentue") And lnTime <= 2400) or (lnTime >= 0 And lnTime <= loRS("closetue")) Then
										lbRet = TRUE
									End If
								Else
									If lnTime >= loRS("opentue") And lnTime <= loRS("closetue") Then
										lbRet = TRUE
									End If
								End If
							Case 4
								If loRS("closewed") < loRS("openwed") Then
									If (lnTime >= loRS("openwed") And lnTime <= 2400) or (lnTime >= 0 And lnTime <= loRS("closewed")) Then
										lbRet = TRUE
									End If
								Else
									If lnTime >= loRS("openwed") And lnTime <= loRS("closewed") Then
										lbRet = TRUE
									End If
								End If
							Case 5
								If loRS("closethu") < loRS("openthu") Then
									If (lnTime >= loRS("openthu") And lnTime <= 2400) or (lnTime >= 0 And lnTime <= loRS("closethu")) Then
										lbRet = TRUE
									End If
								Else
									If lnTime >= loRS("openthu") And lnTime <= loRS("closethu") Then
										lbRet = TRUE
									End If
								End If
							Case 6
								If loRS("closefri") < loRS("openfri") Then
									If (lnTime >= loRS("openfri") And lnTime <= 2400) or (lnTime >= 0 And lnTime <= loRS("closefri")) Then
										lbRet = TRUE
									End If
								Else
									If lnTime >= loRS("openfri") And lnTime <= loRS("closefri") Then
										lbRet = TRUE
									End If
								End If
							Case 7
								If loRS("closesat") < loRS("opensat") Then
									If (lnTime >= loRS("opensat") And lnTime <= 2400) or (lnTime >= 0 And lnTime <= loRS("closesat")) Then
										lbRet = TRUE
									End If
								Else
									If lnTime >= loRS("opensat") And lnTime <= loRS("closesat") Then
										lbRet = TRUE
									End If
								End If
						End Select
					End If
				End If
		End If
		
		DBCloseQuery loRS
	End If
	
	IsStoreEnabled = lbRet
End Function

' **************************************************************************
' Function: GetStorePostalCodes
' Purpose: Retrieves the postal codes associated with a store.
' Parameters:	pnStoreID - The StoreID to search for
'				pasPostalCodes - Array of postal codes
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStorePostalCodes(ByVal pnStoreID, ByRef pasPostalCodes)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select PostalCode from tblStorePostalCodes where StoreID = " & pnStoreID & " order by PostalCode"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve pasPostalCodes(lnPos)
				
				pasPostalCodes(lnPos) = loRS("PostalCode")
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim pasPostalCodes(0)
			
			pasPostalCodes(0) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStorePostalCodes = lbRet
End Function

' **************************************************************************
' Function: CreatePayIn
' Purpose: Creates a pay iin.
' Parameters:	pnStoreID - The StoreID
'				pnEmpID - The EmpID
'				psPayInFrom - Paid in from
'				pnPayInOutMethod - The PaymentTypeID
'				pnPayAmount - The amount
'				psPayInOutCheckNumber - The check number
'				psPayInOutAccountNumber - The account number
'				psPaymentReference - The payment reference
'				pnPayInOutCategory - The pay in category
'				psPayMemo - The memo
'				psTransactionDate - The transaction date
' Return: True if sucessful, False if not
' **************************************************************************
Function CreatePayIn(ByVal pnStoreID, ByVal pnEmpID, ByVal psPayInFrom, ByVal pnPayInOutMethod, ByVal pnPayAmount, ByVal psPayInOutCheckNumber, ByVal psPayInOutAccountNumber, ByVal psPaymentReference, ByVal pnPayInOutCategory, ByVal psPayMemo, ByVal psTransactionDate)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	lsSQL = "Insert Into tblPayInOut(StoreID, EmployeeID, Who, PaymentTypeID, PayInOut, Amount, CheckNumber, AccountNumber, PaymentReference, CategoryID, Memo, TransactionDate) Values (" & pnStoreID & ", " & pnEmpID & ", '" & DBCleanLiteral(psPayInFrom) &"', " & pnPayInOutMethod & ", 'IN', " & pnPayAmount & ", '" & DBCleanLiteral(psPayInOutCheckNumber) & "', '" & DBCleanLiteral(psPayInOutAccountNumber) & "', '" & DBCleanLiteral(psPaymentReference) & "', " & pnPayInOutCategory & ", '" & DBCleanLiteral(psPayMemo) & "', '" & DBCleanLiteral(psTransactionDate) & "')"
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If
	
	CreatePayIn = lbRet
End Function

' **************************************************************************
' Function: GetStores
' Purpose: Gets a list of stores.
' Parameters:	paStoreID - Array of StoreIDs
'				paAddress1 - Array of address line 1s
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStores(ByRef paStoreID, ByRef paAddress1)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select StoreID, Address1 from tblStores where tblStores.StoreID > 0 order by StoreID"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve paStoreID(lnPos), paAddress1(lnPos)
				
				paStoreID(lnPos) = loRS("StoreID")
				If IsNull(loRS("Address1")) Then
					paAddress1(lnPos) = ""
				Else
					paAddress1(lnPos) = loRS("Address1")
				End If
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim paStoreID(0), paAddress1(0)
			
			paStoreID(lnPos) = 0
			paAddress1(lnPos) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStores = lbRet
End Function

' **************************************************************************
' Function: GetDefaultDeliveryCharge
' Purpose: Gets a store's default delivery charge.
' Parameters:	pnStoreID - The StoreID
' Return: The default delivery charge
' **************************************************************************
Function GetDefaultDeliveryCharge(ByVal pnStoreID)
	Dim ldRet, lsSQL, loRS
	
	ldRet = 0.00
	
	lsSQL = "select DefaultDeliveryCharge from tblStores where StoreID = " & pnStoreID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			ldRet = loRS("DefaultDeliveryCharge")
		End If
		
		DBCloseQuery loRS
	End If
	
	GetDefaultDeliveryCharge = ldRet
End Function

' **************************************************************************
' Function: GetDefaultDriverMoney
' Purpose: Gets a store's default driver money.
' Parameters:	pnStoreID - The StoreID
' Return: The default driver money
' **************************************************************************
Function GetDefaultDriverMoney(ByVal pnStoreID)
	Dim ldRet, lsSQL, loRS
	
	ldRet = 0.00
	
	lsSQL = "select DefaultDriverMoney from tblStores where StoreID = " & pnStoreID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			ldRet = loRS("DefaultDriverMoney")
		End If
		
		DBCloseQuery loRS
	End If
	
	GetDefaultDriverMoney = ldRet
End Function

' **************************************************************************
' Function: CheckExtraDelivery
' Purpose: Determines if a delivery charge is higher than normal and not
'			been previously disclosed.
' Parameters:	pnStoreID - The StoreID
' Return: TRUE or FALSE
' **************************************************************************
Function CheckExtraDelivery(ByVal pnStoreID, ByVal pnCustomerID, ByVal pnAddressID, ByVal pdDeliveryCharge)
	Dim lbRet, lsSQL, loRS, ldDefaultCharge
	
	lbRet = FALSE
	
	ldDefaultCharge = GetDefaultDeliveryCharge(pnStoreID)
	If pdDeliveryCharge > ldDefaultCharge Then
		lsSQL = "select WasExtraDeliveryNotified from trelCustomerAddresses where CustomerID = " & pnCustomerID & " and AddressID = " & pnAddressID
		If DBOpenQuery(lsSQL, TRUE, loRS) Then
			If Not loRS.bof And Not loRS.eof Then
				If loRS("WasExtraDeliveryNotified") = 0 Then
					lbRet = TRUE
					
					loRS("WasExtraDeliveryNotified") = 1
					
					loRS.Update
				End If
			End If
			
			DBCloseQuery loRS
		End If
	End If
	
	CheckExtraDelivery = lbRet
End Function

' **************************************************************************
' Function: GetStoreAreaCodes
' Purpose: Retrieves the area codes associated with a store.
' Parameters:	pnStoreID - The StoreID to search for
'				panAreaCodes - Array of area codes
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStoreAreaCodes(ByVal pnStoreID, ByRef panAreaCodes)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select AreaCode from tblStoreAreaCodes where StoreID = " & pnStoreID & " order by StoreAreaCodeID"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve panAreaCodes(lnPos)
				
				panAreaCodes(lnPos) = loRS("AreaCode")
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panAreaCodes(0)
			
			panAreaCodes(0) = 0
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStoreAreaCodes = lbRet
End Function

' **************************************************************************
' Function: GetStoreTaxRate2
' Purpose: Retrieves the second tax rate for a store.
' Parameters:	pnStoreID - The StoreID to search for
' Return: The second tax rate
' **************************************************************************
Function GetStoreTaxRate2(ByVal pnStoreID)
	Dim ldRet, lsSQL, loRS
	
	ldRet = -1
	
	lsSQL = "select TaxRate2 from tblStores where StoreID = " & pnStoreID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			ldRet = (loRS("TaxRate2") / 100)
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStoreTaxRate2 = ldRet
End Function

' **************************************************************************
' Function: IsDeliveryTaxable
' Purpose: Determines if delivery charge is taxable.
' Parameters:	pnStoreID - The StoreID to search for
' Return: True or false
' **************************************************************************
Function IsDeliveryTaxable(ByVal pnStoreID)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	lsSQL = "select IsDeliveryTaxable from tblStores where StoreID = " & pnStoreID & " and IsDeliveryTaxable <> 0"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lbRet = TRUE
		End If
		
		DBCloseQuery loRS
	End If
	
	IsDeliveryTaxable = lbRet
End Function

' **************************************************************************
' Function: GetTotalSales
' Purpose: Returns the total sales from one date to next.
' Parameters:	pnStoreID - The StoreID to search for
'				pdtDate1 - The starting date (inclusive)
'				pdtDate2 - The ending date (not inclusive)
'				pbIncludeDiscounts - Flag indicating if amount should include discounts
' Return: The total sales
' **************************************************************************
Function GetTotalSales(ByVal pnStoreID, ByVal pdtDate1, ByVal pdtDate2, ByVal pbIncludeDiscounts)
	Dim ldRet, lsSQL, loRS, loRS2
	
	ldRet = 0.00
	
'	lsSQL = "select sum(DeliveryCharge - DriverMoney) AS TotalDeliveryCharge from tblOrders where OrderStatusID >= 3 and OrderStatusID <= 10 and TransactionDate > '" & pdtDate1 & "' and TransactionDate <= '" & pdtDate2 & "' and StoreID = " & pnStoreID
'	lsSQL = "select sum(DeliveryCharge) AS TotalDeliveryCharge from tblOrders where OrderStatusID >= 3 and OrderStatusID <= 10 and TransactionDate > '" & pdtDate1 & "' and TransactionDate <= '" & pdtDate2 & "' and StoreID = " & pnStoreID
	lsSQL = "select sum(DeliveryCharge) AS TotalDeliveryCharge from tblOrders where OrderStatusID >= 3 and OrderStatusID <= 10 and ReleaseDate >= '" & pdtDate1 & "' and ReleaseDate < '" & pdtDate2 & "' and StoreID = " & pnStoreID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lsSQL = "select sum(dbo.tblOrderLines.Quantity * (dbo.tblOrderLines.Cost"
			If pbIncludeDiscounts Then
				lsSQL = lsSQL & " - dbo.tblOrderLines.Discount"
			End If
'			lsSQL = lsSQL & ")) as OrderTotal from tblOrderLines inner join tblOrders on tblOrderLines.OrderID = tblOrders.OrderID where OrderStatusID >= 3 and OrderStatusID <= 10 and TransactionDate > '" & pdtDate1 & "' and TransactionDate <= '" & pdtDate2 & "' and StoreID = " & pnStoreID
			lsSQL = lsSQL & ")) as OrderTotal from tblOrderLines inner join tblOrders on tblOrderLines.OrderID = tblOrders.OrderID where OrderStatusID >= 3 and OrderStatusID <= 10 and ReleaseDate >= '" & pdtDate1 & "' and ReleaseDate < '" & pdtDate2 & "' and StoreID = " & pnStoreID
			If DBOpenQuery(lsSQL, FALSE, loRS2) Then
				If Not loRS2.bof And Not loRS2.eof Then
					If Not IsNull(loRS("TotalDeliveryCharge")) Then
						ldRet = loRS("TotalDeliveryCharge")
					End If
					If Not IsNull(loRS2("OrderTotal")) Then
						ldRet = ldRet + loRS2("OrderTotal")
					End If
				End If
				
				DBCloseQuery loRS2
			End If
		End If
		
		DBCloseQuery loRS
	End If
	
	GetTotalSales = ldRet
End Function

' **************************************************************************
' Function: GetTotalPayIns
' Purpose: Returns the total pay ins from one date to next.
' Parameters:	pnStoreID - The StoreID to search for
'				pdtDate1 - The starting date (inclusive)
'				pdtDate2 - The ending date (not inclusive)
' Return: The total pay ins
' **************************************************************************
Function GetTotalPayIns(ByVal pnStoreID, ByVal pdtDate1, ByVal pdtDate2)
	Dim ldRet, lsSQL, loRS
	
	ldRet = 0.00
	
'	lsSQL = "select sum(Amount) AS TotalPayIns from tblPayInOut where TransactionDate > '" & pdtDate1 & "' and TransactionDate <= '" & pdtDate2 & "' and StoreID = " & pnStoreID & " and PayInOut = 'IN'"
	lsSQL = "select sum(Amount) AS TotalPayIns from tblPayInOut where RADRAT >= '" & pdtDate1 & "' and RADRAT < '" & pdtDate2 & "' and StoreID = " & pnStoreID & " and PayInOut = 'IN'"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			If Not IsNull(loRS("TotalPayIns")) Then
				ldRet = loRS("TotalPayIns")
			End If
		End If
		
		DBCloseQuery loRS
	End If
	
	GetTotalPayIns = ldRet
End Function

' **************************************************************************
' Function: GetTotalPayOuts
' Purpose: Returns the total pay outs from one date to next.
' Parameters:	pnStoreID - The StoreID to search for
'				pdtDate1 - The starting date (inclusive)
'				pdtDate2 - The ending date (not inclusive)
' Return: The total pay ins
' **************************************************************************
Function GetTotalPayOuts(ByVal pnStoreID, ByVal pdtDate1, ByVal pdtDate2)
	Dim ldRet, lsSQL, loRS
	
	ldRet = 0.00
	
'	lsSQL = "select sum(Amount) AS TotalPayOuts from tblPayInOut where TransactionDate > '" & pdtDate1 & "' and TransactionDate <= '" & pdtDate2 & "' and StoreID = " & pnStoreID & " and PayInOut = 'Out'"
	lsSQL = "select sum(Amount) AS TotalPayOuts from tblPayInOut where RADRAT >= '" & pdtDate1 & "' and RADRAT < '" & pdtDate2 & "' and StoreID = " & pnStoreID & " and PayInOut = 'Out'"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			If Not IsNull(loRS("TotalPayOuts")) Then
				ldRet = loRS("TotalPayOuts")
			End If
		End If
		
		DBCloseQuery loRS
	End If
	
	GetTotalPayOuts = ldRet
End Function

' **************************************************************************
' Function: GetTotalLabor
' Purpose: Returns the total labor from one date to next.
' Parameters:	pnStoreID - The StoreID to search for
'				pdtDate1 - The starting date (inclusive)
'				pdtDate2 - The ending date (not inclusive)
' Return: The total labor
' **************************************************************************
Function GetTotalLabor(ByVal pnStoreID, ByVal pdtDate1, ByVal pdtDate2)
	Dim ldRet, lsSQL, loRS, ldtStart, ldtEnd, loRS2, lbOTOK, ldTime
	
	ldRet = 0.00
	
	' Full shifts inside the range
	lsSQL = "select SUM(Rate * (datediff(mi, PunchInTime, PunchOutTime) / 60.00)) As TotalWage from tblShifts where PunchInTime >= '" & pdtDate1 & "' and PunchOutTime < '" & pdtDate2 & "' and StoreID = " & pnStoreID & " and PunchOutTime is not null"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			If Not IsNull(loRS("TotalWage")) Then
				ldRet = ldRet + loRS("TotalWage")
			End If
		End If
		
		DBCloseQuery loRS
		
		' Full shifts overlapping the start date
		lsSQL = "select SUM(Rate * (datediff(mi, '" & pdtDate1 & "', PunchOutTime) / 60.00)) As TotalWage from tblShifts where PunchInTime < '" & pdtDate1 & "' and PunchOutTime >= '" & pdtDate1 & "' and PunchOutTime < '" & pdtDate2 & "' and StoreID = " & pnStoreID & " and PunchOutTime is not null"
		If DBOpenQuery(lsSQL, FALSE, loRS) Then
			If Not loRS.bof And Not loRS.eof Then
				If Not IsNull(loRS("TotalWage")) Then
					ldRet = ldRet + loRS("TotalWage")
				End If
			End If
			
			DBCloseQuery loRS
			
			' Full shifts overlapping the end date
			lsSQL = "select SUM(Rate * (datediff(mi, PunchInTime, '" & pdtDate2 & "') / 60.00)) As TotalWage from tblShifts where PunchInTime >= '" & pdtDate1 & "' and PunchInTime < '" & pdtDate2 & "' and PunchOutTime >= '" & pdtDate2 & "' and StoreID = " & pnStoreID & " and PunchOutTime is not null"
			If DBOpenQuery(lsSQL, FALSE, loRS) Then
				If Not loRS.bof And Not loRS.eof Then
					If Not IsNull(loRS("TotalWage")) Then
						ldRet = ldRet + loRS("TotalWage")
					End If
				End If
				
				DBCloseQuery loRS
				
				' Full shifts overlapping the entire range
				lsSQL = "select SUM(Rate * (datediff(mi, '" & pdtDate1 & "', '" & pdtDate2 & "') / 60.00)) As TotalWage from tblShifts where PunchInTime < '" & pdtDate1 & "' and PunchOutTime > '" & pdtDate2 & "' and StoreID = " & pnStoreID & " and PunchOutTime is not null"
				If DBOpenQuery(lsSQL, FALSE, loRS) Then
					If Not loRS.bof And Not loRS.eof Then
						If Not IsNull(loRS("TotalWage")) Then
							ldRet = ldRet + loRS("TotalWage")
						End If
					End If
					
					DBCloseQuery loRS
					
					' Full shifts not yet punched out started before end date
					lsSQL = "select SUM(Rate * (datediff(mi, PunchInTime, '" & pdtDate2 & "') / 60.00)) As TotalWage from tblShifts where PunchInTime < '" & pdtDate2 & "' and PunchOutTime is null and StoreID = " & pnStoreID
					If DBOpenQuery(lsSQL, FALSE, loRS) Then
						If Not loRS.bof And Not loRS.eof Then
							If Not IsNull(loRS("TotalWage")) Then
								ldRet = ldRet + loRS("TotalWage")
							End If
						End If
						
						DBCloseQuery loRS
					End If
				End If
			End If
		End If
	End If
	
	GetTotalLabor = ldRet
End Function

' **************************************************************************
' Function: GetTotalDriveMoney
' Purpose: Returns the total drive money from one date to next.
' Parameters:	pnStoreID - The StoreID to search for
'				pdtDate1 - The starting date (inclusive)
'				pdtDate2 - The ending date (not inclusive)
' Return: The total drive money
' **************************************************************************
Function GetTotalDriveMoney(ByVal pnStoreID, ByVal pdtDate1, ByVal pdtDate2)
	Dim ldRet, lsSQL, loRS
	
	ldRet = 0.00
	
'	lsSQL = "select sum(DriverMoney) AS TotalDriveMoney from tblOrders where OrderStatusID >= 3 and OrderStatusID <= 10 and TransactionDate > '" & pdtDate1 & "' and TransactionDate <= '" & pdtDate2 & "' and StoreID = " & pnStoreID
	lsSQL = "select sum(DriverMoney) AS TotalDriveMoney from tblOrders where OrderStatusID >= 3 and OrderStatusID <= 10 and ReleaseDate >= '" & pdtDate1 & "' and ReleaseDate < '" & pdtDate2 & "' and StoreID = " & pnStoreID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			If Not IsNull(loRS("TotalDriveMoney")) Then
				ldRet = loRS("TotalDriveMoney")
			End If
		End If
		
		DBCloseQuery loRS
	End If
	
	GetTotalDriveMoney = ldRet
End Function

' **************************************************************************
' Function: IsDriverSwipeRequired
' Purpose: Determines if a driver is required to swipe for driver dispatch.
' Parameters:	pnStoreID - The StoreID to search for
' Return: True or false
' **************************************************************************
Function IsDriverSwipeRequired(ByVal pnStoreID)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	lsSQL = "select RequireDriverSwipe from tblStores where StoreID = " & pnStoreID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			If loRS("RequireDriverSwipe") <> 0 Then
				lbRet = TRUE
			End If
		End If
		
		DBCloseQuery loRS
	End If
	
	IsDriverSwipeRequired = lbRet
End Function

' **************************************************************************
' Function: GetStoreOpenCloseTime
' Purpose: Returns a store's opening and closing time for a given day of the week.
' Parameters:	pnStoreID - The StoreID to search for
'				pnDOW - The day of the week (1 = Sunday)
'				pnOpenTime - The opening time found
'				pnCloseTime - The closing time found
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStoreOpenCloseTime(ByVal pnStoreID, ByVal pnDOW, ByRef pnOpenTime, ByRef pnCloseTime)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	lsSQL = "select "
	Select Case pnDOW
		Case 1
			lsSQL = lsSQL & "OpenSun as OpenTime, CloseSun as CloseTime "
		Case 2
			lsSQL = lsSQL & "OpenMon as OpenTime, CloseMon as CloseTime "
		Case 3
			lsSQL = lsSQL & "OpenTue as OpenTime, CloseTue as CloseTime "
		Case 4
			lsSQL = lsSQL & "OpenWed as OpenTime, CloseWed as CloseTime "
		Case 5
			lsSQL = lsSQL & "OpenThu as OpenTime, CloseThu as CloseTime "
		Case 6
			lsSQL = lsSQL & "OpenFri as OpenTime, CloseFri as CloseTime "
		Case 7
			lsSQL = lsSQL & "OpenSat as OpenTime, CloseSat as CloseTime "
	End Select
	lsSQL = lsSQL & "from tblStores where StoreID = " & pnStoreID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lbRet = TRUE
			pnOpenTime = loRS("OpenTime")
			pnCloseTime = loRS("CloseTime")
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStoreOpenCloseTime = lbRet
End Function

' **************************************************************************
' Function: GetPayInOutCategory
' Purpose: Returns a store's opening and closing time for a given day of the week.
' Parameters:	pnPayInOutCategoryID - The PayInOutCategoryID to search for
' Return: The PayInOutCategory
' **************************************************************************
Function GetPayInOutCategory(ByVal pnPayInOutCategoryID)
	Dim lsRet, lsSQL, loRS
	
	lsRet = ""
	
	lsSQL = "select PayInOutCategory from tlkpPayInOutCategories where PayInOutCategoryID = " & pnPayInOutCategoryID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lsRet = loRS("PayInOutCategory")
		End If
		
		DBCloseQuery loRS
	End If
	
	GetPayInOutCategory = lsRet
End Function
%>