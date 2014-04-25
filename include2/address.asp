<%
' **************************************************************************
' File: address.asp
' Purpose: Functions for address related activities.
' Created: 7/19/2011 - TAM
' Description:
'	Include this file on any page where address data is manipulated.
'	This file includes the following functions: GetAddressDetails, GetCityState,
'		GetStoreByAddress, NormalizeAddress, LookupAddress, AddAddress,
'		GetStreetList, AddCASSAddress, GetCASSAddress, GetCASSAddresses,
'		DeleteCASSAddress, UpdateCASSAddress, UpdateAddressNotes,
'		GetStreetListByZipCodes
'
' Revision History:
' 7/19/201 - Created
' **************************************************************************

' **************************************************************************
' Function: GetAddressDetails
' Purpose: Retrieves address details.
' Parameters:	pnAddressID - The AddressID to search for
'				pnStoreID - The StoreID
'				psAddress1 - The address first line
'				psAddress2 - The address second line
'				psCity - The city
'				psState - The state
'				psPostalCode - The postal code
'				psAddressNotes - The address notes
' Return: True if sucessful, False if not
' **************************************************************************
Function GetAddressDetails(ByVal pnAddressID, ByRef pnStoreID, ByRef psAddress1, ByRef psAddress2, ByRef psCity, ByRef psState, ByRef psPostalCode, ByRef psAddressNotes)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	lsSQL = "select StoreID, AddressLine1, AddressLine2, City, State, PostalCode, AddressNotes from tblAddresses where AddressID = " & pnAddressID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lbRet = TRUE
			
			pnStoreID = loRS("StoreID")
			psAddress1 = Trim(loRS("AddressLine1"))
			If IsNull(loRS("AddressLine2")) Then
				psAddress2 = ""
			Else
				psAddress2 = Trim(loRS("AddressLine2"))
			End If
			psCity = Trim(loRS("City"))
			psState = Trim(loRS("State"))
			psPostalCode = Trim(loRS("PostalCode"))
			If IsNull(loRS("AddressNotes")) Then
				psAddressNotes = ""
			Else
				psAddressNotes = Trim(loRS("AddressNotes"))
			End If
		End If
		
		DBCloseQuery loRS
	End If
	
	GetAddressDetails = lbRet
End Function

' **************************************************************************
' Function: GetCityState
' Purpose: Retrieves the city and state of a postal code.
' Parameters:	psPostalCode - The postal code to search for
'				psCity - The city
'				psState - The state
' Return: True if sucessful, False if not
' **************************************************************************
Function GetCityState(ByVal psPostalCode, ByRef psCity, ByRef psState)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	lsSQL = "select distinct City, State from tblCASSAddresses where PostalCode = " & psPostalCode
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			psCity = Trim(loRS("City"))
			psState = Trim(loRS("State"))
		Else
			psCity = "UNKNOWN CITY"
			psState = "US"
		End If
		
		DBCloseQuery(loRS)
	Else
		psCity = "UNKNOWN CITY"
		psState = "US"
	End If
	
	GetCityState = lbRet
End Function

' **************************************************************************
' Function: GetStoreByAddress
' Purpose: Retrieves the store associated with an address. Also returns the
'			normalized address with apt/suite separated as well as the
'			delivery charge and driver money.
' Parameters:	psPostalCode - The postal code to search for
'				psAddress1 - The first address line
'				psAddress2 - The second address line
'				psCity - The city
'				psState - The state
'				pdDeliveryCharge - The delivery charge
'				pdDriverMoney - The driver money
' Return: The StoreID or 0 if not found
' **************************************************************************
Function GetStoreByAddress(ByVal psPostalCode, ByRef psAddress1, ByRef psAddress2, ByRef psCity, ByRef psState, ByRef pdDeliveryCharge, ByRef pdDriverMoney)
	Dim lnRet, lnAddress, lsStreet, lsApt, lsSQL, loRS
	
	lnRet = -1
	pdDeliveryCharge = 0.00
	pdDriverMoney = 0.00
	
	If IsNumeric(psPostalCode) Then
		If NormalizeAddress(psPostalCode, UCase(psAddress1), lnAddress, lsStreet, lsApt) Then
			' Try exact match
'			lsSQL = "select tblCASSAddresses.StoreID, tblCASSAddresses.City, tblCASSAddresses.State, tblCASSAddresses.DeliveryCharge, tblCASSAddresses.DriverMoney from tblCASSAddresses inner join tblStores on tblCASSAddresses.StoreID = tblStores.StoreID where tblStores.StoreID > 0 and tblCASSAddresses.PostalCode = '" & psPostalCode & "' and tblCASSAddresses.Street = '" & DBCleanLiteral(lsStreet) & "' and tblCASSAddresses.LowNumber <= " & lnAddress & " and tblCASSAddresses.HighNumber >= " & lnAddress
			lsSQL = "select tblCASSAddresses.StoreID, tblCASSAddresses.City, tblCASSAddresses.State, tblCASSAddresses.DeliveryCharge, tblCASSAddresses.DriverMoney from tblCASSAddresses inner join tblStores on tblCASSAddresses.StoreID = tblStores.StoreID where tblCASSAddresses.PostalCode = '" & psPostalCode & "' and tblCASSAddresses.Street = '" & DBCleanLiteral(lsStreet) & "' and tblCASSAddresses.LowNumber <= " & lnAddress & " and tblCASSAddresses.HighNumber >= " & lnAddress
			
			If DBOpenQuery(lsSQL, FALSE, loRS) Then
				If Not loRS.bof And Not loRS.eof Then
					' Put 1/2 apt back into address
					If lsApt = "1/2" Then
						psAddress1 = lnAddress & " 1/2 " & lsStreet
						psAddress2 = ""
					Else
						psAddress1 = lnAddress & " " & lsStreet
						psAddress2 = lsApt
					End If
					
					lnRet = loRS("StoreID")
					psCity = Trim(loRS("City"))
					psState = Trim(loRS("State"))
					pdDeliveryCharge = loRS("DeliveryCharge")
					pdDriverMoney = loRS("DriverMoney")
					
					DBCloseQuery(loRS)
				Else
					DBCloseQuery(loRS)
					
					' Try starting with replacing spaces with percent in case they left off the street type and with a direction
'					lsSQL = "select distinct tblCASSAddresses.StoreID, tblCASSAddresses.Street, tblCASSAddresses.City, tblCASSAddresses.State, tblCASSAddresses.DeliveryCharge, tblCASSAddresses.DriverMoney from tblCASSAddresses inner join tblStores on tblCASSAddresses.StoreID = tblStores.StoreID where tblStores.StoreID > 0 and tblCASSAddresses.PostalCode = '" & psPostalCode & "' and tblCASSAddresses.Street like '" & Replace(DBCleanLiteral(lsStreet), " ", "%") & "' and tblCASSAddresses.LowNumber <= " & lnAddress & " and tblCASSAddresses.HighNumber >= " & lnAddress
					lsSQL = "select distinct tblCASSAddresses.StoreID, tblCASSAddresses.Street, tblCASSAddresses.City, tblCASSAddresses.State, tblCASSAddresses.DeliveryCharge, tblCASSAddresses.DriverMoney from tblCASSAddresses inner join tblStores on tblCASSAddresses.StoreID = tblStores.StoreID where tblCASSAddresses.PostalCode = '" & psPostalCode & "' and tblCASSAddresses.Street like '" & Replace(DBCleanLiteral(lsStreet), " ", "%") & "' and tblCASSAddresses.LowNumber <= " & lnAddress & " and tblCASSAddresses.HighNumber >= " & lnAddress
					If DBOpenQuery(lsSQL, FALSE, loRS) Then
						If Not loRS.bof And Not loRS.eof Then
							loRS.MoveNext
							If Not loRS.eof Then
'								psAddress1 = UCase(psAddress1)
								' Put 1/2 apt back into address
								If lsApt = "1/2" Then
									psAddress1 = lnAddress & " 1/2 " & UCase(lsStreet)
									psAddress2 = ""
								Else
									psAddress1 = lnAddress & " " & UCase(lsStreet)
									psAddress2 = lsApt
								End If
								
								GetCityState psPostalCode, psCity, psState
							Else
								loRS.Requery
								
								' Put 1/2 apt back into address
								If lsApt = "1/2" Then
									psAddress1 = lnAddress & " 1/2 " & Trim(loRS("Street"))
									psAddress2 = ""
								Else
									psAddress1 = lnAddress & " " & Trim(loRS("Street"))
									psAddress2 = lsApt
								End If
								
								lnRet = loRS("StoreID")
								psCity = Trim(loRS("City"))
								psState = Trim(loRS("State"))
								pdDeliveryCharge = loRS("DeliveryCharge")
								pdDriverMoney = loRS("DriverMoney")
							End If
							
							DBCloseQuery(loRS)
						Else
							' Try starting with in case they left off the street type and without a direction
'							lsSQL = "select distinct tblCASSAddresses.StoreID, tblCASSAddresses.Street, tblCASSAddresses.City, tblCASSAddresses.State, tblCASSAddresses.DeliveryCharge, tblCASSAddresses.DriverMoney from tblCASSAddresses inner join tblStores on tblCASSAddresses.StoreID = tblStores.StoreID where tblStores.StoreID > 0 and tblCASSAddresses.PostalCode = '" & psPostalCode & "' and tblCASSAddresses.Street like '" & DBCleanLiteral(lsStreet) & " %' and tblCASSAddresses.LowNumber <= " & lnAddress & " and tblCASSAddresses.HighNumber >= " & lnAddress
							lsSQL = "select distinct tblCASSAddresses.StoreID, tblCASSAddresses.Street, tblCASSAddresses.City, tblCASSAddresses.State, tblCASSAddresses.DeliveryCharge, tblCASSAddresses.DriverMoney from tblCASSAddresses inner join tblStores on tblCASSAddresses.StoreID = tblStores.StoreID where tblCASSAddresses.PostalCode = '" & psPostalCode & "' and tblCASSAddresses.Street like '" & DBCleanLiteral(lsStreet) & " %' and tblCASSAddresses.LowNumber <= " & lnAddress & " and tblCASSAddresses.HighNumber >= " & lnAddress
							If DBOpenQuery(lsSQL, FALSE, loRS) Then
								If Not loRS.bof And Not loRS.eof Then
									loRS.MoveNext
									If Not loRS.eof Then
'										psAddress1 = UCase(psAddress1)
										' Put 1/2 apt back into address
										If lsApt = "1/2" Then
											psAddress1 = lnAddress & " 1/2 " & UCase(lsStreet)
											psAddress2 = ""
										Else
											psAddress1 = lnAddress & " " & UCase(lsStreet)
											psAddress2 = lsApt
										End If
										
										GetCityState psPostalCode, psCity, psState
									Else
										loRS.Requery
										
										' Put 1/2 apt back into address
										If lsApt = "1/2" Then
											psAddress1 = lnAddress & " 1/2 " & Trim(loRS("Street"))
											psAddress2 = ""
										Else
											psAddress1 = lnAddress & " " & Trim(loRS("Street"))
											psAddress2 = lsApt
										End If
										
										lnRet = loRS("StoreID")
										psCity = Trim(loRS("City"))
										psState = Trim(loRS("State"))
										pdDeliveryCharge = loRS("DeliveryCharge")
										pdDriverMoney = loRS("DriverMoney")
									End If
									
									DBCloseQuery(loRS)
								Else
									DBCloseQuery(loRS)
									
									' Try half the street in case they left off the street type with a direction
'									lsSQL = "select distinct tblCASSAddresses.StoreID, tblCASSAddresses.Street, tblCASSAddresses.City, tblCASSAddresses.State, tblCASSAddresses.DeliveryCharge, tblCASSAddresses.DriverMoney from tblCASSAddresses inner join tblStores on tblCASSAddresses.StoreID = tblStores.StoreID where tblStores.StoreID > 0 and tblCASSAddresses.PostalCode = '" & psPostalCode & "' and tblCASSAddresses.Street like '" & DBCleanLiteral(Left(lsStreet, (Len(lsStreet) / 2))) & "%' and tblCASSAddresses.LowNumber <= " & lnAddress & " and tblCASSAddresses.HighNumber >= " & lnAddress
									lsSQL = "select distinct tblCASSAddresses.StoreID, tblCASSAddresses.Street, tblCASSAddresses.City, tblCASSAddresses.State, tblCASSAddresses.DeliveryCharge, tblCASSAddresses.DriverMoney from tblCASSAddresses inner join tblStores on tblCASSAddresses.StoreID = tblStores.StoreID where tblCASSAddresses.PostalCode = '" & psPostalCode & "' and tblCASSAddresses.Street like '" & DBCleanLiteral(Left(lsStreet, (Len(lsStreet) / 2))) & "%' and tblCASSAddresses.LowNumber <= " & lnAddress & " and tblCASSAddresses.HighNumber >= " & lnAddress
									If DBOpenQuery(lsSQL, FALSE, loRS) Then
										If Not loRS.bof And Not loRS.eof Then
											loRS.MoveNext
											If Not loRS.eof Then
'												psAddress1 = UCase(psAddress1)
												' Put 1/2 apt back into address
												If lsApt = "1/2" Then
													psAddress1 = lnAddress & " 1/2 " & UCase(lsStreet)
													psAddress2 = ""
												Else
													psAddress1 = lnAddress & " " & UCase(lsStreet)
													psAddress2 = lsApt
												End If
												
												GetCityState psPostalCode, psCity, psState
											Else
												loRS.Requery
												
												' Put 1/2 apt back into address
												If lsApt = "1/2" Then
													psAddress1 = lnAddress & " 1/2 " & Trim(loRS("Street"))
													psAddress2 = ""
												Else
													psAddress1 = lnAddress & " " & Trim(loRS("Street"))
													psAddress2 = lsApt
												End If
												
												lnRet = loRS("StoreID")
												psCity = Trim(loRS("City"))
												psState = Trim(loRS("State"))
												pdDeliveryCharge = loRS("DeliveryCharge")
												pdDriverMoney = loRS("DriverMoney")
											End If
											
											DBCloseQuery(loRS)
										Else
'											psAddress1 = UCase(psAddress1)
											' Put 1/2 apt back into address
											If lsApt = "1/2" Then
												psAddress1 = lnAddress & " 1/2 " & UCase(lsStreet)
												psAddress2 = ""
											Else
												psAddress1 = lnAddress & " " & UCase(lsStreet)
												psAddress2 = lsApt
											End If
											
											GetCityState psPostalCode, psCity, psState
'											lnRet = -1
											lnRet = 0
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
	End If
	
	GetStoreByAddress = lnRet
End Function

' **************************************************************************
' Function: NormalizeAddress
' Purpose: Takes an address and separates the number, street, and apartment portions.
' Parameters:	psPostalCode - The postal code
'				psAddress - The address to normalize
'				pnAddress - The number portion
'				psStreet - The street portion
'				psApt - The apartment portion
' Return: True if sucessful, False if not
' **************************************************************************
Function NormalizeAddress(ByVal psPostalCode, ByVal psAddress, ByRef pnAddress, ByRef psStreet, ByRef psApt)
	Dim bRet, sRemain, aRemain, i, sDir, bCheck2, bDoAptSearch, j
	
	bRet = FALSE
	
	' Extract street number
	pnAddress = 0
	For i = 1 to Len(psAddress)
		If IsNumeric(Mid(psAddress, i, 1)) Then
			pnAddress = pnAddress * 10 + CInt(Mid(psAddress, i, 1))
		Else
			sRemain = Trim(psAddress)
			
			' Strip anything after a comma in case they put city/state
			If InStr(sRemain, ",") > 0 Then
				sRemain = Left(sRemain, (InStr(sRemain, ",") - 1))
			End If
			
			' Put a space after all periods in case it was forgotten
			sRemain = Trim(Replace(sRemain, ".", ". "))
			
			' Kill any periods
			sRemain = Trim(Replace(sRemain, ".", ""))
			
			aRemain = Split(sRemain)
			Exit For
		End If
	Next
	
	If Len(sRemain) > 0 Then
		psStreet = ""
		sDir = ""
		psApt = ""
		
		' Hack for APARTMENT DR WEST in 48144
		bDoAptSearch = TRUE
		If psPostalCode = "48144" And UBound(aRemain) >= 3 Then
			If aRemain(1) = "APARTMENT" Then
				bDoAptSearch = FALSE
				
				If Len(aRemain(UBound(aRemain))) > 1 And Left(aRemain(UBound(aRemain)), 1) = "#" Then
					psApt = Mid(aRemain(UBound(aRemain)), 2)
					ReDim Preserve aRemain(UBound(aRemain)-1)
				Else
					If aRemain(UBound(aRemain) - 1) = "#" Then
						psApt = aRemain(UBound(aRemain))
						ReDim Preserve aRemain(UBound(aRemain)-2)
					End If
				End If
			End If
		End If
		
		If bDoAptSearch Then
			' Look backwards for Apt/Suite
			For i = UBound(aRemain) To 1 Step -1
				If aRemain(i) = "APT" Or aRemain(i) = "APT." Or aRemain(i) = "APARTMENT" Or aRemain(i) = "APART" Or aRemain(i) = "APART." Or aRemain(i) = "#" Or Left(aRemain(i), 1) = "#" Or aRemain(i) = "STE" Or aRemain(i) = "STE." Or aRemain(i) = "SUITE" Or aRemain(i) = "BLDG" Or aRemain(i) = "BLDG." Or aRemain(i) = "BUILDING" Or aRemain(i) = "LOT" Or aRemain(i) = "FLR" Or aRemain(i) = "FLR." Or aRemain(i) = "FLOOR" Or aRemain(i) = "UNIT" or aRemain(i) = "ROOM" or aRemain(i) = "RM" or aRemain(i) = "RM." Then
					If Len(aRemain(i)) > 1 And Left(aRemain(i), 1) = "#" Then
						psApt = Mid(aRemain(i), 2)
					Else
						psApt = aRemain(i+1)
					End If
					ReDim Preserve aRemain(i-1)
					Exit For
				End If
			Next
			
			' Check for double Apt/Suite (i.e. "Apt #")
			i = UBound(aRemain)
			Do While (i > 0) And (Len(Trim(aRemain(i))) = 0)
				i = i - 1
			Loop
			If i > 0 Then
				If aRemain(i) = "APT" Or aRemain(i) = "APT." Or aRemain(i) = "APARTMENT" Or aRemain(i) = "APART" Or aRemain(i) = "APART." Or aRemain(i) = "#" Or Left(aRemain(i), 1) = "#" Or aRemain(i) = "STE" Or aRemain(i) = "STE." Or aRemain(i) = "SUITE" Or aRemain(i) = "BLDG" Or aRemain(i) = "BLDG." Or aRemain(i) = "BUILDING" Or aRemain(i) = "LOT" Or aRemain(i) = "FLR" Or aRemain(i) = "FLR." Or aRemain(i) = "FLOOR" Or aRemain(i) = "UNIT" or aRemain(i) = "ROOM" or aRemain(i) = "RM" or aRemain(i) = "RM." Then
					ReDim Preserve aRemain(i-1)
				End If
			End If
		End If
		
		IF UBound(aRemain) > 0 Then
			' THE FOLLOWING CODE BREAKS A LOT OF ADDRESSES
'			For i = 1 To UBound(aRemain)
'				If Len(Trim(aRemain(i))) > 0 Then
'					' Check for initial direction
'					Select Case aRemain(i)
'						case "N", "N.", "NORTH"
'							If Len(sDir) = 0 Then
'								sDir = "NORTH"
'							Else
'								psStreet = psStreet & " " & aRemain(i)
'							End If
'						case "S", "S.", "SOUTH"
'							If Len(sDir) = 0 Then
'								sDir = "SOUTH"
'							Else
'								psStreet = psStreet & " " & aRemain(i)
'							End If
'						case "E", "E.", "EAST"
'							If Len(sDir) = 0 Then
'								sDir = "EAST"
'							Else
'								psStreet = psStreet & " " & aRemain(i)
'							End If
'						case "W", "W.", "WEST"
'							If Len(sDir) = 0 Then
'								sDir = "WEST"
'							Else
'								psStreet = psStreet & " " & aRemain(i)
'							End If
'						case else
'							psStreet = psStreet & " " & aRemain(i)
'					End Select
'				End If
'			Next
			
' 2013-06-17 TAM: Deal with 1/2 before others
			' Account for 1/2
			For i = 1 To UBound(aRemain)
				If aRemain(i) = "1/2" Then
					If Len(psApt) = 0 Then
						psApt = "1/2"
					End If
					
					For j = (i + 1) to UBound(aRemain)
						aRemain(j - 1) = aRemain(j)
					Next
					ReDim Preserve aRemain(UBound(aRemain) - 1)
					Exit For
				End If
			Next
			
			' Prevent looking for direction if 2 words and last is type of road
			bCheck2 = TRUE
			If UBound(aRemain) = 2 Then
				Select Case aRemain(UBound(aRemain))
					Case "RD", "ROAD", "DR", "DRIVE", "CT", "COURT", "PL", "PLACE", "BLVD", "BOULEVARD", "BL.", "BL", "BOLEVARD", "ST", "STREET", "AVE", "AVENUE", "AV.", "AV", "LN", "LANE", "CIR", "CIRCLE", "TRCE", "TRACE", "TRL", "TRAIL", "CROSSING"
						bCheck2 = FALSE
				End Select
			End If
			
			' THE FOLLOWING CODE BREAKS "N LOCKWOOD" TYPE ADDRESSES - BUT THERE WAS A REASON THIS WAS DONE - WHY?
			If (UBound(aRemain) > 2) Or (UBound(aRemain) = 2 AND bCheck2) Then
				' Check for trailing direction
				Select Case aRemain(UBound(aRemain))
					case "N", "NORTH"
						sDir = "NORTH"
					case "S", "SOUTH"
						sDir = "SOUTH"
					case "E", "EAST"
						sDir = "EAST"
					case "W", "WEST"
						sDir = "WEST"
				End Select
				If Len(sDir) > 0 Then
					For i = 1 to (UBound(aRemain) - 1)
						psStreet = psStreet & " " & aRemain(i)
					Next
				Else
					' Check for initial direction
					Select Case aRemain(1)
						case "N", "NORTH"
							sDir = "NORTH"
						case "S", "SOUTH"
							sDir = "SOUTH"
						case "E", "EAST"
							sDir = "EAST"
						case "W", "WEST"
							sDir = "WEST"
					End Select
					If Len(sDir) > 0 Then
						For i = 2 to UBound(aRemain)
							psStreet = psStreet & " " & aRemain(i)
						Next
					Else
						For i = 1 to UBound(aRemain)
							psStreet = psStreet & " " & aRemain(i)
						Next
					End If
				End If
			Else
				For i = 1 to UBound(aRemain)
					psStreet = psStreet & " " & aRemain(i)
				Next
			End If
			
			aRemain = Split(psStreet)
			
			' Check for FIRST, SECOND, THIRD, etc.
			Select Case aRemain(1)
				Case "FIRST"
					aRemain(1) = "1ST"
				Case "SECOND"
					aReamin(1) = "2ND"
				Case "THIRD"
					aRemain(1) = "3RD"
				Case "FOURTH"
					aRemain(1) = "4TH"
				Case "FIFTH"
					aRemain(1) = "5TH"
				Case "SIXTH"
					aRemain(1) = "6TH"
				Case "SEVENTH"
					aRemain(1) = "7TH"
				Case "EIGHTH"
					aRemain(1) = "8TH"
				Case "NINTH"
					aRemain(1) = "9TH"
				Case "TENTH"
					aRemain(1) = "10TH"
				Case "ELEVENTH"
					aRemain(1) = "11TH"
				Case "TWELVETH"
					aReamin(1) = "12TH"
				Case "THIRTEENTH"
					aRemain(1) = "13TH"
				Case "FOURTEENTH"
					aRemain(1) = "14TH"
				Case "FIFTEENTH"
					aRemain(1) = "15TH"
				Case "SIXTEENTH"
					aReamin(1) = "16TH"
				Case "SEVENTEENTH"
					aRemain(1) = "17TH"
				Case "EIGHTEENTH"
					aRemain(1) = "18TH"
				Case "NINTEENTH"
					aRemain(1) = "19TH"
				Case "TWENTYETH"
					aRemain(1) = "20TH"
			End Select
			
			' Check end for street type and correct
			Select Case aRemain(UBound(aRemain))
				Case "RD", "ROAD"
					aRemain(UBound(aRemain)) = "RD"
				Case "DR", "DRIVE"
					aRemain(UBound(aRemain)) = "DR"
				Case "CT", "COURT"
					aRemain(UBound(aRemain)) = "CT"
				Case "PL", "PLACE"
					aRemain(UBound(aRemain)) = "PL"
				Case "BLVD", "BOULEVARD", "BL.", "BL", "BOLEVARD"
					aRemain(UBound(aRemain)) = "BLVD"
				Case "ST", "STREET"
					aRemain(UBound(aRemain)) = "ST"
				Case "AVE", "AVENUE", "AV.", "AV"
					aRemain(UBound(aRemain)) = "AVE"
				Case "LN", "LANE"
					aRemain(UBound(aRemain)) = "LN"
				Case "CIR", "CIRCLE"
					aRemain(UBound(aRemain)) = "CIR"
				Case "TRCE", "TRACE"
					aRemain(UBound(aRemain)) = "TRCE"
				Case "TRL", "TRAIL"
					aRemain(UBound(aRemain)) = "TRL"
				Case "CROSSING"
					aRemain(UBound(aRemain)) = "XING"
				Case "TERRACE"
					aRemain(UBound(aRemain)) = "TER"
			End Select
			
			' Rebuild Street
			psStreet = ""
			For i = 1 To UBound(aRemain)
' 2013-06-17 TAM: Deal with 1/2 above
'				' Account for 1/2
'				If aRemain(i) = "1/2" and Len(psApt) = 0 Then
'					psApt = "1/2"
'				Else
					psStreet = psStreet + " " + aRemain(i)
'				End If
			Next
			
			' Replace hyphens with spaces
			psStreet = Replace(psStreet, "-", " ")
			
			' Move direction to end
			psStreet = Trim(psStreet + " " + sDir)
			
			' Static Conversions
			If Left(psStreet, 4) = "TWP " Then
				psStreet = "TOWNSHIP " & Mid(psStreet, 5)
			End If
			If Left(psStreet, 12) = "TOWNSHIP RD " Then
				psStreet = "TOWNSHIP ROAD " & Mid(psStreet, 13)
			End If
			If Left(psStreet, 3) = "ST " Then
				' Must differentiate between SAINT and STATE ROUTE
				If Left(psStreet, 6) = "ST RT " Or Left(psStreet, 7) = "ST RTE " Or Left(psStreet, 9) = "ST ROUTE " Then
					psStreet = "STATE " & Mid(psStreet, 4)
				Else
					psStreet = "SAINT " & Mid(psStreet, 4)
				End If
			End If
			If Left(psStreet, 9) = "STATE RT " Then
				psStreet = "STATE ROUTE " & Mid(psStreet, 10)
			End If
			If Left(psStreet, 10) = "STATE RTE " Then
				psStreet = "STATE ROUTE " & Mid(psStreet, 11)
			End If
			If Left(psStreet, 10) = "COUNTY RD " Then
				psStreet = "COUNTY ROAD " & Mid(psStreet, 11)
			End If
			If psPostalCode = "43612" And (psStreet = "TOWNE MALL DR NORTH" or psStreet = "TOWNE MALL NORTH") Then
				psStreet = "NORTH TOWNE MALL DR"
			End If
			If psPostalCode = "45840" And (psStreet = "POINT DR EAST" Or psStreet = "POINT EAST") Then
				psStreet = "EAST POINT DR"
			End If
			If psPostalCode = "45840" And (psStreet = "POINT DR NORTH" Or psStreet = "POINT NORTH") Then
				psStreet = "NORTH POINT DR"
			End If
			If psPostalCode = "45840" And (psStreet = "POINT DR SOUTH" Or psStreet = "POINT SOUTH") Then
				psStreet = "SOUTH POINT DR"
			End If
			If psPostalCode = "45840" And psStreet = "NORTH WEST" Then
				psStreet = "WEST ST NORTH"
			End If
			If psPostalCode = "43611" And psStreet = "TER" Then
				psStreet = "TERRACE DR"
			End If
			If psPostalCode = "43528" And psStreet = "OAK TER" Then
				psStreet = "OAK TERRACE BLVD"
			End If
			If psPostalCode = "43551" And psStreet = "FIVE POINT RD" Then
				psStreet = "5 POINT RD"
			End If
			If psPostalCode = "48162" And (psStreet = "AVENUE DE LAYFAYETTE" Or psStreet = "AVENUE DELAYFAYETTE" Or psStreet = "AVE DELAYFAYETTE") Then
				psStreet = "AVE DE LAYFAYETTE"
			End If
			
			bRet = TRUE
		End If
	End If
	
	NormalizeAddress = bRet
End Function

' **************************************************************************
' Function: LookupAddress
' Purpose: Looks up an address.
' Parameters:	psAddress1 - The address first line to find
'				psAddress2 - The address second line to find
'				psCity - The city to find
'				psState - The state to find
'				psPostalCode - The postal code to find
'				pnAddressID - The AddressID
'				pnStoreID - The StoreID
'				psAddressNotes - The address notes
' Return: True if sucessful, False if not
' **************************************************************************
Function LookupAddress(ByVal psAddress1, ByVal psAddress2, ByVal psCity, ByVal psState, ByVal psPostalCode, ByRef pnAddressID, ByRef pnStoreID, ByRef psAddressNotes)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	lsSQL = "select AddressID, StoreID, AddressNotes from tblAddresses where AddressLine1 = '" & DBCleanLiteral(psAddress1) & "' and "
	If Len(psAddress2) = 0 Then
		lsSQL = lsSQL & "(AddressLine2 = '' Or AddressLine2 is NULL) "
	Else
		lsSQL = lsSQL & "AddressLine2 = '" & DBCleanLiteral(psAddress2) & "' "
	End If
	lsSQL = lsSQL & "and City = '" & DBCleanLiteral(psCity) & "' and State = '" & DBCleanLiteral(psState) & "' and PostalCode = '" & DBCleanLiteral(psPostalCode) & "'"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			pnAddressID = loRS("AddressID")
			pnStoreID = loRS("StoreID")
			If IsNull(loRS("AddressNotes")) Then
				psAddressNotes = ""
			Else
				psAddressNotes = Trim(loRS("AddressNotes"))
			End If
		Else
			pnAddressID = 0
			pnStoreID = 0
			psAddressNotes = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	LookupAddress = lbRet
End Function

' **************************************************************************
' Function: AddAddress
' Purpose: Adds a new address.
' Parameters:	pnStoreID - The StoreID to add
'				psAddress1 - The address first line to add
'				psAddress2 - The address second line to add
'				psCity - The city to add
'				psState - The state to add
'				psPostalCode - The postal code to add
'				psAddressNotes - The address notes to add
'				pbIsManual - Is this a manual add
' Return: The new AddressID
' **************************************************************************
Function AddAddress(ByVal pnStoreID, ByVal psAddress1, ByVal psAddress2, ByVal psCity, ByVal psState, ByVal psPostalCode, ByVal psAddressNotes, ByVal pbIsManual)
	Dim lnRet, lsSQL, loRS
	
	lnRet = 0
	
	lsSQL = "EXEC AddAddress @pStoreID = " & pnStoreID & ", @pAddressLine1 = '" & DBCleanLiteral(psAddress1) & "', @pAddressLine2 = "
'	If Len(psAddress2) = 0 Then
'		lsSQL = lsSQL & "NULL"
'	Else
		lsSQL = lsSQL & "'" & DBCleanLiteral(psAddress2) & "'"
'	End If
	lsSQL = lsSQL & ", @pCity = '" & DBCleanLiteral(psCity) & "', @pState = '" & DBCleanLiteral(psState) & "', @pPostalCode = '" & DBCleanLiteral(psPostalCode) & "', @pAddressNotes = "
	If Len(psAddressNotes) = 0 Then
		lsSQL = lsSQL & "NULL"
	Else
		lsSQL = lsSQL & "'" & DBCleanLiteral(psAddressNotes) & "'"
	End If
	If pbIsManual Then
		lsSQL = lsSQL & ", @pIsManual = 1"
	Else
		lsSQL = lsSQL & ", @pIsManual = 0"
	End If
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lnRet = CLng(loRS(0))
		End If
		
		DBCloseQuery loRS
	End If
	
	AddAddress = lnRet
End Function

' **************************************************************************
' Function: GetStreetList
' Purpose: Gets a list of street names.
' Parameters:	pnPostalCode - The postal code to search for
'				pnStreetNumber - The street number to search for
'				psStreetChar - The first character of the street
'				pasStreet - Array of street names
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStreetList(ByVal pnPostalCode, ByVal pnStreetNumber, ByVal psStreetChar, ByRef pasStreet)
	Dim lbRet, lsEOCode, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	If pnStreetNumber Mod 2 = 1 Then
		lsEOCode = "O"
	Else
		lsEOCode = "E"
	End If
	
	lsSQL = "select distinct street from tblCASSAddresses where PostalCode = " & pnPostalCode & " and LowNumber <= " & pnStreetNumber & " and HighNumber >= " & pnStreetNumber & " and (EOCode = 'B' Or EOCode = '" & lsEOCode & "') and Street like '" & psStreetChar & "%' order by Street"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve pasStreet(lnPos)
				
				pasStreet(lnPos) = Trim(loRS("Street"))
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim pasStreet(0)
			pasStreet(0) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStreetList = lbRet
End Function

' **************************************************************************
' Function: AddCASSAddress
' Purpose: Adds a new CASS address.
' Parameters:	pnStoreID - The StoreID to add
'				psPostalCode - The postal code to add
'				psStreet - The street to add
'				pnLowNumber - The low number to add
'				pnHighNumber - The high number to add
'				psEOCode - The EO code to add
'				psCity - The city to add
'				psState - The state to add
'				pdDeliveryCharge - The delivery charge to add
'				psDriverMoney - The driver money to add
' Return: The new CASSAddressID
' **************************************************************************
Function AddCASSAddress(ByVal pnStoreID, ByVal psPostalCode, ByVal psStreet, ByVal pnLowNumber, ByVal pnHighNumber, ByVal psEOCode, ByVal psCity, ByVal psState, ByVal pdDeliveryCharge, ByVal pdDriverMoney)
	Dim lnRet, lsSQL, loRS
	
	lnRet = 0
	
	lsSQL = "EXEC AddCASSAddress @pStoreID = " & pnStoreID & ", @pPostalCode = '" & DBCleanLiteral(psPostalCode) & "', @pStreet = '" & DBCleanLiteral(psStreet) & "'"
	lsSQL = lsSQL & ", @pLowNumber = " & pnLowNumber & ", @pHighNumber = " & pnHighNumber & ", @pEOCode = '" & DBCleanLiteral(psEOCode) & "'"
	lsSQL = lsSQL & ", @pCity = '" & DBCleanLiteral(psCity) & "', @pState = '" & DBCleanLiteral(psState) & "', @pDeliveryCharge = " & pdDeliveryCharge & ", @pDriverMoney = " & pdDriverMoney
	
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lnRet = CLng(loRS(0))
		End If
		
		DBCloseQuery loRS
	End If
	
	AddCASSAddress = lnRet
End Function

' **************************************************************************
' Function: GetCASSAddress
' Purpose: Retrieves a CASS address.
' Parameters:	pnCASSAddressID - The CASSAddressID to search for
'				pnStoreID - The StoreID
'				psPostalCode - The postal code
'				psStreet - The street
'				pnLowNumber - The low number
'				pnHighNumber - The high number
'				psEOCode - The EO code
'				psCity - The city
'				psState - The state
'				pdDeliveryCharge - The delivery charge
'				psDriverMoney - The driver money
' Return: True if sucessful, False if not
' **************************************************************************
Function GetCASSAddress(ByRef pnCASSAddressID, ByRef pnStoreID, ByRef psPostalCode, ByRef psStreet, ByRef pnLowNumber, ByRef pnHighNumber, ByRef psEOCode, ByRef psCity, ByRef psState, ByRef pdDeliveryCharge, ByRef pdDriverMoney)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	lsSQL = "select StoreID, PostalCode, Street, LowNumber, HighNumber, EOCode, City, State, DeliveryCharge, DriverMoney from tblCASSAddresses where CASSAddressID = " & pnCASSAddressID
	
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			pnStoreID = loRS("StoreID")
			psPostalCode = loRS("PostalCode")
			psStreet = loRS("Street")
			pnLowNumber = loRS("LowNumber")
			pnHighNumber = loRS("HighNumber")
			psEOCode = loRS("EOCode")
			psCity = loRS("City")
			psState = loRS("State")
			pdDeliveryCharge = loRS("DeliveryCharge")
			pdDriverMoney = loRS("DriverMoney")
		Else
			pnCASSAddressID = 0
		End If
		
		DBCloseQuery loRS
	End If
	
	GetCASSAddress = lbRet
End Function

' **************************************************************************
' Function: GetCASSAddresses
' Purpose: Retrieves a list of CASS address based on search criteria.
' Parameters:	psStoreID - The store ID to search for
'				psPostalCode - The postal code to search for
'				psStreet - The street to search for
'				psCity - The city to search for
'				psState - The state to search for
'				panCASSAddressID - Array of CASSAddressIDs
'				panStoreID - Array of StoreIDs
'				pasPostalCode - Array of postal codes
'				pasStreet - Array of streets
'				panLowNumber - Array of low numbers
'				panHighNumber - Array of high numbers
'				pasEOCode - Array of EO codes
'				pasCity - Array of citys
'				pasState - Array of states
'				padDeliveryCharge - Array of delivery charges
'				pasDriverMoney - Array of driver monies
' Return: True if sucessful, False if not
' **************************************************************************
Function GetCASSAddresses(ByVal psStoreID, ByVal psPostalCode, ByVal psStreet, ByVal psCity, ByVal psState, ByRef panCASSAddressID, ByRef panStoreID, ByRef pasPostalCode, ByRef pasStreet, ByRef panLowNumber, ByRef panHighNumber, ByRef pasEOCode, ByRef pasCity, ByRef pasState, ByRef padDeliveryCharge, ByRef padDriverMoney)
	Dim lbRet, lsSQL, lsWhere, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select CASSAddressID, StoreID, PostalCode, Street, LowNumber, HighNumber, EOCode, City, State, DeliveryCharge, DriverMoney from tblCASSAddresses"
	lsWhere = ""
	If Len(psStoreID) <> 0 Then
		If IsNumeric(psStoreID) Then
			lsWhere = "where StoreID = " & psStoreID
		End If
	End If
	If Len(psPostalCode) <> 0 Then
		If Len(lsWhere) = 0 Then
			lsWhere = "where "
		Else
			lsWhere = lsWhere & " and "
		End If
		
		lsWhere = lsWhere & "PostalCode = '" & DBCleanLiteral(psPostalCode) & "'"
	End If
	If Len(psStreet) <> 0 Then
		If Len(lsWhere) = 0 Then
			lsWhere = "where "
		Else
			lsWhere = lsWhere & " and "
		End If
		
		lsWhere = lsWhere & "Street like '%" & DBCleanLiteral(psStreet) & "%'"
	End If
	If Len(psCity) <> 0 Then
		If Len(lsWhere) = 0 Then
			lsWhere = "where "
		Else
			lsWhere = lsWhere & " and "
		End If
		
		lsWhere = lsWhere & "City = '" & DBCleanLiteral(psCity) & "'"
	End If
	If Len(psState) <> 0 Then
		If Len(lsWhere) = 0 Then
			lsWhere = "where "
		Else
			lsWhere = lsWhere & " and "
		End If
		
		lsWhere = lsWhere & "State = '" & DBCleanLiteral(psState) & "'"
	End If
	If Len(lsWhere) <> 0 Then
		lsSQL = lsSQL & " " & lsWhere
	End If
	lsSQL = lsSQL & " order by Street, LowNumber"
	
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve panCASSAddressID(lnPos), panStoreID(lnPos), pasPostalCode(lnPos), pasStreet(lnPos), panLowNumber(lnPos), panHighNumber(lnPos), pasEOCode(lnPos), pasCity(lnPos), pasState(lnPos), padDeliveryCharge(lnPos), padDriverMoney(lnPos)
				
				panCASSAddressID(lnPos) = loRS("CASSAddressID")
				panStoreID(lnPos) = loRS("StoreID")
				pasPostalCode(lnPos) = loRS("PostalCode")
				pasStreet(lnPos) = loRS("Street")
				panLowNumber(lnPos) = loRS("LowNumber")
				panHighNumber(lnPos) = loRS("HighNumber")
				pasEOCode(lnPos) = loRS("EOCode")
				pasCity(lnPos) = loRS("City")
				pasState(lnPos) = loRS("State")
				padDeliveryCharge(lnPos) = loRS("DeliveryCharge")
				padDriverMoney(lnPos) = loRS("DriverMoney")
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panCASSAddressID(0), panStoreID(0), pasPostalCode(0), pasStreet(0), panLowNumber(0), panHighNumber(0), pasEOCode(0), pasCity(0), pasState(0), padDeliveryCharge(0), padDriverMoney(0)
			
			panCASSAddressID(lnPos) = 0
			panStoreID(lnPos) = 0
			pasPostalCode(lnPos) = ""
			pasStreet(lnPos) = ""
			panLowNumber(lnPos) = 0
			panHighNumber(lnPos) = 0
			pasEOCode(lnPos) = ""
			pasCity(lnPos) = ""
			pasState(lnPos) = ""
			padDeliveryCharge(lnPos) = 0.00
			padDriverMoney(lnPos) = 0.00
		End If
		
		DBCloseQuery loRS
	End If
	
	GetCASSAddresses = lbRet
End Function

' **************************************************************************
' Function: DeleteCASSAddress
' Purpose: Deletes a CASS address.
' Parameters:	pnCASSAddressID - The CASSAddressID
' Return: True if sucessful, False if not
' **************************************************************************
Function DeleteCASSAddress(ByVal pnCASSAddressID)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	lsSQL = "delete from tblCASSAddresses where CASSAddressID = " & pnCASSAddressID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If
	
	DeleteCASSAddress = lbRet
End Function

' **************************************************************************
' Function: UpdateCASSAddress
' Purpose: Updates a CASS address.
' Parameters:	pnCASSAddressID - The CASSAddressID
' Return: True if sucessful, False if not
' **************************************************************************
Function UpdateCASSAddress(ByVal pnCASSAddressID, ByVal pnStoreID, ByVal psPostalCode, ByVal psStreet, ByVal pnLowNumber, ByVal pnHighNumber, ByVal psEOCode, ByVal psCity, ByVal psState, ByVal pdDeliveryCharge, ByVal pdDriverMoney)
	Dim lbRet, lsSQL
	
	lbRet = FALSE
	
	lsSQL = "update tblCASSAddresses set StoreID = " & pnStoreID & ", PostalCode = '" & DBCleanLiteral(psPostalCode) & "', Street = '" & DBCleanLiteral(psStreet) & "'"
	lsSQL = lsSQL & ", LowNumber = " & pnLowNumber & ", HighNumber = " & pnHighNumber & ", EOCode = '" & DBCleanLiteral(psEOCode) & "'"
	lsSQL = lsSQL & ", City = '" & DBCleanLiteral(psCity) & "', State = '" & DBCleanLiteral(psState) & "', DeliveryCharge = " & pdDeliveryCharge & ", DriverMoney = " & pdDriverMoney
	lsSQL = lsSQL & " where CASSAddressID = " & pnCASSAddressID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If
	
	UpdateCASSAddress = lbRet
End Function

' **************************************************************************
' Function: UpdateAddressNotes
' Purpose: Updates address notes.
' Parameters:	pnAddressID - The AddressID to update
'				psAddressNotes - The address notes
' Return: True if sucessful, False if not
' **************************************************************************
Function UpdateAddressNotes(ByVal pnAddressID, ByVal psAddressNotes)
	Dim lbRet, lsSQL, loRS
	
	lbRet = FALSE
	
	lsSQL = "update tblAddresses set AddressNotes = '" & DBCleanLiteral(psAddressNotes) & "' where AddressID = " & pnAddressID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If
	
	UpdateAddressNotes = lbRet
End Function

' **************************************************************************
' Function: GetStreetListByZipCodes
' Purpose: Gets a list of street names.
' Parameters:	pasSearchPostalCodes - The postal codes to search for
'				pnStreetNumber - The street number to search for
'				psStreetChar - The first character of the street
'				pasStreet - Array of street names
'				pasPostalCodes - Array of postal codes
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStreetListByZipCodes(ByVal pasSearchPostalCodes, ByVal pnStreetNumber, ByVal psStreetChar, ByRef pasStreet, ByRef pasPostalCode)
	Dim lbRet, lsEOCode, lsSQL, loRS, lnPos, i
	
	lbRet = FALSE
	
	If pnStreetNumber Mod 2 = 1 Then
		lsEOCode = "O"
	Else
		lsEOCode = "E"
	End If
	
	lsSQL = "select distinct street, postalcode from tblCASSAddresses where ("
	For i = 0 To UBound(pasSearchPostalCodes)
		If i > 0 Then
			lsSQL = lsSQL & " or "
		End If
		lsSQL = lsSQL & "PostalCode = '" & pasSearchPostalCodes(i) & "'"
	Next
	lsSQL = lsSQL & ") and LowNumber <= " & pnStreetNumber & " and HighNumber >= " & pnStreetNumber & " and (EOCode = 'B' Or EOCode = '" & lsEOCode & "') and Street like '" & psStreetChar & "%' order by Street"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve pasStreet(lnPos), pasPostalCode(lnPos)
				
				pasStreet(lnPos) = Trim(loRS("Street"))
				pasPostalCode(lnPos) = Trim(loRS("PostalCode"))
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim pasStreet(0), pasPostalCode(0)
			
			pasStreet(0) = ""
			pasPostalCode(0) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStreetListByZipCodes = lbRet
End Function
%>