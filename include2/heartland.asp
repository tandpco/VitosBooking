<%
' **************************************************************************
' File: heartland.asp
' Purpose: Functions for heartland payment gateway.
' Created: 9/22/2011 - TAM
' Description:
'	Include this file on any page where credit cards are processed.
'	This file includes the following functions: CCAuth, CCSale, CCForce, 
'		CCTipAdjust, CCSettle, CCVoid, CCWebOrderForce, CCWebOrderVoid, 
'		CCWebSettle
'
' Revision History:
' 7/19/201 - Created
' **************************************************************************

' **************************************************************************
' Function: CCAuth
' Purpose: Process a credit card authorization.
' Parameters:	pnStoreID - The StoreID
'				psCardNum - The card number or track data
'				psExpDate - The expiration date (if not sending track data)
'				psName - The cardholder name
'				psAddress - The cardholder address
'				psZip - The cardholder zip code
'				pnOrderID - The order ID
'				pnCustomerID - The customer ID
'				pdAmount - The amount of the authorization (grand total)
'				psReference - The reference number returned, "DECLINED" if
'					declined or "ERROR X" if some other error (X is error #).
' Return: True if sucessful, False if not
' **************************************************************************
Function CCAuth(ByVal pnStoreID, ByVal psCardNum, ByVal psExpDate, ByVal psName, ByVal psAddress, ByVal psZip, ByVal pnOrderID, ByVal pnCustomerID, ByVal pdAmount, ByRef psReference)
	Dim lbRet, lsSQL, loRS, lsPGWID, loXMLHttp, lsURL, lsData, loXMLDoc, lsTrack1, lsTrack2
	
	lbRet = FALSE
	
	lsSQL = "select PGWID from tblStores where StoreID = " & pnStoreID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			If Not IsNull(loRS("PGWID")) Then
				lsPGWID = Trim(loRS("PGWID"))
				If Len(lsPGWID) > 0 Then
					lsURL = gsPaymentGatewayURL
					
					lsData = "UserName=" & Server.URLEncode(lsPGWID)
					lsData = lsData & "&Password=" & Server.URLEncode(gsPaymentGatewayPW)
					lsData = lsData & "&TransType=Auth"
					lsData = lsData & "&Amount=" & pdAmount
					lsData = lsData & "&InvNum=" & pnOrderID
					lsData = lsData & "&PNRef="
					lsData = lsData & "&CVNum="
					
					If Len(psCardNum) > 16 Then
						lsTrack1 = Left(psCardNum, (InStr(psCardNum, "?;") - 1))
						lsTrack2 = Mid(psCardNum, (InStr(psCardNum, "?;") + 2))
						lsTrack2 = Left(lsTrack2, (Len(lsTrack2) - 1))
						
						lsData = lsData & "&CardNum=" & Left(lsTrack2, (InStr(lsTrack2, "=") - 1))
						lsData = lsData & "&ExpDate=" & Mid(lsTrack2, (InStr(lsTrack2, "=") + 3), 2) & Mid(lsTrack2, (InStr(lsTrack2, "=") + 1), 2)
						lsData = lsData & "&MagData=" & lsTrack2
						lsData = lsData & "&NameOnCard=" & Mid(lsTrack1, (InStr(lsTrack1, "^") + 1), (InStrRev(lsTrack1, "^") - (InStr(lsTrack1, "^") + 1)))
						lsData = lsData & "&Zip="
						lsData = lsData & "&Street="
						lsData = lsData & "&ExtData=" & Server.URLEncode("<CustomerID>" & pnCustomerID & "</CustomerID><PONum>" & pnOrderID & "</PONum><CVPresence>2</CVPresence><Presentation><CardPresent>True</CardPresent></Presentation><EntryMode>MAGNETICSTRIPE</EntryMode>")
					Else
						lsData = lsData & "&CardNum=" & psCardNum
						lsData = lsData & "&ExpDate=" & psExpDate
						lsData = lsData & "&MagData="
						lsData = lsData & "&NameOnCard=" & Server.URLEncode(psName)
						lsData = lsData & "&Zip=" & psZip
						lsData = lsData & "&Street=" & Server.URLEncode(psAddress)
						lsData = lsData & "&ExtData=" & Server.URLEncode("<CustomerID>" & pnCustomerID & "</CustomerID><PONum>" & pnOrderID & "</PONum><CVPresence>2</CVPresence><Presentation><CardPresent>False</CardPresent></Presentation><EntryMode>MANUAL</EntryMode>")
					End If
					
					On Error Resume Next
					Set loXMLHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
					If Err.Number = 0 Then
						lsURL = lsURL & "?" & lsData
						loXMLHttp.Open "GET", lsURL, false
						If Err.Number = 0 Then
							loXMLHttp.send
							If Err.Number = 0 Then
								Set loXMLDoc = Server.CreateObject("Msxml2.DOMDocument")
								If Err.Number = 0 Then
									loXMLDoc.loadXML(loXMLHttp.responseText)
									If loXMLDoc.parseError.errorCode = 0 Then
										lbRet = TRUE
										
										' NOTE: Result 0 is approved, 12 or 13 is declined, else is error
										If CLng(loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue) = 0 Then
											psReference = loXMLDoc.getElementsByTagName("PNRef")(0).childNodes(0).nodeValue
										Else
											If CLng(loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue) = 12 Or CLng(loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue) = 13 Then
												psReference = "DECLINED"
											Else
'												psReference = "ERROR " & loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue & " - " & loXMLDoc.getElementsByTagName("RespMSG")(0).childNodes(0).nodeValue & " - " & loXMLDoc.getElementsByTagName("Message")(0).childNodes(0).nodeValue
												psReference = "ERROR " & loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue & " - " & loXMLDoc.getElementsByTagName("RespMSG")(0).childNodes(0).nodeValue
											End If
										End If
									Else
										gsDBErrorMessage = "Unable to parse the response from the payment gateway."
									End If
								Else
									gsDBErrorMessage = Err.Description
								End If
							Else
								gsDBErrorMessage = Err.Description
							End If
						Else
							gsDBErrorMessage = Err.Description
						End If
					Else
						gsDBErrorMessage = Err.Description
					End If
				Else
					gsDBErrorMessage = "No payment gateway User ID has been defined for this store."
				End If
			Else
				gsDBErrorMessage = "No payment gateway User ID has been defined for this store."
			End If
		Else
			gsDBErrorMessage = "The specified store was not found."
		End If
		
		DBCloseQuery loRS
	End If
	
	CCAuth = lbRet
End Function

' **************************************************************************
' Function: CCSale
' Purpose: Process a credit card sale.
' Parameters:	pnStoreID - The StoreID
'				psCardNum - The card number or track data
'				psExpDate - The expiration date (if not sending track data)
'				psName - The cardholder name
'				psAddress - The cardholder address
'				psZip - The cardholder zip code
'				pnOrderID - The order ID
'				pnCustomerID - The customer ID
'				pdAmount - The amount of the sale (grand total)
'				pdTax - The tax amount
'				pdTip - The tip amount
'				psReference - The reference number returned, "DECLINED" if
'					declined or "ERROR X" if some other error (X is error #).
' Return: True if sucessful, False if not
' **************************************************************************
Function CCSale(ByVal pnStoreID, ByVal psCardNum, ByVal psExpDate, ByVal psName, ByVal psAddress, ByVal psZip, ByVal pnOrderID, ByVal pnCustomerID, ByVal pdAmount, ByVal pdTax, ByVal pdTip, ByRef psReference)
	Dim lbRet, lsSQL, loRS, lsPGWID, loXMLHttp, lsURL, lsData, loXMLDoc, lsTrack1, lsTrack2
	
	lbRet = FALSE
	
	lsSQL = "select PGWID from tblStores where StoreID = " & pnStoreID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			If Not IsNull(loRS("PGWID")) Then
				lsPGWID = Trim(loRS("PGWID"))
				If Len(lsPGWID) > 0 Then
					lsURL = gsPaymentGatewayURL
					
					lsData = "UserName=" & Server.URLEncode(lsPGWID)
					lsData = lsData & "&Password=" & Server.URLEncode(gsPaymentGatewayPW)
					lsData = lsData & "&TransType=Sale"
					lsData = lsData & "&Amount=" & pdAmount
					lsData = lsData & "&InvNum=" & pnOrderID
					lsData = lsData & "&PNRef="
					lsData = lsData & "&CVNum="
					
					If Len(psCardNum) > 16 Then
						lsTrack1 = Left(psCardNum, (InStr(psCardNum, "?;") - 1))
						lsTrack2 = Mid(psCardNum, (InStr(psCardNum, "?;") + 2))
						lsTrack2 = Left(lsTrack2, (Len(lsTrack2) - 1))
						
						lsData = lsData & "&CardNum=" & Left(lsTrack2, (InStr(lsTrack2, "=") - 1))
						lsData = lsData & "&ExpDate=" & Mid(lsTrack2, (InStr(lsTrack2, "=") + 3), 2) & Mid(lsTrack2, (InStr(lsTrack2, "=") + 1), 2)
						lsData = lsData & "&MagData=" & lsTrack2
						lsData = lsData & "&NameOnCard=" & Mid(lsTrack1, (InStr(lsTrack1, "^") + 1), (InStrRev(lsTrack1, "^") - (InStr(lsTrack1, "^") + 1)))
						lsData = lsData & "&Zip="
						lsData = lsData & "&Street="
						lsData = lsData & "&ExtData=" & Server.URLEncode("<TipAmt>" & pdTip & "</TipAmt><TaxAmt>" & pdTax & "</TaxAmt><CustomerID>" & pnCustomerID & "</CustomerID><PONum>" & pnOrderID & "</PONum><CVPresence>2</CVPresence><Presentation><CardPresent>True</CardPresent></Presentation><EntryMode>MAGNETICSTRIPE</EntryMode>")
					Else
						lsData = lsData & "&CardNum=" & psCardNum
						lsData = lsData & "&ExpDate=" & psExpDate
						lsData = lsData & "&MagData="
						lsData = lsData & "&NameOnCard=" & Server.URLEncode(psName)
						lsData = lsData & "&Zip=" & psZip
						lsData = lsData & "&Street=" & Server.URLEncode(psAddress)
						lsData = lsData & "&ExtData=" & Server.URLEncode("<TipAmt>" & pdTip & "</TipAmt><TaxAmt>" & pdTax & "</TaxAmt><CustomerID>" & pnCustomerID & "</CustomerID><PONum>" & pnOrderID & "</PONum><CVPresence>2</CVPresence><Presentation><CardPresent>False</CardPresent></Presentation><EntryMode>MANUAL</EntryMode>")
					End If
					
					On Error Resume Next
					Set loXMLHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
					If Err.Number = 0 Then
						lsURL = lsURL & "?" & lsData
						loXMLHttp.Open "GET", lsURL, false
						If Err.Number = 0 Then
							loXMLHttp.send
							If Err.Number = 0 Then
								Set loXMLDoc = Server.CreateObject("Msxml2.DOMDocument")
								If Err.Number = 0 Then
									loXMLDoc.loadXML(loXMLHttp.responseText)
									If loXMLDoc.parseError.errorCode = 0 Then
										lbRet = TRUE
										
										' NOTE: Result 0 is approved, 12 or 13 is declined, else is error
										If CLng(loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue) = 0 Then
											psReference = loXMLDoc.getElementsByTagName("PNRef")(0).childNodes(0).nodeValue
										Else
											If CLng(loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue) = 12 Or CLng(loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue) = 13 Then
												psReference = "DECLINED"
											Else
'												psReference = "ERROR " & loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue & " - " & loXMLDoc.getElementsByTagName("RespMSG")(0).childNodes(0).nodeValue & " - " & loXMLDoc.getElementsByTagName("Message")(0).childNodes(0).nodeValue
												psReference = "ERROR " & loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue & " - " & loXMLDoc.getElementsByTagName("RespMSG")(0).childNodes(0).nodeValue
											End If
										End If
									Else
										gsDBErrorMessage = "Unable to parse the response from the payment gateway."
									End If
								Else
									gsDBErrorMessage = Err.Description
								End If
							Else
								gsDBErrorMessage = Err.Description
							End If
						Else
							gsDBErrorMessage = Err.Description
						End If
					Else
						gsDBErrorMessage = Err.Description
					End If
				Else
					gsDBErrorMessage = "No payment gateway User ID has been defined for this store."
				End If
			Else
				gsDBErrorMessage = "No payment gateway User ID has been defined for this store."
			End If
		Else
			gsDBErrorMessage = "The specified store was not found."
		End If
		
		DBCloseQuery loRS
	End If
	
	CCSale = lbRet
End Function

' **************************************************************************
' Function: CCForce
' Purpose: Converts a previous authorization into a sale.
' Parameters:	pnOrderID - The OrderID of the previous authorization
'				psPNRefNum - The reference number of the previous authorization.
'				psReference - The reference number returned, "DECLINED" if
'					declined or "ERROR X" if some other error (X is error #).
' Return: True if sucessful, False if not
' **************************************************************************
Function CCForce(ByVal pnOrderID, ByVal psPNRefNum, ByRef psReference)
	Dim lbRet, lsSQL, loRS, lsPGWID, loXMLHttp, ldTotal, ldDelivery, ldTax, ldTip, lsURL, lsData, loXMLDoc
	
	lbRet = FALSE
	
	lsSQL = "select DeliveryCharge, Tax, Tax2, Tip, Quantity, Cost, Discount, PGWID from tblOrders inner join tblOrderLines on tblOrders.OrderID = tblOrderLines.OrderID inner join tblStores on tblOrders.StoreID = tblStores.StoreID where tblOrders.OrderID = " & pnOrderID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			If Not IsNull(loRS("PGWID")) Then
				lsPGWID = Trim(loRS("PGWID"))
				If Len(lsPGWID) > 0 Then
					ldTotal = 0.00
					ldDelivery = loRS("DeliveryCharge")
					ldTax = loRS("Tax") + loRS("Tax2")
					ldTip = loRS("Tip")
					
					Do While Not loRS.eof
						ldTotal = ldTotal + (loRS("Quantity") * (loRS("Cost") - loRS("Discount")))
						
						loRS.MoveNext
					Loop
					ldTotal = ldTotal + ldDelivery + ldTax + ldTip
					
					lsURL = gsPaymentGatewayURL
					
					lsData = "UserName=" & Server.URLEncode(lsPGWID)
					lsData = lsData & "&Password=" & Server.URLEncode(gsPaymentGatewayPW)
					lsData = lsData & "&TransType=Force"
					lsData = lsData & "&Amount=" & ldTotal
					lsData = lsData & "&InvNum=" & pnOrderID
					lsData = lsData & "&PNRef=" & Server.URLEncode(psPNRefNum)
					lsData = lsData & "&CVNum="
					lsData = lsData & "&CardNum="
					lsData = lsData & "&ExpDate="
					lsData = lsData & "&MagData="
					lsData = lsData & "&NameOnCard="
					lsData = lsData & "&Zip="
					lsData = lsData & "&Street="
					lsData = lsData & "&ExtData="
					If ldTip > 0 Then
						lsData = lsData & Server.URLEncode("<TipAmt>" & ldTip & "</TipAmt>")
					End If
					If ldTax > 0 Then
						lsData = lsData & Server.URLEncode("<TaxAmt>" & ldTax & "</TaxAmt>")
					End If
					
					On Error Resume Next
					Set loXMLHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
					If Err.Number = 0 Then
						lsURL = lsURL & "?" & lsData
						loXMLHttp.Open "GET", lsURL, false
						If Err.Number = 0 Then
							loXMLHttp.send
							If Err.Number = 0 Then
								Set loXMLDoc = Server.CreateObject("Msxml2.DOMDocument")
								If Err.Number = 0 Then
									loXMLDoc.loadXML(loXMLHttp.responseText)
									If loXMLDoc.parseError.errorCode = 0 Then
										lbRet = TRUE
										
										' NOTE: Result 0 is approved, 12 or 13 is declined, else is error
										If CLng(loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue) = 0 Then
											psReference = loXMLDoc.getElementsByTagName("PNRef")(0).childNodes(0).nodeValue
											
											lsSQL = "update tblOrders set PaymentReference = '" & psReference & "' where OrderID = " & pnOrderID
											DBExecuteSQL lsSQL
										Else
											If CLng(loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue) = 12 Or CLng(loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue) = 13 Then
												psReference = "DECLINED"
											Else
												psReference = "ERROR " & loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue & " - " & loXMLDoc.getElementsByTagName("RespMSG")(0).childNodes(0).nodeValue & " - " & loXMLDoc.getElementsByTagName("Message")(0).childNodes(0).nodeValue
											End If
										End If
									Else
										gsDBErrorMessage = "Unable to parse the response from the payment gateway."
									End If
								Else
									gsDBErrorMessage = Err.Description
								End If
							Else
								gsDBErrorMessage = Err.Description
							End If
						Else
							gsDBErrorMessage = Err.Description
						End If
					Else
						gsDBErrorMessage = Err.Description
					End If
				Else
					gsDBErrorMessage = "No payment gateway User ID has been defined for this store."
				End If
			Else
				gsDBErrorMessage = "No payment gateway User ID has been defined for this store."
			End If
		Else
			gsDBErrorMessage = "The specified store was not found."
		End If
		
		DBCloseQuery loRS
	End If
	
	CCForce = lbRet
End Function

' **************************************************************************
' Function: CCTipAdjust
' Purpose: Adjusts the tip on a previous sale.
' Parameters:	pnStoreID - The StoreID
'				psPNRefNum - The reference number of the previous sale
'				pdTip - Amount of tip
'				psReference - The reference number returned, "DECLINED" if
'					declined or "ERROR X" if some other error (X is error #).
' Return: True if sucessful, False if not
' **************************************************************************
Function CCTipAdjust(ByVal pnStoreID, ByVal psPNRefNum, ByVal pdTip, ByRef psReference)
	Dim lbRet, lsSQL, loRS, lsPGWID, loXMLHttp, lsURL, lsData, loXMLDoc
	
	lbRet = FALSE
	
	lsSQL = "select PGWID from tblStores where StoreID = " & pnStoreID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			If Not IsNull(loRS("PGWID")) Then
				lsPGWID = Trim(loRS("PGWID"))
				If Len(lsPGWID) > 0 Then
					lsURL = gsPaymentGatewayURL
					
					lsData = "UserName=" & Server.URLEncode(lsPGWID)
					lsData = lsData & "&Password=" & Server.URLEncode(gsPaymentGatewayPW)
					lsData = lsData & "&TransType=Adjustment"
					lsData = lsData & "&Amount="
					lsData = lsData & "&InvNum="
					lsData = lsData & "&PNRef=" & Server.URLEncode(psPNRefNum)
					lsData = lsData & "&CVNum="
					lsData = lsData & "&CardNum="
					lsData = lsData & "&ExpDate="
					lsData = lsData & "&MagData="
					lsData = lsData & "&NameOnCard="
					lsData = lsData & "&Zip="
					lsData = lsData & "&Street="
					lsData = lsData & "&ExtData=" & Server.URLEncode("<TipAmt>" & pdTip & "</TipAmt>")
					
					On Error Resume Next
					Set loXMLHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
					If Err.Number = 0 Then
						lsURL = lsURL & "?" & lsData
						loXMLHttp.Open "GET", lsURL, false
						If Err.Number = 0 Then
							loXMLHttp.send
							If Err.Number = 0 Then
								Set loXMLDoc = Server.CreateObject("Msxml2.DOMDocument")
								If Err.Number = 0 Then
									loXMLDoc.loadXML(loXMLHttp.responseText)
									If loXMLDoc.parseError.errorCode = 0 Then
										lbRet = TRUE
										
										' NOTE: Result 0 is approved, 12 or 13 is declined, else is error
										If CLng(loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue) = 0 Then
											psReference = loXMLDoc.getElementsByTagName("PNRef")(0).childNodes(0).nodeValue
										Else
											If CLng(loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue) = 12 Or CLng(loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue) = 13 Then
												psReference = "DECLINED"
											Else
												psReference = "ERROR " & loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue & " - " & loXMLDoc.getElementsByTagName("RespMSG")(0).childNodes(0).nodeValue & " - " & loXMLDoc.getElementsByTagName("Message")(0).childNodes(0).nodeValue
											End If
										End If
									Else
										gsDBErrorMessage = "Unable to parse the response from the payment gateway."
									End If
								Else
									gsDBErrorMessage = Err.Description
								End If
							Else
								gsDBErrorMessage = Err.Description
							End If
						Else
							gsDBErrorMessage = Err.Description
						End If
					Else
						gsDBErrorMessage = Err.Description
					End If
				Else
					gsDBErrorMessage = "No payment gateway User ID has been defined for this store."
				End If
			Else
				gsDBErrorMessage = "No payment gateway User ID has been defined for this store."
			End If
		Else
			gsDBErrorMessage = "The specified store was not found."
		End If
		
		DBCloseQuery loRS
	End If
	
	CCTipAdjust = lbRet
End Function

' **************************************************************************
' Function: CCSettle
' Purpose: Closes out the current batch.
' Parameters:	pnStoreID - The StoreID
'				pnCount - The number of transaction in the batch
'				pdTotal - The total amount settled
'				pnBatchNumber - The batch number
'				psReference - The reference number returned, "DECLINED" if
'					declined or "ERROR X" if some other error (X is error #).
' Return: True if sucessful, False if not
' **************************************************************************
Function CCSettle(ByVal pnStoreID, ByRef pnCount, ByRef pdTotal, ByRef pnBatchNumber, ByRef psReference)
	Dim lbRet, lsSQL, loRS, lsPGWID, loXMLHttp, lsURL, lsData, loXMLDoc, laExtra, i
	
	lbRet = FALSE
	
	pnCount = 0
	pdTotal = 0.00
	pnBatchNumber = ""
	
	lsSQL = "select PGWID from tblStores where StoreID = " & pnStoreID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			If Not IsNull(loRS("PGWID")) Then
				lsPGWID = Trim(loRS("PGWID"))
				If Len(lsPGWID) > 0 Then
					lsURL = gsPaymentGatewayURL
					
					lsData = "UserName=" & Server.URLEncode(lsPGWID)
					lsData = lsData & "&Password=" & Server.URLEncode(gsPaymentGatewayPW)
					lsData = lsData & "&TransType=CaptureAll"
					lsData = lsData & "&Amount="
					lsData = lsData & "&InvNum="
					lsData = lsData & "&PNRef="
					lsData = lsData & "&CVNum="
					lsData = lsData & "&CardNum="
					lsData = lsData & "&ExpDate="
					lsData = lsData & "&MagData="
					lsData = lsData & "&NameOnCard="
					lsData = lsData & "&Zip="
					lsData = lsData & "&Street="
					lsData = lsData & "&ExtData="
					
					On Error Resume Next
					Set loXMLHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
					If Err.Number = 0 Then
						lsURL = lsURL & "?" & lsData
						loXMLHttp.Open "GET", lsURL, false
						If Err.Number = 0 Then
							loXMLHttp.send
							If Err.Number = 0 Then
								Set loXMLDoc = Server.CreateObject("Msxml2.DOMDocument")
								If Err.Number = 0 Then
									loXMLDoc.loadXML(loXMLHttp.responseText)
									If loXMLDoc.parseError.errorCode = 0 Then
										lbRet = TRUE
										
										' NOTE: Result 0 is approved, 12 or 13 is declined, else is error
										If CLng(loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue) = 0 Then
											psReference = loXMLDoc.getElementsByTagName("AuthCode")(0).childNodes(0).nodeValue
											
											laExtra = Split(loXMLDoc.getElementsByTagName("ExtData")(0).childNodes(0).nodeValue, ",")
											pnCount = 0
											pdTotal = 0.00
											pnBatchNumber = 0
											For i = 0 To UBound(laExtra)
												If UCase(Left(laExtra(i), 10)) = "NET_COUNT=" Then
													pnCount = CLng(Mid(laExtra(i), 11))
												Else
													If UCase(Left(laExtra(i), 11)) = "NET_AMOUNT=" Then
														pdTotal = CDbl(Mid(laExtra(i), 12))
													Else
														If UCase(Left(laExtra(i), 7)) = "NUMBER=" Then
															pnBatchNumber = Mid(laExtra(i), 8)
														End If
													End If
												End If
											Next
										Else
											If CLng(loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue) = 12 Or CLng(loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue) = 13 Then
												psReference = "DECLINED"
											Else
												psReference = "ERROR " & loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue & " - " & loXMLDoc.getElementsByTagName("RespMSG")(0).childNodes(0).nodeValue & " - " & loXMLDoc.getElementsByTagName("Message")(0).childNodes(0).nodeValue
											End If
										End If
									Else
										gsDBErrorMessage = "Unable to parse the response from the payment gateway."
									End If
								Else
									gsDBErrorMessage = Err.Description
								End If
							Else
								gsDBErrorMessage = Err.Description
							End If
						Else
							gsDBErrorMessage = Err.Description
						End If
					Else
						gsDBErrorMessage = Err.Description
					End If
				Else
					gsDBErrorMessage = "No payment gateway User ID has been defined for this store."
				End If
			Else
				gsDBErrorMessage = "No payment gateway User ID has been defined for this store."
			End If
		Else
			gsDBErrorMessage = "The specified store was not found."
		End If
		
		DBCloseQuery loRS
	End If
	
	CCSettle = lbRet
End Function

' **************************************************************************
' Function: CCVoid
' Purpose: Voids a previous sale.
' Parameters:	pnStoreID - The StoreID
'				psPNRefNum - The reference number of the previous sale
'				psReference - The reference number returned, "DECLINED" if
'					declined or "ERROR X" if some other error (X is error #).
' Return: True if sucessful, False if not
' **************************************************************************
Function CCVoid(ByVal pnStoreID, ByVal psPNRefNum, ByRef psReference)
	Dim lbRet, lsSQL, loRS, lsPGWID, loXMLHttp, lsURL, lsData, loXMLDoc
	
	lbRet = FALSE
	
	lsSQL = "select PGWID from tblStores where StoreID = " & pnStoreID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			If Not IsNull(loRS("PGWID")) Then
				lsPGWID = Trim(loRS("PGWID"))
				If Len(lsPGWID) > 0 Then
					lsURL = gsPaymentGatewayURL
					
					lsData = "UserName=" & Server.URLEncode(lsPGWID)
					lsData = lsData & "&Password=" & Server.URLEncode(gsPaymentGatewayPW)
					lsData = lsData & "&TransType=Void"
					lsData = lsData & "&Amount="
					lsData = lsData & "&InvNum="
					lsData = lsData & "&PNRef=" & Server.URLEncode(psPNRefNum)
					lsData = lsData & "&CVNum="
					lsData = lsData & "&CardNum="
					lsData = lsData & "&ExpDate="
					lsData = lsData & "&MagData="
					lsData = lsData & "&NameOnCard="
					lsData = lsData & "&Zip="
					lsData = lsData & "&Street="
					lsData = lsData & "&ExtData="
					
					On Error Resume Next
					Set loXMLHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
					If Err.Number = 0 Then
						lsURL = lsURL & "?" & lsData
						loXMLHttp.Open "GET", lsURL, false
						If Err.Number = 0 Then
							loXMLHttp.send
							If Err.Number = 0 Then
								Set loXMLDoc = Server.CreateObject("Msxml2.DOMDocument")
								If Err.Number = 0 Then
									loXMLDoc.loadXML(loXMLHttp.responseText)
									If loXMLDoc.parseError.errorCode = 0 Then
										lbRet = TRUE
										
										' NOTE: Result 0 is approved, 12 or 13 is declined, else is error
										If CLng(loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue) = 0 Then
											psReference = loXMLDoc.getElementsByTagName("PNRef")(0).childNodes(0).nodeValue
										Else
											If CLng(loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue) = 12 Or CLng(loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue) = 13 Then
												psReference = "DECLINED"
											Else
												psReference = "ERROR " & loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue & " - " & loXMLDoc.getElementsByTagName("RespMSG")(0).childNodes(0).nodeValue & " - " & loXMLDoc.getElementsByTagName("Message")(0).childNodes(0).nodeValue
											End If
										End If
									Else
										gsDBErrorMessage = "Unable to parse the response from the payment gateway."
									End If
								Else
									gsDBErrorMessage = Err.Description
								End If
							Else
								gsDBErrorMessage = Err.Description
							End If
						Else
							gsDBErrorMessage = Err.Description
						End If
					Else
						gsDBErrorMessage = Err.Description
					End If
				Else
					gsDBErrorMessage = "No payment gateway User ID has been defined for this store."
				End If
			Else
				gsDBErrorMessage = "No payment gateway User ID has been defined for this store."
			End If
		Else
			gsDBErrorMessage = "The specified store was not found."
		End If
		
		DBCloseQuery loRS
	End If
	
	CCVoid = lbRet
End Function

' **************************************************************************
' Function: CCWebOrderForce
' Purpose: Converts a previous authorization into a sale.
' Parameters:	pnOrderID - The OrderID of the previous authorization
'				psPNRefNum - The reference number of the previous authorization.
'				psReference - The reference number returned, "DECLINED" if
'					declined or "ERROR X" if some other error (X is error #).
' Return: True if sucessful, False if not
' **************************************************************************
Function CCWebOrderForce(ByVal pnOrderID, ByVal psPNRefNum, ByRef psReference)
	Dim lbRet, lsSQL, loRS, lsPGWID, loXMLHttp, ldTotal, ldDelivery, ldTax, ldTip, lsURL, lsData, loXMLDoc
	
	lbRet = FALSE
	
	lsSQL = "select DeliveryCharge, Tax, Tax2, Tip, Quantity, Cost, Discount, PGWWebID2 from tblOrders inner join tblOrderLines on tblOrders.OrderID = tblOrderLines.OrderID inner join tblStores on tblOrders.StoreID = tblStores.StoreID where tblOrders.OrderID = " & pnOrderID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			If Not IsNull(loRS("PGWWebID2")) Then
				lsPGWID = Trim(loRS("PGWWebID2"))
				If Len(lsPGWID) > 0 Then
					ldTotal = 0.00
					ldDelivery = loRS("DeliveryCharge")
					ldTax = loRS("Tax") + loRS("Tax2")
					ldTip = loRS("Tip")
					
					Do While Not loRS.eof
						ldTotal = ldTotal + (loRS("Quantity") * (loRS("Cost") - loRS("Discount")))
						
						loRS.MoveNext
					Loop
					ldTotal = ldTotal + ldDelivery + ldTax + ldTip
					
					lsURL = gsPaymentGatewayURL
					
					lsData = "UserName=" & Server.URLEncode(lsPGWID)
					lsData = lsData & "&Password=" & Server.URLEncode(gsPaymentGatewayPW)
					lsData = lsData & "&TransType=Force"
					lsData = lsData & "&Amount=" & ldTotal
					lsData = lsData & "&InvNum=" & pnOrderID
					lsData = lsData & "&PNRef=" & Server.URLEncode(psPNRefNum)
					lsData = lsData & "&CVNum="
					lsData = lsData & "&CardNum="
					lsData = lsData & "&ExpDate="
					lsData = lsData & "&MagData="
					lsData = lsData & "&NameOnCard="
					lsData = lsData & "&Zip="
					lsData = lsData & "&Street="
					lsData = lsData & "&ExtData="
					If ldTip > 0 Then
						lsData = lsData & Server.URLEncode("<TipAmt>" & ldTip & "</TipAmt>")
					End If
					If ldTax > 0 Then
						lsData = lsData & Server.URLEncode("<TaxAmt>" & ldTax & "</TaxAmt>")
					End If
					
					On Error Resume Next
					Set loXMLHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
					If Err.Number = 0 Then
						lsURL = lsURL & "?" & lsData
						loXMLHttp.Open "GET", lsURL, false
						If Err.Number = 0 Then
							loXMLHttp.send
							If Err.Number = 0 Then
								Set loXMLDoc = Server.CreateObject("Msxml2.DOMDocument")
								If Err.Number = 0 Then
									loXMLDoc.loadXML(loXMLHttp.responseText)
									If loXMLDoc.parseError.errorCode = 0 Then
										lbRet = TRUE
										
										' NOTE: Result 0 is approved, 12 or 13 is declined, else is error
										If CLng(loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue) = 0 Then
											psReference = loXMLDoc.getElementsByTagName("PNRef")(0).childNodes(0).nodeValue
											
'											lsSQL = "update tblOrders set PaidDate = GetDate(), IsPaid = 1, PaymentReference = '" & psReference & "', PaymentEmpID = 1 where OrderID = " & pnOrderID
'											DBExecuteSQL lsSQL
										Else
											If CLng(loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue) = 12 Or CLng(loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue) = 13 Then
												psReference = "DECLINED"
											Else
												psReference = "ERROR " & loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue & " - " & loXMLDoc.getElementsByTagName("RespMSG")(0).childNodes(0).nodeValue & " - " & loXMLDoc.getElementsByTagName("Message")(0).childNodes(0).nodeValue
											End If
										End If
									Else
										gsDBErrorMessage = "Unable to parse the response from the payment gateway."
									End If
								Else
									gsDBErrorMessage = Err.Description
								End If
							Else
								gsDBErrorMessage = Err.Description
							End If
						Else
							gsDBErrorMessage = Err.Description
						End If
					Else
						gsDBErrorMessage = Err.Description
					End If
				Else
					gsDBErrorMessage = "No payment gateway User ID has been defined for this store."
				End If
			Else
				gsDBErrorMessage = "No payment gateway User ID has been defined for this store."
			End If
		Else
			gsDBErrorMessage = "The specified store was not found."
		End If
		
		DBCloseQuery loRS
	End If
	
	CCWebOrderForce = lbRet
End Function

' **************************************************************************
' Function: CCWebOrderVoid
' Purpose: Voids a previous web sale.
' Parameters:	pnStoreID - The StoreID
'				psPNRefNum - The reference number of the previous sale
'				psReference - The reference number returned, "DECLINED" if
'					declined or "ERROR X" if some other error (X is error #).
' Return: True if sucessful, False if not
' **************************************************************************
Function CCWebOrderVoid(ByVal pnStoreID, ByVal psPNRefNum, ByRef psReference)
	Dim lbRet, lsSQL, loRS, lsPGWID, loXMLHttp, lsURL, lsData, loXMLDoc
	
	lbRet = FALSE
	
	lsSQL = "select PGWWebID2 from tblStores where StoreID = " & pnStoreID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			If Not IsNull(loRS("PGWWebID2")) Then
				lsPGWID = Trim(loRS("PGWWebID2"))
				If Len(lsPGWID) > 0 Then
					lsURL = gsPaymentGatewayURL
					
					lsData = "UserName=" & Server.URLEncode(lsPGWID)
					lsData = lsData & "&Password=" & Server.URLEncode(gsPaymentGatewayPW)
					lsData = lsData & "&TransType=Void"
					lsData = lsData & "&Amount="
					lsData = lsData & "&InvNum="
					lsData = lsData & "&PNRef=" & Server.URLEncode(psPNRefNum)
					lsData = lsData & "&CVNum="
					lsData = lsData & "&CardNum="
					lsData = lsData & "&ExpDate="
					lsData = lsData & "&MagData="
					lsData = lsData & "&NameOnCard="
					lsData = lsData & "&Zip="
					lsData = lsData & "&Street="
					lsData = lsData & "&ExtData="
					
					On Error Resume Next
					Set loXMLHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
					If Err.Number = 0 Then
						lsURL = lsURL & "?" & lsData
						loXMLHttp.Open "GET", lsURL, false
						If Err.Number = 0 Then
							loXMLHttp.send
							If Err.Number = 0 Then
								Set loXMLDoc = Server.CreateObject("Msxml2.DOMDocument")
								If Err.Number = 0 Then
									loXMLDoc.loadXML(loXMLHttp.responseText)
									If loXMLDoc.parseError.errorCode = 0 Then
										lbRet = TRUE
										
										' NOTE: Result 0 is approved, 12 or 13 is declined, else is error
										If CLng(loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue) = 0 Then
											psReference = loXMLDoc.getElementsByTagName("PNRef")(0).childNodes(0).nodeValue
										Else
											If CLng(loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue) = 12 Or CLng(loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue) = 13 Then
												psReference = "DECLINED"
											Else
												psReference = "ERROR " & loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue & " - " & loXMLDoc.getElementsByTagName("RespMSG")(0).childNodes(0).nodeValue & " - " & loXMLDoc.getElementsByTagName("Message")(0).childNodes(0).nodeValue
											End If
										End If
									Else
										gsDBErrorMessage = "Unable to parse the response from the payment gateway."
									End If
								Else
									gsDBErrorMessage = Err.Description
								End If
							Else
								gsDBErrorMessage = Err.Description
							End If
						Else
							gsDBErrorMessage = Err.Description
						End If
					Else
						gsDBErrorMessage = Err.Description
					End If
				Else
					gsDBErrorMessage = "No payment gateway User ID has been defined for this store."
				End If
			Else
				gsDBErrorMessage = "No payment gateway User ID has been defined for this store."
			End If
		Else
			gsDBErrorMessage = "The specified store was not found."
		End If
		
		DBCloseQuery loRS
	End If
	
	CCWebOrderVoid = lbRet
End Function

' **************************************************************************
' Function: CCWebSettle
' Purpose: Closes out the current web batch.
' Parameters:	pnStoreID - The StoreID
'				pnCount - The number of transaction in the batch
'				pdTotal - The total amount settled
'				pnBatchNumber - The batch number
'				psReference - The reference number returned, "DECLINED" if
'					declined or "ERROR X" if some other error (X is error #).
' Return: True if sucessful, False if not
' **************************************************************************
Function CCWebSettle(ByVal pnStoreID, ByRef pnCount, ByRef pdTotal, ByRef pnBatchNumber, ByRef psReference)
	Dim lbRet, lsSQL, loRS, lsPGWID, loXMLHttp, lsURL, lsData, loXMLDoc, laExtra, i
	
	lbRet = FALSE
	
	pnCount = 0
	pdTotal = 0.00
	pnBatchNumber = ""
	
	lsSQL = "select PGWWebID2 from tblStores where StoreID = " & pnStoreID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			If Not IsNull(loRS("PGWWebID2")) Then
				lsPGWID = Trim(loRS("PGWWebID2"))
				If Len(lsPGWID) > 0 Then
					lsURL = gsPaymentGatewayURL
					
					lsData = "UserName=" & Server.URLEncode(lsPGWID)
					lsData = lsData & "&Password=" & Server.URLEncode(gsPaymentGatewayPW)
					lsData = lsData & "&TransType=CaptureAll"
					lsData = lsData & "&Amount="
					lsData = lsData & "&InvNum="
					lsData = lsData & "&PNRef="
					lsData = lsData & "&CVNum="
					lsData = lsData & "&CardNum="
					lsData = lsData & "&ExpDate="
					lsData = lsData & "&MagData="
					lsData = lsData & "&NameOnCard="
					lsData = lsData & "&Zip="
					lsData = lsData & "&Street="
					lsData = lsData & "&ExtData="
					
					On Error Resume Next
					Set loXMLHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
					If Err.Number = 0 Then
						lsURL = lsURL & "?" & lsData
						loXMLHttp.Open "GET", lsURL, false
						If Err.Number = 0 Then
							loXMLHttp.send
							If Err.Number = 0 Then
								Set loXMLDoc = Server.CreateObject("Msxml2.DOMDocument")
								If Err.Number = 0 Then
									loXMLDoc.loadXML(loXMLHttp.responseText)
									If loXMLDoc.parseError.errorCode = 0 Then
										lbRet = TRUE
										
										' NOTE: Result 0 is approved, 12 or 13 is declined, else is error
										If CLng(loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue) = 0 Then
											psReference = loXMLDoc.getElementsByTagName("AuthCode")(0).childNodes(0).nodeValue
											
											laExtra = Split(loXMLDoc.getElementsByTagName("ExtData")(0).childNodes(0).nodeValue, ",")
											pnCount = 0
											pdTotal = 0.00
											pnBatchNumber = 0
											For i = 0 To UBound(laExtra)
												If UCase(Left(laExtra(i), 10)) = "NET_COUNT=" Then
													pnCount = CLng(Mid(laExtra(i), 11))
												Else
													If UCase(Left(laExtra(i), 11)) = "NET_AMOUNT=" Then
														pdTotal = CDbl(Mid(laExtra(i), 12))
													Else
														If UCase(Left(laExtra(i), 7)) = "NUMBER=" Then
															pnBatchNumber = Mid(laExtra(i), 8)
														End If
													End If
												End If
											Next
										Else
											If CLng(loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue) = 12 Or CLng(loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue) = 13 Then
												psReference = "DECLINED"
											Else
												psReference = "ERROR " & loXMLDoc.getElementsByTagName("Result")(0).childNodes(0).nodeValue & " - " & loXMLDoc.getElementsByTagName("RespMSG")(0).childNodes(0).nodeValue & " - " & loXMLDoc.getElementsByTagName("Message")(0).childNodes(0).nodeValue
											End If
										End If
									Else
										gsDBErrorMessage = "Unable to parse the response from the payment gateway."
									End If
								Else
									gsDBErrorMessage = Err.Description
								End If
							Else
								gsDBErrorMessage = Err.Description
							End If
						Else
							gsDBErrorMessage = Err.Description
						End If
					Else
						gsDBErrorMessage = Err.Description
					End If
				Else
					gsDBErrorMessage = "No payment gateway User ID has been defined for this store."
				End If
			Else
				gsDBErrorMessage = "No payment gateway User ID has been defined for this store."
			End If
		Else
			gsDBErrorMessage = "The specified store was not found."
		End If
		
		DBCloseQuery loRS
	End If
	
	CCWebSettle = lbRet
End Function
%>
