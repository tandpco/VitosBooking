<%
' **************************************************************************
' File: system.asp
' Purpose: Functions for system wide activities.
' Created: 7/19/2011 - TAM
' Description:
'	Include this file on any page where system wide data is manipulated.
'	This file includes the following functions: GetMarqueeText,
'		GetCellPhoneCarriers, GetNotificationTypes, GetStates,
'		GetPaymentTerms
'
' Revision History:
' 7/19/201 - Created
' **************************************************************************

' **************************************************************************
' Function: GetMarqueeText
' Purpose: Retrieves the marquee text.
' Parameters:	<none>
' Return: The text to display on the marquee
' **************************************************************************
Function GetMarqueeText()
	Dim lsRet, lsSQL, loRS
	
	lsRet = ""
	
	lsSQL = "select marqueeMain, marqueeSub from tblMarquee where getdate() between startdate and enddate"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			If Not IsNull(loRS("marqueeMain")) Then
				lsRet = Trim(loRS("marqueeMain"))
				
				If Not IsNull(loRS("marqueeSub")) Then
					If Len(Trim(loRS("marqueeSub"))) > 0 Then
						lsRet = lsRet & " (Promo Code: " & Trim(loRS("marqueeSub")) & ")"
					End If
				End If
			End If
		End If
		
		DBCloseQuery loRS
	End If
	
	GetMarqueeText = lsRet
End Function

' **************************************************************************
' Function: GetCellPhoneCarriers
' Purpose: Retrieves a list of cell phone carriers.
' Parameters:	panCarrierIDs - Array of CellPhoneCarrierIDs found
'				pasCarrierNames - Array of carrier names found
' Return: True if sucessful, False if not
' **************************************************************************
Function GetCellPhoneCarriers(ByRef panCarrierIDs, ByRef pasCarrierNames)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select CellPhoneCarrierID, CellPhoneCarrier from tlkpCellPhoneCarriers"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve panCarrierIDs(lnPos), pasCarrierNames(lnPos)
				
				panCarrierIDs(lnPos) = loRS("CellPhoneCarrierID")
				If IsNull(loRS("CellPhoneCarrier")) Then
					pasCarrierNames(lnPos) = ""
				Else
					pasCarrierNames(lnPos) = loRS("CellPhoneCarrier")
				End If
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panCarrierIDs(0), pasCarrierNames(0)
			panCarrierIDs(0) = 0
			pasCarrierNames(0) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetCellPhoneCarriers = lbRet
End Function

' **************************************************************************
' Function: GetNotificationTypes
' Purpose: Retrieves a list of notification types.
' Parameters:	panNotificationTypeIDs - Array of NotificationTypeIDs found
'				pasNotificationTypes - Array of notification types found
' Return: True if sucessful, False if not
' **************************************************************************
Function GetNotificationTypes(ByRef panNotificationTypeIDs, ByRef pasNotificationTypes)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select NotificationTypeID, NotificationType from tlkpNotificationType"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve panNotificationTypeIDs(lnPos), pasNotificationTypes(lnPos)
				
				panNotificationTypeIDs(lnPos) = loRS("NotificationTypeID")
				If IsNull(loRS("NotificationType")) Then
					pasNotificationTypes(lnPos) = ""
				Else
					pasNotificationTypes(lnPos) = loRS("NotificationType")
				End If
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panNotificationTypeIDs(0), pasNotificationTypes(0)
			panNotificationTypeIDs(0) = 0
			pasNotificationTypes(0) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetNotificationTypes = lbRet
End Function

' **************************************************************************
' Function: GetStates
' Purpose: Retrieves a list of states.
' Parameters:	pasStateIDs - Array of StateIDs found
'				pasStates - Array of states found
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStates(ByRef pasStateIDs, ByRef pasStates)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select StateID, State from tlkpStates"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve pasStateIDs(lnPos), pasStates(lnPos)
				
				pasStateIDs(lnPos) = loRS("StateID")
				If IsNull(loRS("State")) Then
					pasStates(lnPos) = ""
				Else
					pasStates(lnPos) = loRS("State")
				End If
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim pasStateIDs(0), pasStates(0)
			pasStateIDs(0) = ""
			pasStates(0) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetStates = lbRet
End Function

' **************************************************************************
' Function: GetPaymentTerms
' Purpose: Retrieves a list of payment terms.
' Parameters:	panPaymentTermsIDs - Array of PaymentTermsIDs found
'				pasPaymentTerms - Array of payment terms found
' Return: True if sucessful, False if not
' **************************************************************************
Function GetPaymentTerms(ByRef panPaymentTermsIDs, ByRef pasPaymentTerms)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select PaymentTermsID, PaymentTermsDescription from tlkpPaymentTerms"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve panPaymentTermsIDs(lnPos), pasPaymentTerms(lnPos)
				
				panPaymentTermsIDs(lnPos) = loRS("PaymentTermsID")
				If IsNull(loRS("PaymentTermsDescription")) Then
					pasPaymentTerms(lnPos) = ""
				Else
					pasPaymentTerms(lnPos) = loRS("PaymentTermsDescription")
				End If
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panPaymentTermsIDs(0), pasPaymentTerms(0)
			panPaymentTermsIDs(0) = 0
			pasPaymentTerms(0) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetPaymentTerms = lbRet
End Function
%>