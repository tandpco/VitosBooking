<%
' **************************************************************************
' File: globals.asp
' Purpose: Various global settings.
' Created: 6/20/2011 - TAM
' Description:
'	Include this file at the top of every page (it should be the first file
'	included). This contains settings that affect virtually every page.
'
' Revision History:
' 6/20/2011 - Created
' **************************************************************************

' **************************************************************************
' Dev/Test mode - if dev mode, test mode will also be true
Dim gbTestMode, gbDevMode, gbWeb, gbSystemActive

gbSystemActive = TRUE
gbTestMode = TRUE
gbDevMode = TRUE
gbWeb = FALSE

' **************************************************************************
' The amount of time a page can be idle before automatically sigining off
Dim gnRedirectTime

If gbTestMode Then
	gnRedirectTime = 120
Else
	gnRedirectTime = 20
End If

' **************************************************************************
' The maximum quantity of items that can be added to a unit
Dim gnMaxItemsPerUnit

gnMaxItemsPerUnit = 24

' **************************************************************************
' The maximum quantity of toppers that can be added to a unit
Dim gnMaxToppersPerUnit

gnMaxToppersPerUnit = 10

' **************************************************************************
' The maximum quantity of free sides that can be added to a unit
Dim gnMaxFreeSidesPerUnit

gnMaxFreeSidesPerUnit = 20

' **************************************************************************
' The maximum quantity of additional sides that can be added to a unit
Dim gnMaxAddSidesPerUnit

gnMaxAddSidesPerUnit = 20

' **************************************************************************
' Hold order globals
Dim gnHoldOrderDisplayHours

gnHoldOrderDisplayHours = 24

' **************************************************************************
' Check acceptance criteria
Dim gnCheckAcceptMaxDaysSinceOrdered

gnCheckAcceptMaxDaysSinceOrdered = 90

' **************************************************************************
' Credit card globals
Dim gsPaymentGatewayURL, gsPaymentGatewayPW

gsPaymentGatewayURL = "https://transactions.test.secureexchange.net/SmartPayments/transact.asmx/ProcessCreditCard"
gsPaymentGatewayPW = "pizza4U08"

' **************************************************************************
' Mileage and tip rate
Dim gnMileageRate, gnMaximumMilesPerDrive, gnMinimumMilesPerDrive, gnMinimumTipPerDrive

gnMileageRate = .51
gnMaximumMilesPerDrive = 4
gnMinimumMilesPerDrive = 1
gnMinimumTipPerDrive = 1

' **************************************************************************
' Labor
Dim gnLaborFactor, gnMinimumWage, gnMaximumWage, gnMaximumShiftLength

gnLaborFactor = 1.25
gnMinimumWage = 4.00
gnMaximumWage = 20.00
gnMaximumShiftLength = 11

' **************************************************************************
Dim gsMailSystem, gsSMTPUserID, gsSMTPPassword, gsSMTPFrom

gsSMTPFrom = "Vito's Pizza and Subs <system@vitos.com>"

gsMailSystem = "mail.vitos.com"
gsSMTPUserID = "ordering@vitos.com"
gsSMTPPassword = "pizza4U08"

' **************************************************************************
' Litmos
Dim gsLitmosCheckURL

gsLitmosCheckURL = "https://dev.vitos.com/Litmos/LitmosCheck.aspx"

' **************************************************************************
' Inventory
Dim gdInventoryMinHour, gdInventoryRecountVariance

gdInventoryMinHour = 15
gdInventoryRecountVariance = 0.2

' **************************************************************************
' Close out past due threshold
Dim gnCloseOutThreshold

gnCloseOutThreshold = 4

' **************************************************************************
' Function for checking if an array is initialized
Function IsArrayInitialized(ByVal pa)    
	Dim lbRet
	
	lbRet = FALSE
	Err.Clear
	On Error Resume Next
	UBound(pa)
	If (Err.Number = 0) Then 
		lbRet = TRUE
	End If
	
	IsArrayInitialized = lbRet
End Function

' **************************************************************************
' Function for centering text within a field width
Function CenterText(ByVal psText, ByVal pnFieldWidth)
	Dim lsRet, lnLen
	
	lsRet = ""
	lnLen = Int((pnFieldWidth - Len(psText)) / 2)
	If lnLen > 0 Then
		lsRet = String(lnLen, " ") + psText
	End If
	
	CenterText = lsRet
End Function


Function IIf( expr, truepart, falsepart )
   IIf = falsepart
   If expr Then IIf = truepart
End Function
%>
