<%
' **************************************************************************
' File: math.asp
' Purpose: General mathematic functions.
' Created: 10/14/2011 - TAM
' Description:
'	Include this file at the top of any page where you will be performing
'	extended math functions. This file includes the following useful
'	functions: Round2
'
' Revision History:
' 10/14/2011 - Created
' **************************************************************************

' **************************************************************************
' Function: Round2
' Purpose: Rounds according to the .5 up rule.
' Parameters:	pnNumber - The number to be rounded
'				pnNumDigitsAfterDecimal - The number of decimal places to round to
' Return: The rounded result
' **************************************************************************
Function Round2(ByVal pnNumber, ByVal pnNumDigitsAfterDecimal)
	Round2 = CDbl(FormatNumber(pnNumber, pnNumDigitsAfterDecimal))
End Function
%>
