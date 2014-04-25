<%
' **************************************************************************
' File: litmos.asp
' Purpose: Functions for litmos training.
' Created: 10/11/2012 - TAM
' Description:
'	Include this file on any page where litmos training is needed.
'	This file includes the following functions: IsCoursesDue
'
' Revision History:
' 10/11/2012 - Created
' **************************************************************************

' **************************************************************************
' Function: IsCoursesDue
' Purpose: Determines if a course is due.
' Parameters:	pnEmpID - The EmpID to check
'				pnStoreID - The StoreID to check
'				pbResult - True if a course is due
'
' Return: True if successful, False if not
' **************************************************************************
Function IsCoursesDue(ByVal pnEmpID, ByVal pnStoreID, ByRef pbResult)
	Dim lbRet, loXMLHttp, lsURL, lsData, loXMLDoc
	
	lbRet = FALSE
	
	lsURL = gsLitmosCheckURL & "?EmpID=" & pnEmpID & "&StoreID=" & pnStoreID
	
	On Error Resume Next
	Set loXMLHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
	If Err.Number = 0 Then
		loXMLHttp.Open "GET", lsURL, false
		If Err.Number = 0 Then
			loXMLHttp.send
			If Err.Number = 0 Then
				lbRet = TRUE
				
				If Left(Trim(loXMLHttp.responseText), 11) = "COURSES DUE" Then
					pbResult = TRUE
				Else
					pbResult = FALSE
				End If
			Else
				gsDBErrorMessage = Err.Desciption
			End If
		Else
			gsDBErrorMessage = Err.Desciption
		End If
	Else
		gsDBErrorMessage = Err.Desciption
	End If
	
	IsCoursesDue = lbRet
End Function
%>