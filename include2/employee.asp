<%
' **************************************************************************
' File: employee.asp
' Purpose: Functions for employee related activities.
' Created: 7/19/2011 - TAM
' Description:
'	Include this file on any page where employee data is manipulated.
'	This file includes the following functions: GetAllStoreManagers,
'	GetEmployeeName, GetEmployeeShortName
'
' Revision History:
' 8/23/2011 - Created
' **************************************************************************

' **************************************************************************
' Function: GetAllStoreManagers
' Purpose: Retrieves all store managers.
' Parameters:	pnStoreID - The StoreID to search for
'				panEmpIDs - Array of EmpIDs
'				panEmployeeIDs - Array of EmployeeIDs
'				pasCardIDs - Array of CardIDs
' Return: True if sucessful, False if not
' **************************************************************************
Function GetAllStoreManagers(ByVal pnStoreID, ByRef panEmpIDs, ByRef panEmployeeIDs, ByRef pasCardIDs)
	Dim lbRet, lsSQL, loRS, lnPos
	
	lbRet = FALSE
	
	lsSQL = "select EmpID, EmployeeID, CardID from tblEmployee where (StoreID = 0 or StoreID = " & pnStoreID & ") and SystemRoleID > 1 and IsActive <> 0"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE
		
		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0
			
			Do While Not loRS.eof
				ReDim Preserve panEmpIDs(lnPos), panEmployeeIDs(lnPos), pasCardIDs(lnPos)
				
				panEmpIDs(lnPos) = loRS("EmpID")
				panEmployeeIDs(lnPos) = loRS("EmployeeID")
				pasCardIDs(lnPos) = Trim(loRS("CardID"))
				
				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panEmpIDs(0), panEmployeeIDs(0), pasCardIDs(0)
			panEmpIDs(0) = 0
			panEmployeeIDs(0) = 0
			pasCardIDs(0) = ""
		End If
		
		DBCloseQuery loRS
	End If
	
	GetAllStoreManagers = lbRet
End Function

' **************************************************************************
' Function: GetEmployeeName
' Purpose: Retrieves the name of an employee.
' Parameters:	pnEmpID - The StoreID to search for
' Return: The name of the employee
' **************************************************************************
Function GetEmployeeName(ByVal pnEmpID)
	Dim lsRet, lsSQL, loRS
	
	lsRet = ""
	
	lsSQL = "select FirstName, LastName from tblEmployee where EmpID = " & pnEmpID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lsRet = loRS("FirstName") & " " & loRS("LastName")
		End If
		
		DBCloseQuery loRS
	End If
	
	GetEmployeeName = lsRet
End Function

' **************************************************************************
' Function: GetEmployeeShortName
' Purpose: Retrieves the first name and last initial of an employee.
' Parameters:	pnEmpID - The EmpID to search for
' Return: The short name of the employee
' **************************************************************************
Function GetEmployeeShortName(ByVal pnEmpID)
	Dim lsRet, lsSQL, loRS
	
	lsRet = ""
	
	lsSQL = "select FirstName, LastName from tblEmployee where EmpID = " & pnEmpID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lsRet = loRS("FirstName") & " " & Left(Trim(loRS("LastName")), 1)
		End If
		
		DBCloseQuery loRS
	End If
	
	GetEmployeeShortName = lsRet
End Function
%>