<%
' **************************************************************************
' File: customer.asp
' Purpose: Functions for customer related activities.
' Created: 7/19/2011 - TAM
' Description:
'	Include this file on any page where customer data is manipulated.
'	This file includes the following functions: GetCustomersByPhone,
'		GetCustomerDetails, GetCustomerAddressDetails, GetCustomersByAddress
'		AddCustomerAddress, AssignCustomerPhone, AddCustomer, IsCustomerCheckOK,
'		GetLastCustomerOrderDate, GetCustomerAccounts, GetCustomerStoreAccounts,
'		GetStoreAccounts, GetAllAccounts, DebitAccountLedger, CreditAccountLedger, CustomerLogin,
'		LogWebActivity, UpdateWebActivityOrder, UpdateWebActivityStore, GetAccountName,
'		UpdateCustomerAddressNotes, UpdateCustomer, GetCustomerAddresses,
'		SetPrimaryAddress, DeleteCustomerAddress, GetAccountDetails, AddAccount,
'		AddStoreAccount, UpdateAccount, GetAccountContact, AddCustomerPhoneName,
'		IsAccountCollegeDebit, GetLastCustomerOrder
'
' Revision History:
' 7/19/201 - Created
' **************************************************************************

' **************************************************************************
' Function: GetCustomersByPhone
' Purpose: Finds customers based on phone number.
' Parameters:	pnPhone - The phone number to search for
'				panCustomerIDs - Array of CustomerIDs found
'				panPrimaryAddressIDs - Array of PrimaryAddressIDs found
'				panAddressIDs - Array of AddressIDs found
'				panStoreIDs - Array of StoreIDs found
'				pasAddresses - Array of addresses found
'				pasNames - Array of names found
' Return: True if sucessful, False if not
' **************************************************************************
Function GetCustomersByPhone(ByVal pnPhone, ByRef panCustomerIDs, ByRef panPrimaryAddressIDs, ByRef panAddressIDs, ByRef panStoreIDs, ByRef pasAddresses, ByRef pasNames, ByRef rowCount)
	Dim lbRet, lsSQL, loRS, lnPos

	lbRet = FALSE

'	lsSQL = "SELECT tblCustomers.CustomerID, PrimaryAddressID, tblAddresses.AddressID, StoreID, AddressLine1, AddressLine2, FirstName, LastName from tblCustomers left outer join trelCustomerAddresses on tblCustomers.CustomerID = trelCustomerAddresses.CustomerID left outer join tblAddresses on trelCustomerAddresses.AddressID = tblAddresses.AddressID where HomePhone = '" & pnPhone & "' or CellPhone = '" & pnPhone & "' or WorkPhone = '" & pnPhone & "' or FAXPhone = '" & pnPhone & "' order by tblCustomers.CustomerID, tblAddresses.AddressID"
    lsSQL = "SELECT TOP " + rowCount + " tblCustomers.CustomerID,PrimaryAddressID,tblAddresses.AddressID,tOrders.StoreID,AddressLine1,AddressLine2,FirstName,LastName FROM tblCustomers OUTER APPLY (SELECT TOP 1 * FROM tblOrders WHERE (tblOrders.CustomerID = tblCustomers.CustomerID) ORDER BY tblOrders.TransactionDate) tOrders LEFT OUTER JOIN trelCustomerAddresses ON tOrders.CustomerID = trelCustomerAddresses.CustomerID LEFT OUTER JOIN tblAddresses on trelCustomerAddresses.AddressID = tblAddresses.AddressID WHERE HomePhone = '" & pnPhone & "' or CellPhone = '" & pnPhone & "' or WorkPhone = '" & pnPhone & "' or FAXPhone = '" & pnPhone & "' ORDER BY tOrders.TransactionDate DESC"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE

		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0

			Do While Not loRS.eof
				ReDim Preserve panCustomerIDs(lnPos), panPrimaryAddressIDs(lnPos), panAddressIDs(lnPos), panStoreIDs(lnPos), pasAddresses(lnPos), pasNames(lnPos)

				panCustomerIDs(lnPos) = loRS("CustomerID")
				panPrimaryAddressIDs(lnPos) = loRS("PrimaryAddressID")
				If IsNull(loRS("AddressID")) Then
					panAddressIDs(lnPos) = 0
				Else
					panAddressIDs(lnPos) = loRS("AddressID")
					panStoreIDs(lnPos) = loRS("StoreID")
					If IsNull(loRS("AddressLine2")) Then
						pasAddresses(lnPos) = Trim(loRS("AddressLine1"))
					Else
						If Len(loRS("AddressLine2")) = 0 Then
							pasAddresses(lnPos) = Trim(loRS("AddressLine1"))
						Else
							pasAddresses(lnPos) = Trim(loRS("AddressLine1")) & " #" & Trim(loRS("AddressLine2"))
						End If
					End If
				End If
				If IsNull(loRS("FirstName")) Then
					pasNames(lnPos) = ""
				Else
					If IsNull(loRS("LastName")) Then
						pasNames(lnPos) = Trim(loRS("FirstName"))
					Else
						If Len(loRS("LastName")) = 0 Then
							pasNames(lnPos) = Trim(loRS("FirstName"))
						Else
							pasNames(lnPos) = Trim(loRS("FirstName")) & " " & Trim(loRS("LastName"))
						End If
					End If
				End If

				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panCustomerIDs(0), panPrimaryAddressIDs(0), panAddressIDs(0), panStoreIDs(0), pasAddresses(0), pasNames(0)
			panCustomerIDs(0) = 0
			panPrimaryAddressIDs(0) = 0
			panAddressIDs(0) = 0
			panStoreIDs(0) = 0
			pasAddresses(0) = ""
			pasNames(0) = ""
		End If

		DBCloseQuery loRS
	End If

	GetCustomersByPhone = lbRet
End Function

Function GetAddressesByPhone(ByVal pnPhone, ByVal rowCount, ByRef panAddressIDs, ByRef panStoreIDs, ByRef pasAddresses)
	Dim lbRet, lsSQL, loRS, lnPos

	lbRet = FALSE
  lsSQL = "SELECT TOP " + rowCount + " tblAddresses.AddressID,MAX(tblAddresses.StoreID) as StoreID,MAX(AddressLine1) as AddressLine1,MAX(AddressLine2) as AddressLine2		  FROM tblCustomers		  OUTER APPLY (SELECT TOP 1 * FROM tblOrders WHERE (tblOrders.CustomerID = tblCustomers.CustomerID) ORDER BY tblOrders.TransactionDate) tOrders 		  LEFT OUTER JOIN trelCustomerAddresses ON tOrders.CustomerID = trelCustomerAddresses.CustomerID 		  LEFT OUTER JOIN tblAddresses on trelCustomerAddresses.AddressID = tblAddresses.AddressID 		  WHERE HomePhone = '" & pnPhone & "' or CellPhone = '" & pnPhone & "' or WorkPhone = '" & pnPhone & "' or FAXPhone = '" & pnPhone & "' 		  GROUP BY tblAddresses.AddressID		  ORDER BY MAX(tOrders.TransactionDate) DESC"

	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE

		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0

			Do While Not loRS.eof
				ReDim Preserve panAddressIDs(lnPos), panStoreIDs(lnPos), pasAddresses(lnPos)

				If IsNull(loRS("AddressID")) Then
					panAddressIDs(lnPos) = 0
				Else
					panAddressIDs(lnPos) = loRS("AddressID")
					panStoreIDs(lnPos) = loRS("StoreID")
					If IsNull(loRS("AddressLine2")) Then
						pasAddresses(lnPos) = Trim(loRS("AddressLine1"))
					Else
						If Len(loRS("AddressLine2")) = 0 Then
							pasAddresses(lnPos) = Trim(loRS("AddressLine1"))
						Else
							pasAddresses(lnPos) = Trim(loRS("AddressLine1")) & " #" & Trim(loRS("AddressLine2"))
						End If
					End If
				End If

				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panAddressIDs(0), panStoreIDs(0), pasAddresses(0)
			panAddressIDs(0) = 0
			panStoreIDs(0) = 0
			pasAddresses(0) = ""
		End If

		DBCloseQuery loRS
	End If

	GetAddressesByPhone = lbRet
End Function




' **************************************************************************
' Function: GetCustomerDetails
' Purpose: Retrieves customer details.
' Parameters:	pnCustomerID - The CustomerID to search for
'				psEMail - The e-mail address
'				psFirstName - The first name
'				psLastName - The last name
'				pdtBirthdate - The birth date
'				pnPrimaryAddressID - The primary AddressID
'				psHomePhone - The home phone number
'				psCellPhone - The cell phone number
'				psWorkPhone - The work phone number
'				psFAXPhone - The FAX phone number
'				pbIsEMailList - Flag for e-mail list
'				pbIsTextList - Flag for SMS list
' Return: True if sucessful, False if not
' **************************************************************************
Function GetCustomerDetails(ByVal pnCustomerID, ByRef psEMail, ByRef psFirstName, ByRef psLastName, ByRef pdtBirthdate, ByRef pnPrimaryAddressID, ByRef psHomePhone, ByRef psCellPhone, ByRef psWorkPhone, ByRef psFAXPhone, ByRef pbIsEMailList, ByRef pbIsTextList)
	Dim lbRet, lsSQL, loRS

	lbRet = FALSE

	lsSQL = "select EMail, FirstName, LastName, Birthdate, PrimaryAddressID, HomePhone, CellPhone, WorkPhone, FAXPhone, IsEMailList, IsTextList from tblCustomers where CustomerID = " & pnCustomerID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lbRet = TRUE

			If IsNull(loRS("EMail")) Then
				psEMail = ""
			Else
				psEMail = Trim(loRS("EMail"))
			End If
			If IsNull(loRS("FirstName")) Then
				psFirstName = ""
			Else
				psFirstName = Trim(loRS("FirstName"))
			End If
			If IsNull(loRS("LastName")) Then
				psLastName = ""
			Else
				psLastName = Trim(loRS("LastName"))
			End If
			If IsNull(loRS("Birthdate")) Then
				pdtBirthdate = DateValue("1/1/1900")
			Else
				pdtBirthdate = loRS("Birthdate")
			End If
			pnPrimaryAddressID = loRS("PrimaryAddressID")
			If IsNull(loRS("HomePhone")) Then
				psHomePhone = ""
			Else
				psHomePhone = Trim(loRS("HomePhone"))
			End If
			If IsNull(loRS("CellPhone")) Then
				psCellPhone = ""
			Else
				psCellPhone = Trim(loRS("CellPhone"))
			End If
			If IsNull(loRS("WorkPhone")) Then
				psWorkPhone = ""
			Else
				psWorkPhone = Trim(loRS("WorkPhone"))
			End If
			If IsNull(loRS("FAXPhone")) Then
				psFAXPhone = ""
			Else
				psFAXPhone = Trim(loRS("FAXPhone"))
			End If
			If loRS("IsEMailList") <> 0 Then
				pbIsEMailList = TRUE
			Else
				pbIsEMailList = FALSE
			End If
			If loRS("IsTextList") <> 0 Then
				pbIsTextList = TRUE
			Else
				pbIsTextList = FALSE
			End If
		End If

		DBCloseQuery loRS
	End If

	GetCustomerDetails = lbRet
End Function
Function GetCustomerDetails2(ByVal pnCustomerID, ByRef psEMail, ByRef psFirstName, ByRef psLastName, ByRef pdtBirthdate, ByRef pnPrimaryAddressID, ByRef psHomePhone, ByRef psCellPhone, ByRef psWorkPhone, ByRef psFAXPhone, ByRef pbIsEMailList, ByRef pbIsTextList,ByRef psExtension)
	Dim lbRet, lsSQL, loRS

	lbRet = FALSE

	lsSQL = "select EMail, FirstName, LastName, Birthdate, PrimaryAddressID, HomePhone, CellPhone, WorkPhone, FAXPhone,extension, IsEMailList, IsTextList from tblCustomers where CustomerID = " & pnCustomerID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lbRet = TRUE

			If IsNull(loRS("EMail")) Then
				psEMail = ""
			Else
				psEMail = Trim(loRS("EMail"))
			End If
			If IsNull(loRS("FirstName")) Then
				psFirstName = ""
			Else
				psFirstName = Trim(loRS("FirstName"))
			End If
			If IsNull(loRS("LastName")) Then
				psLastName = ""
			Else
				psLastName = Trim(loRS("LastName"))
			End If
			If IsNull(loRS("extension")) Then
				psExtension = ""
			Else
				psExtension = Trim(loRS("extension"))
			End If
			If IsNull(loRS("Birthdate")) Then
				pdtBirthdate = DateValue("1/1/1900")
			Else
				pdtBirthdate = loRS("Birthdate")
			End If
			pnPrimaryAddressID = loRS("PrimaryAddressID")
			If IsNull(loRS("HomePhone")) Then
				psHomePhone = ""
			Else
				psHomePhone = Trim(loRS("HomePhone"))
			End If
			If IsNull(loRS("CellPhone")) Then
				psCellPhone = ""
			Else
				psCellPhone = Trim(loRS("CellPhone"))
			End If
			If IsNull(loRS("WorkPhone")) Then
				psWorkPhone = ""
			Else
				psWorkPhone = Trim(loRS("WorkPhone"))
			End If
			If IsNull(loRS("FAXPhone")) Then
				psFAXPhone = ""
			Else
				psFAXPhone = Trim(loRS("FAXPhone"))
			End If
			If loRS("IsEMailList") <> 0 Then
				pbIsEMailList = TRUE
			Else
				pbIsEMailList = FALSE
			End If
			If loRS("IsTextList") <> 0 Then
				pbIsTextList = TRUE
			Else
				pbIsTextList = FALSE
			End If
		End If

		DBCloseQuery loRS
	End If

	GetCustomerDetails2 = lbRet
End Function

' **************************************************************************
' Function: GetCustomerAddressDetails
' Purpose: Retrieves customer address details.
' Parameters:	pnCustomerID - The CustomerID to search for
'				pnAddressID - The AddressID to search for
'				psCustomerAddressDescription - The address description
'				psCustomerAddressNotes - The customer notes
' Return: True if sucessful, False if not
' **************************************************************************
Function GetCustomerAddressDetails(ByVal pnCustomerID, ByVal pnAddressID, ByRef psCustomerAddressDescription, ByRef psCustomerAddressNotes)
	Dim lbRet, lsSQL, loRS

	lbRet = FALSE

	lsSQL = "select CustomerAddressDescription, CustomerAddressNotes from trelCustomerAddresses where CustomerID = " & pnCustomerID & " and AddressID = " & pnAddressID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lbRet = TRUE

			psCustomerAddressDescription = Trim(loRS("CustomerAddressDescription"))
			If IsNull(loRS("CustomerAddressNotes")) Then
				psCustomerAddressNotes = ""
			Else
				psCustomerAddressNotes = Trim(loRS("CustomerAddressNotes"))
			End If
		End If

		DBCloseQuery loRS
	End If

	GetCustomerAddressDetails = lbRet
End Function

    ' **************************************************************************
' Function: GetCustomerPrimaryAddressDetails
' Purpose: Retrieves customer addresses for selected PrimaryAddressID.
' Parameters:	pnPrimaryAddressID
'				psEMail - The e-mail address
'				psFirstName - The first name
'				psLastName - The last name
' Return: True if sucessful, False if not
' **************************************************************************
Function GetCustomerPrimaryAddressDetails(ByVal pnPrimaryAddressID, ByRef pasNames, ByRef panCustomerIDs, ByRef pasEMails, ByRef extensions, ByRef pasPhones)
	Dim lbRet, lsSQL, loRS, lnPos

	lbRet = FALSE

    'lsSQL = "SELECT DISTINCT CustomerID,* FROM tblCustomers WHERE PrimaryAddressID = "&pnPrimaryAddressID


	lsSQL = "select tblCustomers.CustomerID, HomePhone,CellPhone,WorkPhone,FirstName,extension, LastName,Email from tblCustomers left join trelCustomerAddresses on tblCustomers.CustomerID = trelCustomerAddresses.CustomerID where AddressID = " & pnPrimaryAddressID & " order by tblCustomers.CustomerID"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE

		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0

			Do While Not loRS.eof
				ReDim Preserve panCustomerIDs(lnPos),pasNames(lnPos),pasEMails(lnPos),panAddressIDs(lnPos),extensions(lnPos),pasPhones(lnPos)

				panCustomerIDs(lnPos) = loRS("CustomerID")

				If IsNull(loRS("FirstName")) Then
					pasNames(lnPos) = "Unknown"
				Else
					pasNames(lnPos) = Trim(loRS("FirstName") & " " & loRS("LastName"))
				End If

				pasEMails(lnPos) = Trim(loRS("EMail"))
				pasPhones(lnPos) = Trim(loRS("HomePhone")&" "&loRS("CellPhone")&" "&loRS("WorkPhone"))
				extensions(lnPos) = Trim(loRS("extension"))

				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panCustomerIDs(0),pasNames(0),pasEMails(0),extensions(0),pasPhones(0)
			pasNames(0) = 0
			pasEMails(0) = ""
			pasPhones(0) = ""
			extensions(0) = ""
      panCustomerIDs(0) = 0
		End If

		DBCloseQuery loRS
	End If

	GetCustomerPrimaryAddressDetails = lbRet
End Function

' **************************************************************************
' Function: GetCustomersByAddress
' Purpose: Finds customers based on address.
' Parameters:	pnAddressID - The address to search for
'				panCustomerIDs - Array of CustomerIDs found
'				pasNames - Array of names found
' Return: True if sucessful, False if not
' **************************************************************************
Function GetCustomersByAddress(ByVal pnAddressID, ByRef panCustomerIDs, ByRef pasNames)
	Dim lbRet, lsSQL, loRS, lnPos

	lbRet = FALSE

	lsSQL = "select tblCustomers.CustomerID, FirstName, LastName from trelCustomerAddresses inner join tblCustomers on tblCustomers.CustomerID = trelCustomerAddresses.CustomerID where AddressID = " & pnAddressID & " order by tblCustomers.CustomerID"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE

		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0

			Do While Not loRS.eof
				ReDim Preserve panCustomerIDs(lnPos), pasNames(lnPos)

				panCustomerIDs(lnPos) = loRS("CustomerID")
				If IsNull(loRS("FirstName")) Then
					pasNames(lnPos) = ""
				Else
					If IsNull(loRS("LastName")) Then
						pasNames(lnPos) = Trim(loRS("FirstName"))
					Else
						If Len(loRS("LastName")) = 0 Then
							pasNames(lnPos) = Trim(loRS("FirstName"))
						Else
							pasNames(lnPos) = Trim(loRS("FirstName")) & " " & Trim(loRS("LastName"))
						End If
					End If
				End If

				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panCustomerIDs(0), pasNames(0)
			panCustomerIDs(0) = 0
			pasNames(0) = ""
		End If

		DBCloseQuery loRS
	End If

	GetCustomersByAddress = lbRet
End Function

' **************************************************************************
' Function: AddCustomerAddress
' Purpose: Adds an address to a customer.
' Parameters:	pnCustomerID - The CustomerID
'				pnAddressID - The AddressID
'				psCustomerAddressDescription - The description of the address
' Return: True if sucessful, False if not
' **************************************************************************
Function AddCustomerAddress(ByVal pnCustomerID, ByVal pnAddressID, psCustomerAddressDescription)
	Dim lbRet, lsSQL

	lbRet = FALSE

	lsSQL = "insert into trelCustomerAddresses (CustomerID, AddressID, CustomerAddressDescription) values (" & pnCustomerID & ", " & pnAddressID & ", '" & DBCleanLiteral(psCustomerAddressDescription) & "')"
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If

	AddCustomerAddress = lbRet
End Function

' **************************************************************************
' Function: AssignCustomerPhone
' Purpose: Assigns a phone number to a customer.
' Parameters:	pnCustomerID - The CustomerID
'				psPhoneNumber - The phone number
'				pnPhoneType - The type of phone number
' Return: True if sucessful, False if not
' **************************************************************************
Function AssignCustomerPhone(ByVal pnCustomerID, ByVal psPhoneNumber, pnPhoneType)
	Dim lbRet, lsSQL

	lbRet = FALSE

	lsSQL = "update tblCustomers set "
	Select Case pnPhoneType
		Case 1
			lsSQL = lsSQL & "CellPhone = '"
		Case 2
			lsSQL = lsSQL & "WorkPhone = '"
		Case 3
			lsSQL = lsSQL & "FAXPhone = '"
		Case Else
			lsSQL = lsSQL & "HomePhone = '"
	End Select
	lsSQL = lsSQL & DBCleanLiteral(psPhoneNumber) & "' where CustomerID = " & pnCustomerID

	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If

	AssignCustomerPhone = lbRet
End Function

' **************************************************************************
' Function: AddCustomer
' Purpose: Adds a new customer.
' Parameters:	psEMail - The e-mail address
'				psPassword - The password
'				psFirstName - The first name
'				psLastName - The last name
'				pdtBirthdate - The birth date
'				pnPrimaryAddressID - The primary AddressID
'				psHomePhone - The home phone number
'				psCellPhone - The cell phone number
'				psWorkPhone - The work phone number
'				psFAXPhone - The FAX phone number
'				pbIsEMailList - Flag for e-mail list
'				pbIsTextList - Flag for SMS list
' Return: The new CustomerID
' **************************************************************************
Function AddCustomer(ByVal psEMail, ByVal psPassword, ByVal psFirstName, ByVal psLastName, ByVal pdtBirthdate, ByVal pnPrimaryAddressID, ByVal psHomePhone, ByVal psCellPhone, ByVal psWorkPhone, ByVal psFAXPhone, ByVal pbIsEMailList, ByVal pbIsTextList)
	Dim lnRet, lsSQL, loRS

	lnRet = 0

	lsSQL = "EXEC AddCustomer @pEMail = "
	If Len(psEMail) = 0 Then
		lsSQL = lsSQL & "NULL"
	Else
		lsSQL = lsSQL & "'" & DBCleanLiteral(psEMail) & "'"
	End If
	lsSQL = lsSQL & ", @pPassword = "
	If Len(psPassword) = 0 Then
		lsSQL = lsSQL & "NULL"
	Else
		lsSQL = lsSQL & "'" & DBCleanLiteral(psPassword) & "'"
	End If
	lsSQL = lsSQL & ", @pFirstName = "
	If Len(psFirstName) = 0 Then
		lsSQL = lsSQL & "NULL"
	Else
		lsSQL = lsSQL & "'" & DBCleanLiteral(psFirstName) & "'"
	End If
	lsSQL = lsSQL & ", @pLastName = "
	If Len(psLastName) = 0 Then
		lsSQL = lsSQL & "NULL"
	Else
		lsSQL = lsSQL & "'" & DBCleanLiteral(psLastName) & "'"
	End If
	lsSQL = lsSQL & ", @pBirthdate = "
	If pdtBirthdate = DateValue("1/1/1900") Then
		lsSQL = lsSQL & "NULL"
	Else
		lsSQL = lsSQL & "'" & DBCleanLiteral(pdtBirthdate) & "'"
	End If
	lsSQL = lsSQL & ", @pPrimaryAddressID = " & pnPrimaryAddressID
	lsSQL = lsSQL & ", @pHomePhone = "
	If Len(psHomePhone) = 0 Then
		lsSQL = lsSQL & "NULL"
	Else
		lsSQL = lsSQL & "'" & DBCleanLiteral(psHomePhone) & "'"
	End If
	lsSQL = lsSQL & ", @pCellPhone = "
	If Len(psCellPhone) = 0 Then
		lsSQL = lsSQL & "NULL"
	Else
		lsSQL = lsSQL & "'" & DBCleanLiteral(psCellPhone) & "'"
	End If
	lsSQL = lsSQL & ", @pWorkPhone = "
	If Len(psWorkPhone) = 0 Then
		lsSQL = lsSQL & "NULL"
	Else
		lsSQL = lsSQL & "'" & DBCleanLiteral(psWorkPhone) & "'"
	End If
	lsSQL = lsSQL & ", @pFAXPhone = "
	If Len(psFAXPhone) = 0 Then
		lsSQL = lsSQL & "NULL"
	Else
		lsSQL = lsSQL & "'" & DBCleanLiteral(psFAXPhone) & "'"
	End If
	If pbIsEMailList Then
		lsSQL = lsSQL & ", @pIsEmailList = 1"
	Else
		lsSQL = lsSQL & ", @pIsEmailList = 0"
	End If
	If pbIsTextList Then
		lsSQL = lsSQL & ", @pIsTextList = 1"
	Else
		lsSQL = lsSQL & ", @pIsTextList = 0"
	End If
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lnRet = CLng(loRS(0))
		End If

		DBCloseQuery loRS
	End If

	AddCustomer = lnRet
End Function

' **************************************************************************
' Function: UpdateCustomer
' Purpose: Updates an existing customer.
' Parameters:	psEMail - The e-mail address
'				psFirstName - The first name
'				psLastName - The last name
'                psExtension - The Customers Phone Extension
' Return: True/False
' **************************************************************************
Function UpdateCustomer_New(ByVal psEMail, ByVal psFirstName, ByVal psLastName, ByVal psExtension, ByVal custID)
	Dim lnRet, lsSQL, loRS

	lnRet = FALSE

	lsSQL = "UPDATE tblCustomers SET EMail = '" & psEMail & "', firstName = '" & psFirstName & "',  lastName = '" & psLastname & "', extension = '" & psExtension & "' WHERE CustomerID = " & custID
	If DBExecuteSQL(lsSQL) Then
		lnRet = TRUE
	End If

	UpdateCustomer_New = lnRet
End Function

Function UpdateCustomer_New2(ByVal pnCustomerID, ByVal psEMail, ByVal psFirstName, ByVal psLastName, ByVal pdtBirthdate, ByVal psHomePhone, ByVal psCellPhone, ByVal psWorkPhone, ByVal psFAXPhone, ByVal pbIsEMailList, ByVal pbIsTextList, ByVal pbNoChecks, ByVal psExtension)
	Dim lbRet, lsSQL, loRS

	lbRet = FALSE

	lsSQL = "update tblCustomers set "
	If Len(Trim(psEMail)) = 0 Then
		lsSQL = lsSQL & "EMail = NULL"
	Else
		lsSQL = lsSQL & "EMail = '" & DBCleanLiteral(psEMail) & "'"
	End If
	If Len(Trim(psFirstName)) = 0 Then
		lsSQL = lsSQL & ", FirstName = NULL"
	Else
		lsSQL = lsSQL & ", FirstName = '" & DBCleanLiteral(psFirstName) & "'"
	End If
	If Len(Trim(psLastName)) = 0 Then
		lsSQL = lsSQL & ", LastName = NULL"
	Else
		lsSQL = lsSQL & ", LastName = '" & DBCleanLiteral(psLastName) & "'"
	End If
	If Len(Trim(pdtBirthdate)) = 0 Then
		lsSQL = lsSQL & ", BirthDate = NULL"
	Else
		lsSQL = lsSQL & ", BirthDate = '" & DBCleanLiteral(pdtBirthdate) & "'"
	End If
	If Len(Trim(psHomePhone)) = 0 Then
		lsSQL = lsSQL & ", HomePhone = NULL"
	Else
		lsSQL = lsSQL & ", HomePhone = '" & DBCleanLiteral(psHomePhone) & "'"
	End If
	If Len(Trim(psCellPhone)) = 0 Then
		lsSQL = lsSQL & ", CellPhone = NULL"
	Else
		lsSQL = lsSQL & ", CellPhone = '" & DBCleanLiteral(psCellPhone) & "'"
	End If
	If Len(Trim(psWorkPhone)) = 0 Then
		lsSQL = lsSQL & ", WorkPhone = NULL"
	Else
		lsSQL = lsSQL & ", WorkPhone = '" & DBCleanLiteral(psWorkPhone) & "'"
	End If
	If Len(Trim(psExtension)) = 0 Then
		lsSQL = lsSQL & ", extension = NULL"
	Else
		lsSQL = lsSQL & ", extension = '" & DBCleanLiteral(psExtension) & "'"
	End If
	If Len(Trim(psFAXPhone)) = 0 Then
		lsSQL = lsSQL & ", FAXPhone = NULL"
	Else
		lsSQL = lsSQL & ", FAXPhone = '" & DBCleanLiteral(psFAXPhone) & "'"
	End If
	If pbIsEMailList Then
		lsSQL = lsSQL & ", IsEMailList = 1"
	Else
		lsSQL = lsSQL & ", IsEMailList = 0"
	End If
	If pbIsTextList Then
		lsSQL = lsSQL & ", IsTextList = 1"
	Else
		lsSQL = lsSQL & ", IsTextList = 0"
	End If
	If pbNoChecks Then
		lsSQL = lsSQL & ", NoChecks= 1"
	Else
		lsSQL = lsSQL & ", NoChecks= 0"
	End If
	lsSQL = lsSQL & " where CustomerID = " & pnCustomerID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If

	UpdateCustomer_New2 = lbRet
End Function
' **************************************************************************
' Function: IsCustomerCheckOK
' Purpose: Determines if a check can be accepted from a customer.
' Parameters:	pnCustomerID - The CustomerID to find
' Return: True if checks are OK, False if not
' **************************************************************************
Function IsCustomerCheckOK(ByVal pnCustomerID)
	Dim lbRet, lsSQL, loRS, lnPos

	lbRet = FALSE

	lsSQL = "select NoChecks from tblCustomers where CustomerID = " & pnCustomerID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			If loRS("NoChecks") = 0 Then
				lbRet = TRUE
			End If
		End If

		DBCloseQuery loRS
	End If

	IsCustomerCheckOK = lbRet
End Function

' **************************************************************************
' Function: GetLastCustomerOrderDate
' Purpose: Returns the date of the last customer order.
' Parameters:	pnCustomerID - The CustomerID to find
' Return: The date of the last order or 1/1/1900 if never ordered
' **************************************************************************
Function GetLastCustomerOrderDate(ByVal pnCustomerID)
	Dim ldtRet, lsSQL, loRS, lnPos

	ldtRet = DateValue("1/1/1900")

	lsSQL = "select MAX(ReleaseDate) as LastOrderDate from tblOrders where CustomerID = " & pnCustomerID & " and IsPaid <> 0 and OrderStatusID = 10"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			If Not IsNull(loRS("LastOrderDate")) Then
				ldtRet = loRS("LastOrderDate")
			End If
		End If

		DBCloseQuery loRS
	End If

	GetLastCustomerOrderDate = ldtRet
End Function

' **************************************************************************
' Function: GetCustomerAccounts
' Purpose: Retrieves a list of accounts that a customer can charge to.
' Parameters:	pnCustomerID - The customer ID to search for
'				panAccountIDs - Array of AccountIDs found
'				pasAccountNames - Array of account names found
'				pabAccountOnHolds - Array of on hold flags
' Return: True if sucessful, False if not
' **************************************************************************
Function GetCustomerAccounts(ByVal pnCustomerID, ByRef panAccountIDs, ByRef pasAccountNames, ByRef pabAccountOnHolds)
	Dim lbRet, lsSQL, loRS, lnPos

	lbRet = FALSE

	lsSQL = "select trelAccountsCustomers.AccountID, AccountName, OnHold from trelAccountsCustomers inner join tblAccounts on trelAccountsCustomers.AccountID = tblAccounts.AccountID where trelAccountsCustomers.CustomerID = " & pnCustomerID & " and tblAccounts.IsActive <> 0"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE

		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0

			Do While Not loRS.eof
				ReDim Preserve panAccountIDs(lnPos), pasAccountNames(lnPos), pabAccountOnHolds(lnPos)

				panAccountIDs(lnPos) = loRS("AccountID")
				If IsNull(loRS("AccountName")) Then
					pasAccountNames(lnPos) = ""
				Else
					pasAccountNames(lnPos) = loRS("AccountName")
				End If
				If loRS("OnHold") <> 0 Then
					pabAccountOnHolds(lnPos) = TRUE
				Else
					pabAccountOnHolds(lnPos) = FALSE
				End If

				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panAccountIDs(0), pasAccountNames(0), pabAccountOnHolds(0)
			panAccountIDs(0) = 0
			pasAccountNames(0) = ""
			pabAccountOnHolds(0) = FALSE
		End If

		DBCloseQuery loRS
	End If

	GetCustomerAccounts = lbRet
End Function

' **************************************************************************
' Function: GetCustomerStoreAccounts
' Purpose: Retrieves a list of accounts that a customer can charge to for a store.
' Parameters:	pnCustomerID - The customer ID to search for
'				pnStoreID - The store ID to search for
'				panAccountIDs - Array of AccountIDs found
'				pasAccountNames - Array of account names found
'				pabAccountOnHolds - Array of on hold flags
' Return: True if sucessful, False if not
' **************************************************************************
Function GetCustomerStoreAccounts(ByVal pnCustomerID, ByVal pnStoreID, ByRef panAccountIDs, ByRef pasAccountNames, ByRef pabAccountOnHolds)
	Dim lbRet, lsSQL, loRS, lnPos

	lbRet = FALSE

	lsSQL = "select trelAccountsCustomers.AccountID, AccountName, OnHold from trelAccountsCustomers inner join tblAccounts on trelAccountsCustomers.AccountID = tblAccounts.AccountID inner join trelAccountsStores on tblAccounts.AccountID = trelAccountsStores.AccountID where trelAccountsCustomers.CustomerID = " & pnCustomerID & " and trelAccountsStores.StoreID = " & pnStoreID & " and tblAccounts.IsActive <> 0 and tblAccounts.CollegeDebitStoreID is null"
	lsSQL = lsSQL & " union all select AccountID, AccountName, OnHold from tblAccounts where tblAccounts.IsActive <> 0 and tblAccounts.CollegeDebitStoreID = " & pnStoreID
If Session("SecurityID") > 1 Then
	lsSQL = lsSQL & " union all select AccountID, AccountName, OnHold from tblAccounts where AccountID = 1 and tblAccounts.IsActive <> 0"
	lsSQL = lsSQL & " order by 1"
End If
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE

		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0

			Do While Not loRS.eof
				ReDim Preserve panAccountIDs(lnPos), pasAccountNames(lnPos), pabAccountOnHolds(lnPos)

				panAccountIDs(lnPos) = loRS("AccountID")
				If IsNull(loRS("AccountName")) Then
					pasAccountNames(lnPos) = ""
				Else
					pasAccountNames(lnPos) = loRS("AccountName")
				End If
				If loRS("OnHold") <> 0 Then
					pabAccountOnHolds(lnPos) = TRUE
				Else
					pabAccountOnHolds(lnPos) = FALSE
				End If

				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panAccountIDs(0), pasAccountNames(0), pabAccountOnHolds(0)
			panAccountIDs(0) = 0
			pasAccountNames(0) = ""
			pabAccountOnHolds(0) = FALSE
		End If

		DBCloseQuery loRS
	End If

	GetCustomerStoreAccounts = lbRet
End Function

' **************************************************************************
' Function: GetStoreAccounts
' Purpose: Retrieves a list of all accounts for a store.
' Parameters:	pnStoreID - The store ID to search for (for College Debit account)
'				panAccountIDs - Array of AccountIDs found
'				pasAccountNames - Array of account names found
'				pabAccountOnHolds - Array of on hold flags
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStoreAccounts(ByVal pnStoreID, ByRef panAccountIDs, ByRef pasAccountNames, ByRef pabAccountOnHolds)
	Dim lbRet, lsSQL, loRS, lnPos

	lbRet = FALSE

	lsSQL = "select StoreID, tblAccounts.AccountID, AccountName, OnHold from tblAccounts inner join trelAccountsStores on tblAccounts.AccountID = trelAccountsStores.AccountID and trelAccountsStores.StoreID = " & pnStoreID & " where IsActive <> 0 and (CollegeDebitStoreID is null or CollegeDebitStoreID = " & pnStoreID & ") order by CollegeDebitStoreID desc, AccountName"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE

		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0

			Do While Not loRS.eof
				ReDim Preserve panAccountIDs(lnPos), pasAccountNames(lnPos), pabAccountOnHolds(lnPos)

				panAccountIDs(lnPos) = loRS("AccountID")
				If IsNull(loRS("AccountName")) Then
					pasAccountNames(lnPos) = ""
				Else
					pasAccountNames(lnPos) = loRS("AccountName")
				End If
				If loRS("OnHold") <> 0 Then
					pabAccountOnHolds(lnPos) = TRUE
				Else
					pabAccountOnHolds(lnPos) = FALSE
				End If

				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panAccountIDs(0), pasAccountNames(0), pabAccountOnHolds(0)
			panAccountIDs(0) = 0
			pasAccountNames(0) = ""
			pabAccountOnHolds(0) = FALSE
		End If

		DBCloseQuery loRS
	End If

	GetStoreAccounts = lbRet
End Function

' **************************************************************************
' Function: GetStoreCollegeDebitAccounts
' Purpose: Retrieves a list of all college debit accounts for a store.
' Parameters:	pnStoreID - The store ID to search for (for College Debit account)
'				panAccountIDs - Array of AccountIDs found
'				pasAccountNames - Array of account names found
'				pabAccountOnHolds - Array of on hold flags
' Return: True if sucessful, False if not
' **************************************************************************
Function GetStoreCollegeDebitAccounts(ByVal pnStoreID, ByRef panAccountIDs, ByRef pasAccountNames, ByRef pabAccountOnHolds)
	Dim lbRet, lsSQL, loRS, lnPos

	lbRet = FALSE

	lsSQL = "select StoreID, tblAccounts.AccountID, AccountName, OnHold from tblAccounts inner join trelAccountsStores on tblAccounts.AccountID = trelAccountsStores.AccountID and trelAccountsStores.StoreID = " & pnStoreID & " where IsActive <> 0 and CollegeDebitStoreID = " & pnStoreID & " order by AccountName"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE

		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0

			Do While Not loRS.eof
				ReDim Preserve panAccountIDs(lnPos), pasAccountNames(lnPos), pabAccountOnHolds(lnPos)

				panAccountIDs(lnPos) = loRS("updateAccountID")
				If IsNull(loRS("AccountName")) Then
					pasAccountNames(lnPos) = ""
				Else
					pasAccountNames(lnPos) = loRS("AccountName")
				End If
				If loRS("OnHold") <> 0 Then
					pabAccountOnHolds(lnPos) = TRUE
				Else
					pabAccountOnHolds(lnPos) = FALSE
				End If

				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panAccountIDs(0), pasAccountNames(0), pabAccountOnHolds(0)
			panAccountIDs(0) = 0
			pasAccountNames(0) = ""
			pabAccountOnHolds(0) = FALSE
		End If

		DBCloseQuery loRS
	End If

	GetStoreCollegeDebitAccounts = lbRet
End Function

' **************************************************************************
' Function: GetAllAccounts
' Purpose: Retrieves a list of all accounts.
' Parameters:	pnStoreID - The store ID to search for (for College Debit account)
'				panAllAccountIDs - Array of AccountIDs found
'				pasAllAccountNames - Array of account names found
'				pabAllAccountOnHolds - Array of on hold flags
' Return: True if sucessful, False if not
' **************************************************************************
Function GetAllAccounts(ByVal pnStoreID, ByRef panAllAccountIDs, ByRef pasAllAccountNames, ByRef pabAllAccountOnHolds)
	Dim lbRet, lsSQL, loRS, lnPos

	lbRet = FALSE

	lsSQL = "select AccountID, AccountName, OnHold from tblAccounts where IsActive <> 0 and (CollegeDebitStoreID is null or CollegeDebitStoreID = " & pnStoreID & ") order by AccountName"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE

		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0

			Do While Not loRS.eof
				ReDim Preserve panAllAccountIDs(lnPos), pasAllAccountNames(lnPos), pabAllAccountOnHolds(lnPos)

				panAllAccountIDs(lnPos) = loRS("AccountID")
				If IsNull(loRS("AccountName")) Then
					pasAllAccountNames(lnPos) = ""
				Else
					pasAllAccountNames(lnPos) = loRS("AccountName")
				End If
				If loRS("OnHold") <> 0 Then
					pabAllAccountOnHolds(lnPos) = TRUE
				Else
					pabAllAccountOnHolds(lnPos) = FALSE
				End If

				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panAllAccountIDs(0), pasAllAccountNames(0), pabAllAccountOnHolds(0)
			panAllAccountIDs(0) = 0
			pasAllAccountNames(0) = ""
			pabAllAccountOnHolds(0) = FALSE
		End If

		DBCloseQuery loRS
	End If

	GetAllAccounts = lbRet
End Function

' **************************************************************************
' Function: DebitAccountLedger
' Purpose: Adds a debit to an account.
' Parameters:	pnStoreID - The StoreID
'				pdtTransactionDate - The transaction date
'				pnAccountID - The AccountID
'				pnOrderID - The OrderID (0 if not for an order)
'				psLedgerDescription - The description of the ledger entry
'				pdDebit - The amount to debit
' Return: True if sucessful, False if not
' **************************************************************************
Function DebitAccountLedger(ByVal pnStoreID, ByVal pdtTransactionDate, ByVal pnAccountID, ByVal pnOrderID, ByVal psLedgerDescription, ByVal pdDebit)
	Dim lbRet, lsSQL, loRS

	lbRet = FALSE

	' Ensure not already entered
	If pnOrderID = 0 Then
		lsSQL = "select LedgerID from tblLedger where LedgerDate = '" & pdtTransactionDate & "' and StoreID = " & pnStoreID & " and AccountID = " & pnAccountID & " and LedgerDescription = '" & DBCleanLiteral(psLedgerDescription) & "' and Debit = " & pdDebit & ""
	Else
		lsSQL = "select LedgerID from tblLedger where LedgerDate = '" & pdtTransactionDate & "' and StoreID = " & pnStoreID & " and OrderID = " & pnOrderID & " and AccountID = " & pnAccountID & " and LedgerDescription = '" & DBCleanLiteral(psLedgerDescription) & "' and Debit = " & pdDebit & ""
	End If
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lbRet = TRUE
		Else
			If pnOrderID = 0 Then
				lsSQL = "insert into tblLedger (LedgerDate, StoreID, AccountID, LedgerDescription, Debit) values ('" & pdtTransactionDate & "', " & pnStoreID & ", " & pnAccountID & ", '" & DBCleanLiteral(psLedgerDescription) & "', " & pdDebit & ")"
			Else
				lsSQL = "insert into tblLedger (LedgerDate, StoreID, AccountID, OrderID, LedgerDescription, Debit) values ('" & pdtTransactionDate & "', " & pnStoreID & ", " & pnAccountID & ", " & pnOrderID & ", '" & DBCleanLiteral(psLedgerDescription) & "', " & pdDebit & ")"
			End If
			If DBExecuteSQL(lsSQL) Then
				lbRet = TRUE

				lsSQL = "insert into trelAccountsStores (AccountID, StoreID) values (" & pnAccountID & ", " & pnStoreID & ")"
				DBExecuteSQL lsSQL
			End If
		End If

		DBCloseQuery loRS
	End If

	DebitAccountLedger = lbRet
End Function

' **************************************************************************
' Function: CreditAccountLedger
' Purpose: Adds a credit to an account.
' Parameters:	pnAccountID - The AccountID
'				pnOrderID - The OrderID (0 if not for an order)
'				psLedgerDescription - The description of the ledger entry
'				pdCredit - The amount to credit
' Return: True if sucessful, False if not
' **************************************************************************
Function CreditAccountLedger(ByVal pnAccountID, ByVal pnOrderID, ByVal psLedgerDescription, ByVal pdCredit)
	Dim lbRet, lsSQL

	lbRet = FALSE

	If pnOrderID = 0 Then
		lsSQL = "insert into tblLedger (AccountID, LedgerDescription, Credit) values (" & pnAccountID & ", '" & DBCleanLiteral(psLedgerDescription) & "', " & pdCredit & ")"
	Else
		lsSQL = "insert into tblLedger (AccountID, OrderID, LedgerDescription, Credit) values (" & pnAccountID & ", " & pnOrderID & ", '" & DBCleanLiteral(psLedgerDescription) & "', " & pdCredit & ")"
	End If
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If

	CreditAccountLedger = lbRet
End Function

' **************************************************************************
' Function: CustomerLogin
' Purpose: Validates a customer login and retrieves customer details.
' Parameters:	psEMail - The customer's e-mail address
'				psPassword - The customer's password
'				pnCustomerID - The CustomerID
'				psFirstName - The first name
'				psLastName - The last name
'				pnPrimaryAddressID - The primary AddressID
'				psHomePhone - The home phone number
'				psCellPhone - The cell phone number
'				psWorkPhone - The work phone number
'				psFAXPhone - The FAX phone number
'				pdtBirthdate - The birth date
' Return: True if sucessful, False if not
' **************************************************************************
Function CustomerLogin(ByVal psEMail, ByVal psPassword, ByRef pnCustomerID, ByRef psFirstName, ByRef psLastName, ByRef pnPrimaryAddressID, ByRef psHomePhone, ByRef psCellPhone, ByRef psWorkPhone, ByRef psFAXPhone, ByRef pdtBirthdate)
	Dim lbRet, lsSQL, loRS

	lbRet = FALSE

	lsSQL = "select CustomerID, FirstName, LastName, PrimaryAddressID, HomePhone, CellPhone, WorkPhone, FAXPhone, Birthdate from tblCustomers where EMail = '" & DBCleanLiteral(psEMail) & "' and Password = '" & DBCleanLiteral(MD5(psPassword)) & "'"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lbRet = TRUE

			pnCustomerID = Trim(loRS("CustomerID"))
			If IsNull(loRS("FirstName")) Then
				psFirstName = ""
			Else
				psFirstName = Trim(loRS("FirstName"))
			End If
			If IsNull(loRS("LastName")) Then
				psLastName = ""
			Else
				psLastName = Trim(loRS("LastName"))
			End If
			pnPrimaryAddressID = loRS("PrimaryAddressID")
			If IsNull(loRS("HomePhone")) Then
				psHomePhone = ""
			Else
				psHomePhone = Trim(loRS("HomePhone"))
			End If
			If IsNull(loRS("CellPhone")) Then
				psCellPhone = ""
			Else
				psCellPhone = Trim(loRS("CellPhone"))
			End If
			If IsNull(loRS("WorkPhone")) Then
				psWorkPhone = ""
			Else
				psWorkPhone = Trim(loRS("WorkPhone"))
			End If
			If IsNull(loRS("FAXPhone")) Then
				psFAXPhone = ""
			Else
				psFAXPhone = Trim(loRS("FAXPhone"))
			End If
			If IsNull(loRS("Birthdate")) Then
				pdtBirthdate = DateValue("1/1/1900")
			Else
				pdtBirthdate = loRS("Birthdate")
			End If
		End If

		DBCloseQuery loRS
	End If

	CustomerLogin = lbRet
End Function

' **************************************************************************
' Function: LogWebActivity
' Purpose: Logs website activity.
' Parameters:	psEMail - The e-mail address
'				psAddress1 - The address line 1
'				psAddress2 - The address line 2
'				psCity - The city
'				psState - The state
'				psPostalCode - The postal code
'				pnStoreID - The StoreID
' Return: The new ActivityID
' **************************************************************************
Function LogWebActivity(ByVal psEmail, ByVal psAddress1, ByVal psAddress2, ByVal psCity, ByVal psState, ByVal pnPostalCode, ByVal pnStoreID)
	Dim lnRet, lsSQL, loRS

	lnRet = 0

	lsSQL = "EXEC LogWebActivity @pSessionID = " & Session.SessionID & ", @pEMail = '" & psEmail & "', @pAddress1 = '" & psAddress1 & "', @pAddress2 = '" & psAddress2 & "', @pCity = '" & psCity & "', @pState = '" & psState & "', @pPostalCode = '" & pnPostalCode & "', @pStoreID = " & pnStoreID & ", @pIPAddress = '" & Request.ServerVariables("REMOTE_ADDR") & "', @pRefID = '" & Session("RefID") & "'"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lnRet = CLng(loRS(0))
		End If

		DBCloseQuery loRS
	End If

	LogWebActivity = lnRet
End Function

' **************************************************************************
' Function: UpdateWebActivityOrder
' Purpose: Updates the web activity log with order ID.
' Parameters:	pnActivityID - The WebActivityID
'				pnOrderID - The OrderID
' Return: True if sucessful, False if not
' **************************************************************************
Function UpdateWebActivityOrder(ByVal pnActivityID, ByVal pnOrderID)
	Dim lbRet, lsSQL

	lbRet = FALSE

	lsSQL = "update tblWebActivity set OrderID = " & pnOrderID & " where WebActivityID = " & pnActivityID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If

	UpdateWebActivityOrder = lbRet
End Function

' **************************************************************************
' Function: UpdateWebActivityStore
' Purpose: Updates the web activity log with store ID and order type ID.
' Parameters:	pnActivityID - The WebActivityID
'				pnStoreID - The StoreID
'				pnOrderType - The OrderTypeID
' Return: True if sucessful, False if not
' **************************************************************************
Function UpdateWebActivityStore(ByVal pnActivityID, ByVal pnStoreID, ByVal pnOrderType)
	Dim lbRet, lsSQL

	lbRet = FALSE

	lsSQL = "update tblWebActivity set StoreID = " & pnStoreID & ", OrderTypeID = " & pnOrderType & " where WebActivityID = " & pnActivityID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If

	UpdateWebActivityStore = lbRet
End Function

' **************************************************************************
' Function: GetAccountName
' Purpose: Retrieves the name associated with an account.
' Parameters:	pnAccountID - The account ID to search for
' Return: The account name
' **************************************************************************
Function GetAccountName(ByVal pnAccountID)
	Dim lsRet, lsSQL, loRS

	lsRet = ""

	lsSQL = "select AccountName from tblAccounts where AccountID = " & pnAccountID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lsRet = loRS("AccountName")
		End If

		DBCloseQuery loRS
	End If

	GetAccountName = lsRet
End Function

' **************************************************************************
' Function: UpdateCustomerAddressNotes
' Purpose: Updates customer address notes.
' Parameters:	pnCustomerAddressID - The CustomerID to update
'				pnAddressID - The AddressID to update
'				psCustomerAddressNotes - The customer address notes
' Return: True if sucessful, False if not
' **************************************************************************
Function UpdateCustomerAddressNotes(ByVal pnCustomerID, ByVal pnAddressID, ByVal psCustomerAddressNotes)
	Dim lbRet, lsSQL, loRS

	lbRet = FALSE

	lsSQL = "update trelCustomerAddresses set CustomerAddressNotes = '" & DBCleanLiteral(psCustomerAddressNotes) & "' where CustomerID = " & pnCustomerID & " and AddressID = " & pnAddressID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If

	UpdateCustomerAddressNotes = lbRet
End Function

' **************************************************************************
' Function: UpdateCustomer
' Purpose: Updates customer information
' Parameters:	pnCustomerAddressID - The CustomerID to update
'				psEMail - The e-mail address
'				psFirstName - The first name
'				psLastName - The last name
'				pdtBirthdate - The birth date
'				psHomePhone - The home phone number
'				psCellPhone - The cell phone number
'				psWorkPhone - The work phone number
'				psFAXPhone - The FAX phone number
'				pbIsEMailList - Flag for e-mail list
'				pbIsTextList - Flag for SMS list
'				pbNoChecks - Flag for no checks
' Return: True if sucessful, False if not
' **************************************************************************
Function UpdateCustomer(ByVal pnCustomerID, ByVal psEMail, ByVal psFirstName, ByVal psLastName, ByVal pdtBirthdate, ByVal psHomePhone, ByVal psCellPhone, ByVal psWorkPhone, ByVal psFAXPhone, ByVal pbIsEMailList, ByVal pbIsTextList, ByVal pbNoChecks)
	Dim lbRet, lsSQL, loRS

	lbRet = FALSE

	lsSQL = "update tblCustomers set "
	If Len(Trim(psEMail)) = 0 Then
		lsSQL = lsSQL & "EMail = NULL"
	Else
		lsSQL = lsSQL & "EMail = '" & DBCleanLiteral(psEMail) & "'"
	End If
	If Len(Trim(psFirstName)) = 0 Then
		lsSQL = lsSQL & ", FirstName = NULL"
	Else
		lsSQL = lsSQL & ", FirstName = '" & DBCleanLiteral(psFirstName) & "'"
	End If
	If Len(Trim(psLastName)) = 0 Then
		lsSQL = lsSQL & ", LastName = NULL"
	Else
		lsSQL = lsSQL & ", LastName = '" & DBCleanLiteral(psLastName) & "'"
	End If
	If Len(Trim(pdtBirthdate)) = 0 Then
		lsSQL = lsSQL & ", BirthDate = NULL"
	Else
		lsSQL = lsSQL & ", BirthDate = '" & DBCleanLiteral(pdtBirthdate) & "'"
	End If
	If Len(Trim(psHomePhone)) = 0 Then
		lsSQL = lsSQL & ", HomePhone = NULL"
	Else
		lsSQL = lsSQL & ", HomePhone = '" & DBCleanLiteral(psHomePhone) & "'"
	End If
	If Len(Trim(psCellPhone)) = 0 Then
		lsSQL = lsSQL & ", CellPhone = NULL"
	Else
		lsSQL = lsSQL & ", CellPhone = '" & DBCleanLiteral(psCellPhone) & "'"
	End If
	If Len(Trim(psWorkPhone)) = 0 Then
		lsSQL = lsSQL & ", WorkPhone = NULL"
	Else
		lsSQL = lsSQL & ", WorkPhone = '" & DBCleanLiteral(psWorkPhone) & "'"
	End If
	If Len(Trim(psFAXPhone)) = 0 Then
		lsSQL = lsSQL & ", FAXPhone = NULL"
	Else
		lsSQL = lsSQL & ", FAXPhone = '" & DBCleanLiteral(psFAXPhone) & "'"
	End If
	If pbIsEMailList Then
		lsSQL = lsSQL & ", IsEMailList = 1"
	Else
		lsSQL = lsSQL & ", IsEMailList = 0"
	End If
	If pbIsTextList Then
		lsSQL = lsSQL & ", IsTextList = 1"
	Else
		lsSQL = lsSQL & ", IsTextList = 0"
	End If
	If pbNoChecks Then
		lsSQL = lsSQL & ", NoChecks= 1"
	Else
		lsSQL = lsSQL & ", NoChecks= 0"
	End If
	lsSQL = lsSQL & " where CustomerID = " & pnCustomerID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If

	UpdateCustomer = lbRet
End Function

' **************************************************************************
' Function: GetCustomerAddresses
' Purpose: Finds addresses associated with a customer
' Parameters:	pnCustomerID - The CustomerID to search for
'				panAddressIDs - Array of AddressIDs found
'				pasAddresses - Array of addresses found
' Return: True if sucessful, False if not
' **************************************************************************
Function GetCustomerAddresses(ByVal pnCustomerID, ByRef panAddressIDs, ByRef pasAddresses)
	Dim lbRet, lsSQL, loRS, lnPos

	lbRet = FALSE

	lsSQL = "select tblAddresses.AddressID, AddressLine1, AddressLine2, City, State, PostalCode from trelCustomerAddresses inner join tblAddresses on tblAddresses.AddressID = trelCustomerAddresses.AddressID where CustomerID = " & pnCustomerID & " order by tblAddresses.AddressID"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE

		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0

			Do While Not loRS.eof
				ReDim Preserve panAddressIDs(lnPos), pasAddresses(lnPos)

				panAddressIDs(lnPos) = loRS("AddressID")
				If IsNull(loRS("AddressLine2")) Then
					pasAddresses(lnPos) = Trim(loRS("AddressLine1"))
				Else
					If Len(loRS("AddressLine2")) = 0 Then
						pasAddresses(lnPos) = Trim(loRS("AddressLine1"))
					Else
						pasAddresses(lnPos) = Trim(loRS("AddressLine1")) & " #" & Trim(loRS("AddressLine2"))
					End If
				End If
				pasAddresses(lnPos) = pasAddresses(lnPos) & ", " & Trim(loRS("City")) & ", " & Trim(loRS("State")) & " " & Trim(loRS("PostalCode"))

				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			ReDim panAddressIDs(0), pasAddresses(0)
			panAddressIDs(0) = 0
			pasAddresses(0) = ""
		End If

		DBCloseQuery loRS
	End If

	GetCustomerAddresses = lbRet
End Function
' **************************************************************************
' Function: GetCustomerAddresses
' Purpose: Finds addresses associated with a customer
' Parameters:	pnCustomerID - The CustomerID to search for
'				panAddressIDs - Array of AddressIDs found
'				pasAddresses - Array of addresses found
' Return: True if sucessful, False if not
' **************************************************************************
Function GetAddressDetails2(ByVal addressID, ByRef addressStreet, ByRef addressZip)
	Dim lbRet, lsSQL, loRS, lnPos

	lbRet = FALSE
	lsSQL = "select TOP 1 * from tblAddresses where AddressID = " & addressID

	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		lbRet = TRUE

		If Not loRS.bof And Not loRS.eof Then
			lnPos = 0

			Do While Not loRS.eof

				addressZip= Trim(loRS("PostalCode"))
				If IsNull(loRS("AddressLine2")) Then
					addressStreet = Trim(loRS("AddressLine1"))
				Else
					If Len(loRS("AddressLine2")) = 0 Then
						addressStreet = Trim(loRS("AddressLine1"))
					Else
						addressStreet = Trim(loRS("AddressLine1")) & " #" & Trim(loRS("AddressLine2"))
					End If
				End If

				lnPos = lnPos + 1
				loRS.MoveNext
			Loop
		Else
			addressStreet = ""
			addressZip = ""
		End If

		DBCloseQuery loRS
	End If

	GetAddressDetails2 = lbRet
End Function

' **************************************************************************
' Function: SetPrimaryAddress
' Purpose: Sets a customer's primary address.
' Parameters:	pnCustomerID - The CustomerID to update
'				pnAddressID - The AddressID to set as primary
' Return: True if sucessful, False if not
' **************************************************************************
Function SetPrimaryAddress(ByVal pnCustomerID, ByVal pnAddressID)
	Dim lbRet, lsSQL, loRS

	lbRet = FALSE

	lsSQL = "update tblCustomers set PrimaryAddressID = " & pnAddressID & " where CustomerID = " & pnCustomerID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If

	SetPrimaryAddress = lbRet
End Function

' **************************************************************************
' Function: DeleteCustomerAddress
' Purpose: Deletes an address from a customer
' Parameters:	pnCustomerID - The CustomerID to search for
'				pnAddressID - The AddressID to delete
' Return: True if sucessful, False if not
' **************************************************************************
Function DeleteCustomerAddress(ByVal pnCustomerID, ByVal pnAddressID)
	Dim lbRet, lsSQL, loRS

	lbRet = FALSE

	lsSQL = "delete from trelCustomerAddresses where CustomerID = " & pnCustomerID & " and AddressID = " & pnAddressID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If

	DeleteCustomerAddress = lbRet
End Function

' **************************************************************************
' Function: GetAccountDetails
' Purpose: Retrieves account details.
' Parameters:	pnAccountID - The CustomerID to search for
'				pbIsActive - Is active flag
'				pbOnHold - On hold flag
'				psAccountName - The account name
'				psPrimaryContactName - The primary contact name
'				psPrimaryContactTelephone - The primary contact telephone
'				pnCellPhoneCarrierID - The cell phone carrier ID
'				psPrimaryContactEmail - The primary contact email
'				pnNotificationTypeID - The notification type ID
'				psAccountAddressLine1 - The address line 1
'				psAccountAddressLine2 - The address line 2
'				psAccountCity - The city
'				psAccountStateID - The state ID
'				psAccountZip - The postal code
'				psAccountTelephone - The telephone
'				psAccountFax - The FAX number
'				pnPaymentTermsID - The payment terms ID
'				pdCreditLimit - The credit limit
'				psAccountNotes - The notes
' Return: True if sucessful, False if not
' **************************************************************************
Function GetAccountDetails(ByVal pnAccountID, ByRef pbIsActive, ByRef pbOnHold, ByRef psAccountName, ByRef psPrimaryContactName, ByRef psPrimaryContactTelephone, ByRef pnCellPhoneCarrierID, ByRef psPrimaryContactEmail, ByRef pnNotificationTypeID, ByRef psAccountAddressLine1, ByRef psAccountAddressLine2, ByRef psAccountCity, ByRef psAccountStateID, ByRef psAccountZip, ByRef psAccountTelephone, ByRef psAccountFax, ByRef pnPaymentTermsID, ByRef pdCreditLimit, ByRef psAccountNotes)
	Dim lbRet, lsSQL, loRS

	lbRet = FALSE

	lsSQL = "select IsActive, OnHold, AccountName, PrimaryContactName, PrimaryContactTelephone, CellPhoneCarrierID, PrimaryContactEmail, NotificationTypeID, AccountAddressLine1, AccountAddressLine2, AccountCity, AccountStateID, AccountZip, AccountTelephone, AccountFax, PaymentTermsID, CreditLimit, AccountNotes from tblAccounts where AccountID = " & pnAccountID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lbRet = TRUE

			If loRS("IsActive") <> 0 Then
				pbIsActive = TRUE
			Else
				pbIsActive = FALSE
			End If
			If loRS("OnHold") <> 0 Then
				pbOnHold = TRUE
			Else
				pbOnHold = FALSE
			End If
			psAccountName = loRS("AccountName")
			psPrimaryContactName = loRS("PrimaryContactName")
			psPrimaryContactTelephone = loRS("PrimaryContactTelephone")
			If IsNull(loRS("CellPhoneCarrierID")) Then
				pnCellPhoneCarrierID = 0
			Else
				pnCellPhoneCarrierID = loRS("CellPhoneCarrierID")
			End If
			psPrimaryContactEmail = loRS("PrimaryContactEmail")
			If IsNull(loRS("NotificationTypeID")) Then
				pnNotificationTypeID = 0
			Else
				pnNotificationTypeID = loRS("NotificationTypeID")
			End If
			psAccountAddressLine1 = loRS("AccountAddressLine1")
			If IsNull(loRS("NotificationTypeID")) Then
				psAccountAddressLine2 = ""
			Else
				psAccountAddressLine2 = loRS("AccountAddressLine2")
			End If
			psAccountCity = loRS("AccountCity")
			psAccountStateID = loRS("AccountStateID")
			psAccountZip = loRS("AccountZip")
			If IsNull(loRS("AccountTelephone")) Then
				psAccountTelephone = ""
			Else
				psAccountTelephone = loRS("AccountTelephone")
			End If
			If IsNull(loRS("AccountFax")) Then
				psAccountFax = ""
			Else
				psAccountFax = loRS("AccountFax")
			End If
			pnPaymentTermsID = loRS("PaymentTermsID")
			pdCreditLimit = loRS("CreditLimit")
			If IsNull(loRS("AccountNotes")) Then
				psAccountNotes = ""
			Else
				psAccountNotes = loRS("AccountNotes")
			End If
		End If

		DBCloseQuery loRS
	End If

	GetAccountDetails = lbRet
End Function

' **************************************************************************
' Function: AddAccount
' Purpose: Adds a new account.
' Parameters:	pbIsActive - Is active flag
'				pbOnHold - On hold flag
'				psAccountName - The account name
'				psPrimaryContactName - The primary contact name
'				psPrimaryContactTelephone - The primary contact telephone
'				pnCellPhoneCarrierID - The cell phone carrier ID
'				psPrimaryContactEmail - The primary contact email
'				pnNotificationTypeID - The notification type ID
'				psAccountAddressLine1 - The address line 1
'				psAccountAddressLine2 - The address line 2
'				psAccountCity - The city
'				psAccountStateID - The state ID
'				psAccountZip - The postal code
'				psAccountTelephone - The telephone
'				psAccountFax - The FAX number
'				pnPaymentTermsID - The payment terms ID
'				pdCreditLimit - The credit limit
'				psAccountNotes - The notes
' Return: The new CustomerID
' **************************************************************************
Function AddAccount(ByVal pbIsActive, ByVal pbOnHold, ByVal psAccountName, ByVal psPrimaryContactName, ByVal psPrimaryContactTelephone, ByVal pnCellPhoneCarrierID, ByVal psPrimaryContactEmail, ByVal pnNotificationTypeID, ByVal psAccountAddressLine1, ByVal psAccountAddressLine2, ByVal psAccountCity, ByVal psAccountStateID, ByVal psAccountZip, ByVal psAccountTelephone, ByVal psAccountFax, ByVal pnPaymentTermsID, ByVal pdCreditLimit, ByVal psAccountNotes)
	Dim lnRet, lsSQL, loRS

	lnRet = 0

	lsSQL = "EXEC AddAccount "
	If pbIsActive Then
		lsSQL = lsSQL & "@pbIsActive = 1"
	Else
		lsSQL = lsSQL & "@pbIsActive = 0"
	End If
	If pbOnHold Then
		lsSQL = lsSQL & ", @pbOnHold = 1"
	Else
		lsSQL = lsSQL & ", @pbOnHold = 0"
	End If
	lsSQL = lsSQL & ", @psAccountName = '" & DBCleanLiteral(psAccountName) & "'"
	lsSQL = lsSQL & ", @psPrimaryContactName = '" & DBCleanLiteral(psPrimaryContactName) & "'"
	lsSQL = lsSQL & ", @psPrimaryContactTelephone = '" & DBCleanLiteral(psPrimaryContactTelephone) & "'"
	If pnCellPhoneCarrierID = 0 Then
		lsSQL = lsSQL & ", @pnCellPhoneCarrierID = NULL"
	Else
		lsSQL = lsSQL & ", @pnCellPhoneCarrierID = " & DBCleanLiteral(pnCellPhoneCarrierID)
	End If
	lsSQL = lsSQL & ", @psPrimaryContactEmail  = '" & DBCleanLiteral(psPrimaryContactEmail) & "'"
	If pnNotificationTypeID = 0 Then
		lsSQL = lsSQL & ", @pnNotificationTypeID = NULL"
	Else
		lsSQL = lsSQL & ", @pnNotificationTypeID = " & DBCleanLiteral(pnNotificationTypeID)
	End If
	lsSQL = lsSQL & ", @psAccountAddressLine1 = '" & DBCleanLiteral(psAccountAddressLine1) & "'"
	If Len(psAccountAddressLine1) = 0 Then
		lsSQL = lsSQL & ", @psAccountAddressLine2 = NULL"
	Else
		lsSQL = lsSQL & ", @psAccountAddressLine2 = '" & DBCleanLiteral(psAccountAddressLine2) & "'"
	End If
	lsSQL = lsSQL & ", @psAccountCity = '" & DBCleanLiteral(psAccountCity) & "'"
	lsSQL = lsSQL & ", @psAccountStateID = '" & DBCleanLiteral(psAccountStateID) & "'"
	lsSQL = lsSQL & ", @psAccountZip = '" & DBCleanLiteral(psAccountZip) & "'"
	If Len(psAccountTelephone) = 0 Then
		lsSQL = lsSQL & ", @psAccountTelephone = NULL"
	Else
		lsSQL = lsSQL & ", @psAccountTelephone = '" & DBCleanLiteral(psAccountTelephone) & "'"
	End If
	If Len(psAccountFax) = 0 Then
		lsSQL = lsSQL & ", @psAccountFax = NULL"
	Else
		lsSQL = lsSQL & ", @psAccountFax = '" & DBCleanLiteral(psAccountFax) & "'"
	End If
	lsSQL = lsSQL & ", @pnPaymentTermsID = " & DBCleanLiteral(pnPaymentTermsID)
	lsSQL = lsSQL & ", @pdCreditLimit = " & DBCleanLiteral(pdCreditLimit)
	If Len(psAccountNotes) = 0 Then
		lsSQL = lsSQL & ", @psAccountNotes = NULL"
	Else
		lsSQL = lsSQL & ", @psAccountNotes = '" & DBCleanLiteral(psAccountNotes) & "'"
	End If
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lnRet = CLng(loRS(0))
		End If

		DBCloseQuery loRS
	End If

	AddAccount = lnRet
End Function

' **************************************************************************
' Function: AddStoreAccount
' Purpose: Associates a store to an account.
' Parameters:	pnStoreID - The Store ID
'				pnAccountID - The Account ID
' Return: True if sucessful, False if not
' **************************************************************************
Function AddStoreAccount(ByVal pnStoreID, ByVal pnAccountID)
	Dim lbRet, lsSQL

	lbRet = FALSE

	lsSQL = "insert into trelAccountsStores (AccountID, StoreID) values (" & pnAccountID & ", " & pnStoreID & ")"
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If

	AddStoreAccount = lbRet
End Function

' **************************************************************************
' Function: UpdateAccount
' Purpose: Updates an account.
' Parameters:	pnAccountID - The Account ID to update
'				pbIsActive - Is active flag
'				pbOnHold - On hold flag
'				psAccountName - The account name
'				psPrimaryContactName - The primary contact name
'				psPrimaryContactTelephone - The primary contact telephone
'				pnCellPhoneCarrierID - The cell phone carrier ID
'				psPrimaryContactEmail - The primary contact email
'				pnNotificationTypeID - The notification type ID
'				psAccountAddressLine1 - The address line 1
'				psAccountAddressLine2 - The address line 2
'				psAccountCity - The city
'				psAccountStateID - The state ID
'				psAccountZip - The postal code
'				psAccountTelephone - The telephone
'				psAccountFax - The FAX number
'				pnPaymentTermsID - The payment terms ID
'				pdCreditLimit - The credit limit
'				psAccountNotes - The notes
' Return: True if sucessful, False if not
' **************************************************************************
Function UpdateAccount(ByVal pnAccountID, ByVal pbIsActive, ByVal pbOnHold, ByVal psAccountName, ByVal psPrimaryContactName, ByVal psPrimaryContactTelephone, ByVal pnCellPhoneCarrierID, ByVal psPrimaryContactEmail, ByVal pnNotificationTypeID, ByVal psAccountAddressLine1, ByVal psAccountAddressLine2, ByVal psAccountCity, ByVal psAccountStateID, ByVal psAccountZip, ByVal psAccountTelephone, ByVal psAccountFax, ByVal pnPaymentTermsID, ByVal pdCreditLimit, ByVal psAccountNotes)
	Dim lbRet, lsSQL

	lbRet = FALSE

	lsSQL = "update tblAccounts set "
	If pbIsActive Then
		lsSQL = lsSQL & "IsActive = 1"
	Else
		lsSQL = lsSQL & "IsActive = 0"
	End If
	If pbOnHold Then
		lsSQL = lsSQL & ", OnHold = 1"
	Else
		lsSQL = lsSQL & ", OnHold = 0"
	End If
	lsSQL = lsSQL & ", AccountName = '" & DBCleanLiteral(psAccountName) & "'"
	lsSQL = lsSQL & ", PrimaryContactName = '" & DBCleanLiteral(psPrimaryContactName) & "'"
	lsSQL = lsSQL & ", PrimaryContactTelephone = '" & DBCleanLiteral(psPrimaryContactTelephone) & "'"
	If pnCellPhoneCarrierID = 0 Then
		lsSQL = lsSQL & ", CellPhoneCarrierID = NULL"
	Else
		lsSQL = lsSQL & ", CellPhoneCarrierID = " & DBCleanLiteral(pnCellPhoneCarrierID)
	End If
	lsSQL = lsSQL & ", PrimaryContactEmail  = '" & DBCleanLiteral(psPrimaryContactEmail) & "'"
	If pnNotificationTypeID = 0 Then
		lsSQL = lsSQL & ", NotificationTypeID = NULL"
	Else
		lsSQL = lsSQL & ", NotificationTypeID = " & DBCleanLiteral(pnNotificationTypeID)
	End If
	lsSQL = lsSQL & ", AccountAddressLine1 = '" & DBCleanLiteral(psAccountAddressLine1) & "'"
	If Len(psAccountAddressLine1) = 0 Then
		lsSQL = lsSQL & ", AccountAddressLine2 = NULL"
	Else
		lsSQL = lsSQL & ", AccountAddressLine2 = '" & DBCleanLiteral(psAccountAddressLine2) & "'"
	End If
	lsSQL = lsSQL & ", AccountCity = '" & DBCleanLiteral(psAccountCity) & "'"
	lsSQL = lsSQL & ", AccountStateID = '" & DBCleanLiteral(psAccountStateID) & "'"
	lsSQL = lsSQL & ", AccountZip = '" & DBCleanLiteral(psAccountZip) & "'"
	If Len(psAccountTelephone) = 0 Then
		lsSQL = lsSQL & ", AccountTelephone = NULL"
	Else
		lsSQL = lsSQL & ", AccountTelephone = '" & DBCleanLiteral(psAccountTelephone) & "'"
	End If
	If Len(psAccountFax) = 0 Then
		lsSQL = lsSQL & ", AccountFax = NULL"
	Else
		lsSQL = lsSQL & ", AccountFax = '" & DBCleanLiteral(psAccountFax) & "'"
	End If
	lsSQL = lsSQL & ", PaymentTermsID = " & DBCleanLiteral(pnPaymentTermsID)
	lsSQL = lsSQL & ", CreditLimit = " & DBCleanLiteral(pdCreditLimit)
	If Len(psAccountNotes) = 0 Then
		lsSQL = lsSQL & ", AccountNotes = NULL"
	Else
		lsSQL = lsSQL & ", AccountNotes = '" & DBCleanLiteral(psAccountNotes) & "'"
	End If
	lsSQL = lsSQL & " where AccountID = " & pnAccountID
	If DBExecuteSQL(lsSQL) Then
		lbRet = TRUE
	End If

	UpdateAccount = lbRet
End Function

' **************************************************************************
' Function: GetAccountContact
' Purpose: Retrieves account contact details.
' Parameters:	pnAccountID - The CustomerID to search for
'				psAccountName - The account name
'				psPrimaryContactName - The primary contact name
'				psPrimaryContactEmail - The primary contact email
'				psSMSEmail - The email for SMS
' Return: True if sucessful, False if not
' **************************************************************************
Function GetAccountContact(ByVal pnAccountID, ByRef psAccountName, ByRef psPrimaryContactName, ByRef psPrimaryContactEmail, ByRef psSMSEmail)
	Dim lbRet, lsSQL, loRS, loRS2

	lbRet = FALSE

	lsSQL = "select AccountName, PrimaryContactName, PrimaryContactEmail, CellPhoneCarrierID, NotificationTypeID, PrimaryContactTelephone from tblAccounts where AccountID = " & pnAccountID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lbRet = TRUE

			psAccountName = loRS("AccountName")
			psPrimaryContactName = loRS("PrimaryContactName")
			psPrimaryContactEmail = loRS("PrimaryContactEmail")

			If Not IsNull(loRS("CellPhoneCarrierID")) And Not IsNull(loRS("NotificationTypeID")) Then
				If loRS("NotificationTypeID") = 3 Then
					lsSQL = "select CellPhoneCarrierEmail from tlkpCellPhoneCarriers where CellPhoneCarrierID = " & loRS("CellPhoneCarrierID")
					If DBOpenQuery(lsSQL, FALSE, loRS2) Then
						If Not loRS2.bof And Not loRS2.eof Then
							If Not IsNull(loRS2("CellPhoneCarrierEmail")) Then
								If Len(Trim(loRS2("CellPhoneCarrierEmail"))) > 0 Then
									psSMSEmail = Replace(Replace(Replace(Replace(loRS("PrimaryContactTelephone"), "-", ""), "(", ""), ")", ""), " ", "") & "@" & Trim(loRS2("CellPhoneCarrierEmail"))
								Else
									psSMSEmail = ""
								End If
							Else
								psSMSEmail = ""
							End If
						Else
							psSMSEmail = ""
						End If
					Else
						psSMSEmail = ""
					End If
				Else
					psSMSEmail = ""
				End If
			End If
		End If

		DBCloseQuery loRS
	End If

	GetAccountContact = lbRet
End Function

' **************************************************************************
' Function: AddCustomerPhoneName
' Purpose: Adds a new customer phone/name only if it doesn't exist.
' Parameters:	psFirstName - The first name
'				psLastName - The last name
'				psCellPhone - The cell phone number
' Return: The new CustomerID
' **************************************************************************
Function AddCustomerPhoneName(ByVal psFirstName, ByVal psLastName, ByVal psCellPhone)
	Dim lnRet, lsSQL, loRS

	lnRet = 0

	lsSQL = "select CustomerID from tblCustomers where CellPhone = '" & DBCleanLiteral(psCellPhone) & "' And FirstName "
	If Len(psFirstName) = 0 Then
		lsSQL = lsSQL & "IS NULL"
	Else
		lsSQL = lsSQL & "= '" & DBCleanLiteral(psFirstName) & "'"
	End If
	lsSQL = lsSQL & " And LastName "
	If Len(psLastName) = 0 Then
		lsSQL = lsSQL & "IS NULL"
	Else
		lsSQL = lsSQL & "= '" & DBCleanLiteral(psLastName) & "'"
	End If
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lnRet = loRS("CustomerID")
		Else
			loRS.Close

			lsSQL = "EXEC AddCustomer @pEMail = "
			lsSQL = lsSQL & "NULL"
			lsSQL = lsSQL & ", @pPassword = "
			lsSQL = lsSQL & "NULL"
			lsSQL = lsSQL & ", @pFirstName = "
			If Len(psFirstName) = 0 Then
				lsSQL = lsSQL & "NULL"
			Else
				lsSQL = lsSQL & "'" & DBCleanLiteral(psFirstName) & "'"
			End If
			lsSQL = lsSQL & ", @pLastName = "
			If Len(psLastName) = 0 Then
				lsSQL = lsSQL & "NULL"
			Else
				lsSQL = lsSQL & "'" & DBCleanLiteral(psLastName) & "'"
			End If
			lsSQL = lsSQL & ", @pBirthdate = "
			lsSQL = lsSQL & "NULL"
			lsSQL = lsSQL & ", @pPrimaryAddressID = 1"
			lsSQL = lsSQL & ", @pHomePhone = "
			lsSQL = lsSQL & "NULL"
			lsSQL = lsSQL & ", @pCellPhone = "
			lsSQL = lsSQL & "'" & DBCleanLiteral(psCellPhone) & "'"
			lsSQL = lsSQL & ", @pWorkPhone = "
			lsSQL = lsSQL & "NULL"
			lsSQL = lsSQL & ", @pFAXPhone = "
			lsSQL = lsSQL & "NULL"
			lsSQL = lsSQL & ", @pIsEmailList = 0"
			lsSQL = lsSQL & ", @pIsTextList = 0"
			If DBOpenQuery(lsSQL, FALSE, loRS) Then
				If Not loRS.bof And Not loRS.eof Then
					lnRet = CLng(loRS(0))
				End If

				DBCloseQuery loRS
			End If
		End If
	End If

	AddCustomerPhoneName = lnRet
End Function

' **************************************************************************
' Function: IsAccountCollegeDebit
' Purpose: Determines if an account is a college debit account.
' Parameters:	pnAccountID - The AccountID to search for
' Return: True or False
' **************************************************************************
Function IsAccountCollegeDebit(ByVal pnAccountID)
	Dim lbRet, lsSQL, loRS

	lbRet = FALSE

	lsSQL = "select CollegeDebitStoreID from tblAccounts where CollegeDebitStoreID is not null and AccountID = " & pnAccountID
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			lbRet = TRUE
		End If

		DBCloseQuery loRS
	End If

	IsAccountCollegeDebit = lbRet
End Function

' **************************************************************************
' Function: GetLastCustomerOrder
' Purpose: Returns the last customer order.
' Parameters:	pnCustomerID - The CustomerID to find
' Return: The OrderID found or 0 if never ordered
' **************************************************************************
Function GetLastCustomerOrder(ByVal pnCustomerID)
	Dim lnRet, lsSQL, loRS, lnPos

	lnRet = 0

	lsSQL = "select MAX(OrderID) as LastOrderID from tblOrders where CustomerID = " & pnCustomerID & " and IsPaid <> 0 and OrderStatusID = 10"
	If DBOpenQuery(lsSQL, FALSE, loRS) Then
		If Not loRS.bof And Not loRS.eof Then
			If Not IsNull(loRS("LastOrderID")) Then
				lnRet = loRS("LastOrderID")
			End If
		End If

		DBCloseQuery loRS
	End If

	GetLastCustomerOrder = lnRet
End Function
%>