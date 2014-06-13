<%
' **************************************************************************
' File: clearorder.asp
' Purpose: Clears all order related session variables.
' Created: 7/19/2011 - TAM
' Description:
'	Include this file where needed to clear all order related session variables.
'
' Revision History:
' 7/19/201 - Created
' **************************************************************************

Session("QuickMode") = FALSE
Session("OrderID") = 0
Session("SessionID") = Session.SessionID
Session("IPAddress") = Request.ServerVariables("REMOTE_ADDR")
Session("OrderEmpID") = 0
Session("RefID") = ""
Session("SubmitDate") = DateValue("1/1/1900")
Session("ReleaseDate") = DateValue("1/1/1900")
Session("ExpectedDate") = DateValue("1/1/1900")
Session("IsPaid") = FALSE
Session("PaymentTypeID") = 1
Session("PaymentReference") = ""
Session("AccountID") = 0
Session("DeliveryCharge") = 0.00
Session("DriverMoney") = 0.00
Session("Tax") = 0.00
Session("Tax2") = 0.00
Session("Tip") = 0.00
Session("OrderStatusID") = 1
Session("OrderNotes") = ""
Session("OrderTotal") = 0.00
Session("OrderTypeID") = 1
Session("OrderTypeDescription") = ""
Session("CustomerID") = 1
Session("EMail") = ""
Session("FirstName") = ""
Session("LastName") = ""
Session("Birthdate") = DateValue("1/1/1900")
Session("PrimaryAddressID") = 1
Session("HomePhone") = ""
Session("CellPhone") = ""
Session("WorkPhone") = ""
Session("FAXPhone") = ""
Session("IsEmailList") = FALSE
Session("IsTextList") = FALSE
Session("CustomerName") = ""
Session("CustomerPhone") = ""
Session("AddressID") = 1
Session("Address1") = ""
Session("Address2") = ""
Session("City") = ""
Session("State") = ""
Session("PostalCode") = ""
Session("AddressNotes") = ""
Session("AddressDescription") = ""
Session("CustomerNotes") = ""
Session("OrderEdited") = FALSE
Session("NewOrder") = FALSE
Session("OrderLineCount") = 0
Session("CouponIDs") = ""
Session("DEBUGIDEALCOST") = FALSE
Session("EditReason") = ""
Session("Extension") = ""
Session("optRedirect") = ""
%>