<!-- #Include File="../../include2/utility.asp" -->
<!-- #Include File="../../include2/globals.asp" -->
<!-- #Include File="../../include2/math.asp" -->
<!-- #Include File="../../include2/db-connect.asp" -->
<!-- #Include File="../../include2/order.asp" -->
<!-- #Include File="../../include2/customer.asp" -->
<!-- #Include File="../../include2/json.asp" -->
<%
OpenSQLConn
Response.ContentType = "application/json"
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
  Set Out = jsObject()
  Out("rest") = "what"
  Out.Flush()
End If
If Request.ServerVariables("REQUEST_METHOD") = "GET" Then

  ' GET The Address Details
  If Request("AddressID").Count > 0 Then
    QueryToJSONRow(Conn, "select StoreID, AddressLine1, AddressLine2, City, State, PostalCode, AddressNotes from tblAddresses where AddressID = " & Request("AddressID")).Flush
  End If
  If Request("method").Count > 0 Then

    ' GET City / State from ZipCode
    If Request("method") = "GetCityState" Then
      QueryToJSONRow(Conn, "select distinct City, State from tblCASSAddresses where PostalCode = " & Request("PostalCode")).Flush
    End If


  ' GET a List of Addresses for a Specific Phone #
    If Request("method") = "GetAddressesByPhone" Then
      QueryToJSON(Conn, "SELECT tblAddresses.AddressID,MAX(tblAddresses.StoreID) as StoreID,MAX(AddressLine1) as AddressLine1,MAX(AddressLine2) as AddressLine2,(SELECT COUNT(DISTINCT CustomerID) FROM trelCustomerAddresses WHERE trelCustomerAddresses.AddressID = tblAddresses.AddressID) as CustomerCount, MAX(tOrders.TransactionDate) as LastOrderDate     FROM tblCustomers     OUTER APPLY (SELECT TOP 1 * FROM tblOrders WHERE (tblOrders.CustomerID = tblCustomers.CustomerID) ORDER BY tblOrders.TransactionDate) tOrders       LEFT OUTER JOIN trelCustomerAddresses ON tOrders.CustomerID = trelCustomerAddresses.CustomerID      LEFT OUTER JOIN tblAddresses on trelCustomerAddresses.AddressID = tblAddresses.AddressID      WHERE HomePhone = '" & Request("Phone") & "' or CellPhone = '" & Request("Phone") & "' or WorkPhone = '" & Request("Phone") & "' or FAXPhone = '" & Request("Phone") & "'       GROUP BY tblAddresses.AddressID     ORDER BY MAX(tOrders.TransactionDate) DESC").Flush
    End If
  End If
End If
%>
<!-- #Include File="../../include2/db-disconnect.asp" -->