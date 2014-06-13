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
  If Request("id").Count > 0 Then
    QueryToJSON(Conn, "select StoreID, AddressLine1, AddressLine2, City, State, PostalCode, AddressNotes from tblAddresses where AddressID = " & Request("id")).Flush
  End If
  If Request("method").Count > 0 Then
    If Request("method") = "ByAddress" Then
      QueryToJSON(Conn, "select tblCustomers.CustomerID, FirstName, LastName,CellPhone,HomePhone,WorkPhone, (SELECT TOP 1 TransactionDate FROM tblOrders WHERE CustomerID = tblCustomers.CustomerID ORDER BY TransactionDate DESC) as LastOrderDate, (SELECT COUNT(OrderID) FROM tblOrders WHERE CustomerID = tblCustomers.CustomerID) as TotalOrders from trelCustomerAddresses inner join tblCustomers on tblCustomers.CustomerID = trelCustomerAddresses.CustomerID where AddressID = " &  Request("AddressID") & " order by tblCustomers.CustomerID").Flush
    End If
  End If
End If
%>
<!-- #Include File="../../include2/db-disconnect.asp" -->