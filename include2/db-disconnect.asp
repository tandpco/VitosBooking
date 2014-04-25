<%
' **************************************************************************
' File: db-disconnect.asp
' Purpose: Terminates a connection to the database.
' Created: 6/20/2011 - TAM
' Description:
'	Include this file at the bottom of any page where you have included
'	db-connect.asp in order to gracefully disconnect.
'
' Revision History:
' 6/20/201 - Created
' **************************************************************************

If IsObject(goDBConn) Then
	If goDBConn.State = adStateOpen Then
		goDBConn.Close
	End If
	
	Set goDBConn = Nothing
End If
%>