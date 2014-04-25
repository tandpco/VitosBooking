<!-- #Include File="adovbs.asp" -->
<%
' **************************************************************************
' File: db-connect.asp
' Purpose: Establishes a connection to the database.
' Created: 6/20/2011 - TAM
' Description:
'	Include this file at the top of any page where you will be performing
'	database operations. Include db-disconnect.asp at the bottom of the
'	page to gracefully disconnect. This file includes the following useful
'	database helper functions: DBCleanLiteral, DBOpenQuery, DBExecuteSQL.
'		DBCloseQuery, DBSleep
'
' Revision History:
' 6/20/201 - Created
' **************************************************************************

Dim gsDBProvider, gsDBServer, gsDBDatabase, gsDBUID, gsDBPW
Dim goDBConn, gsDBErrorMessage

gsDBProvider = "SQLNCLI10"
gsDBServer = "vitossvr02.vitos.com"
gsDBDatabase = "outsidedev"
gsDBUID = "outsidedev"
gsDBPW = "sql40uts1d3d3v"

gsDBErrorMessage = ""

On Error Resume Next

Set goDBConn = Server.CreateObject("ADODB.Connection")
If Err.Number = 0 And IsObject(goDBConn) Then
	goDBConn.Provider = gsDBProvider
	
	goDBConn.Properties("Data Source") = gsDBServer
	goDBConn.Properties("Initial Catalog") = gsDBDatabase
	goDBConn.Properties("User ID") = gsDBUID
	goDBConn.Properties("Password") = gsDBPW
		
	goDBConn.Open
	
	If Err.Number <> 0 Or goDBConn.State <> adStateOpen Then
		gsDBErrorMessage = Err.Description
		Set goDBConn = Nothing
		Response.Redirect "/error.asp?err=" & Server.URLEncode(gsDBErrorMessage)
	End If
Else
	gsDBErrorMessage = Err.Description
	Response.Redirect "/error.asp?err=" & Server.URLEncode(gsDBErrorMessage)
End If

On Error Goto 0

' **************************************************************************
' Function: DBCleanLiteral
' Purpose: Examines a literal sting and adds any neccessary escape clauses.
' Parameters: psLiteral - The literal to clean
' Return: String containing cleaned literal
' **************************************************************************
Function DBCleanLiteral(ByVal psLiteral)
	Dim lsRet, i, c
	
	lsRet = ""
	gsDBErrorMessage = ""
	
	For i = 1 to Len(psLiteral)
		c = Mid(psLiteral, i, 1)
		
		Select Case c
			Case "'" c = "''"
			Case "\" c = "\\"
		End Select
		
		lsRet = lsRet & c
	Next
	
	DBCleanLiteral = lsRet
End Function

' **************************************************************************
' Function: DBOpenQuery
' Purpose: Opens a query and returns a recordset.
' Parameters:	psSQL - The SQL statement to run
'				pbOpenRW - Flag if True open Read-Write if False Read Only
'				poRS - The recordset to be returned
' Return: True if sucessful, False if not 
' **************************************************************************
Function DBOpenQuery(ByVal psSQL, ByVal pbOpenRW, ByRef poRS)
	Dim lbRet
	
	lbRet = FALSE
	gsDBErrorMessage = ""
	
	On Error Resume Next
	
	Set poRS = Server.CreateObject("ADODB.Recordset")
	If Err.Number = 0 And IsObject(poRS) Then
		If pbOpenRW Then
			poRS.Open psSQL, goDBConn, adOpenDynamic, adLockOptimistic
		Else
			poRS.Open psSQL, goDBConn, adOpenDynamic, adLockReadOnly
		End If
		
		If Err.Number = 0 And poRS.State = adStateOpen Then
			lbRet = TRUE
		Else
			gsDBErrorMessage = Err.Description & " SQL: " & psSQL
		End If
	Else
		gsDBErrorMessage = Err.Description & " SQL: " & psSQL
	End If
	
	DBOpenQuery = lbRet
End Function

' **************************************************************************
' Function: DBExecuteSQL
' Purpose: Executes an SQL statement.
' Parameters:	psSQL - The SQL statement to run
' Return: True if sucessful, False if not 
' **************************************************************************
Function DBExecuteSQL(ByVal psSQL)
	Dim lbRet
	
	lbRet = FALSE
	gsDBErrorMessage = ""
	
	On Error Resume Next
	
	goDBConn.Execute psSQL
	If Err.Number = 0 Then
		lbRet = TRUE
	Else
		gsDBErrorMessage = Err.Description & " SQL: " & psSQL
	End If
	
	DBExecuteSQL = lbRet
End Function

' **************************************************************************
' Function: DBCloseQuery
' Purpose: Closes a query.
' Parameters:	poRS - The recordset to be closed
' Return: True if sucessful, False if not 
' **************************************************************************
Function DBCloseQuery(ByRef poRS)
	Dim lbRet
	
	lbRet = FALSE
	gsDBErrorMessage = ""
	
	On Error Resume Next
	
	poRS.Close
	If Err.Number = 0 Then
		lbRet = TRUE
	Else
		gsDBErrorMessage = Err.Description
	End If
	Set poRS= Nothing
	
	DBCloseQuery = lbRet
End Function

' **************************************************************************
' Function: DBSleep
' Purpose: Sleeps for a given amount of time. Microsoft SQL Server dependent.
' Parameters:	pnSeconds - The number of seconds to sleep
'				pnMinutes - The number of minutes to sleep
'				pnHours - The number of hours to sleep
' Return: True if sucessful, False if not 
' **************************************************************************
Function DBSleep(ByVal pnSeconds, ByVal pnMinutes, ByVal pnHours)
	Dim lbRet, lnOldTimeout, lsSQL
	
	lbRet = FALSE
	gsDBErrorMessage = ""
	
	If pnSeconds > 59 Or pnMinutes > 59 Or pnHours > 23 Then
		gsDBErrorMessage = "Invalid time specified."
	Else
		On Error Resume Next
		
		lnOldTimeout = goDBConn.commandTimeout
		goDBConn.commandTimeout = pnSeconds + 5
		
		lsSQL = "WAITFOR DELAY '" & Right("0" & pnHours, 2) & ":" & Right("0" & pnMinutes, 2) & ":" & Right("0" & pnSeconds, 2) & "'"
		
		goDBConn.Execute lsSQL,,129
		If Err.Number = 0 Then
			lbRet = TRUE
		Else
			gsDBErrorMessage = Err.Description
		End If
		
		goDBConn.commandTimeout = lnOldTimeout
	End If
	
	DBSleep = lbRet
End Function
%>
