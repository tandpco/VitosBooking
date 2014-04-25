<%

'------------------------------------------------------------------------------------------------------------------
'SQL Server Database Connection
'------------------------------------------------------------------------------------------------------------------
Dim Conn, mydb

Sub OpenSQLConn()


	mydb="PROVIDER=SQLNCLI10;DRIVER={SQL Server}; SERVER=vitossvr02.vitos.com;DATABASE=outsidedev; UID=outsidedev;PWD=sql40uts1d3d3v;" '---------------DEV

	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open mydb 
	Conn.CommandTimeout = 15000

End Sub

'------------------------------------------------------------------------------------
'Printing Routine
'------------------------------------------------------------------------------------

Function SendToPrinterX(ByVal psPrinterIP, ByVal psPrintQueue, ByVal psData)
	Dim lbRet, loPrinter
	
	lbRet = FALSE
'	On Error Resume Next
	
	Set loPrinter = CreateObject("ARLlpr.ARLlprPrinter")
	If Err.Number = 0 And IsObject(loPrinter) Then
		If loPrinter.OpenPrinter(psPrinterIP) Then
			If loPrinter.SendToPrinter(psPrintQueue, psData, Len(psData)) Then
				lbRet = True
			Else
				Response.Redirect("/error.asp?err=1")
			End If
			loPrinter.ClosePrinter()
		End If
		Set loPrinter = Nothing
	Else
	Response.Redirect("/error.asp?err=2")
	End If
	SendToPrinterX = lbRet
End Function

'------------------------------------------------------------------------------------------------------------------
'---Proper Case Function
'------------------------------------------------------------------------------------------------------------------

Function PCase(strInput)
 Dim iPosition  ' Our current position in the string (First character = 1)
 Dim iSpace     ' The position of the next space after our iPosition
 Dim strOutput  ' Our temporary string used to build the function's output
 iPosition = 1
 Do While InStr(iPosition, strInput, " ", 1) <> 0
  iSpace = InStr(iPosition, strInput, " ", 1)
  strOutput = strOutput & UCase(Mid(strInput, iPosition, 1))
  strOutput = strOutput & LCase(Mid(strInput, iPosition + 1, iSpace - iPosition))
 
  iPosition = iSpace + 1
 Loop
 
 strOutput = strOutput & UCase(Mid(strInput, iPosition, 1))
 strOutput = strOutput & LCase(Mid(strInput, iPosition + 1))
 
 PCase = strOutput
End Function
'------------------------------------------------------------------------------------------------------------------
'---Format Telephone Number Function
'------------------------------------------------------------------------------------------------------------------

Private Function MkPhoneNum(byVal number)
	Dim tmp
	number = CStr( number )
	number = Trim( number )
	number = Replace( number, " ", "" )	
	number = Replace( number, "-", "" )
	number = Replace( number, "(", "" )
	number = Replace( number, ")", "" )
	Select Case Len( number )
		Case 7
			tmp = tmp & Mid( number, 1, 3 ) & "-"
			tmp = tmp & Mid( number, 4, 4 )
		Case 10
			tmp = tmp & "(" & Mid( number, 1, 3 ) & ") "
			tmp = tmp & Mid( number, 4, 3 ) & "-"
			tmp = tmp & Mid( number, 7, 4 )
		Case 11
			tmp = tmp & Mid( number, 1, 1 ) & " "
			tmp = tmp & "(" & Mid( number, 2, 3 ) & ") "
			tmp = tmp & Mid( number, 5, 3 ) & "-"
			tmp = tmp & Mid( number, 8, 4 )
		Case Else
			MkPhoneNum = Null
			Exit Function
	End Select
	MkPhoneNum = tmp
End Function

'------------------------------------------------------------------------------------------------------------------
'---Format Social Security Function
'------------------------------------------------------------------------------------------------------------------

Private Function MkSSN(byVal SocialSecurityNumber, _
    byVal BoolAddSpacers)
	dim strIn, i, tmp, lngCt, strOut
	strIn = SocialSecurityNumber
	i = 1
	do
		tmp = Mid( strIn, i, 1 )
		If Not IsNumeric( tmp ) then
			strIn = Trim( Replace( _
			    strIn, tmp, "" ) )
		Else
			i = i + 1
		End If
	loop until i > Len( strIn )
	strIn = Trim( strIn )
	lngCt = Len( strIn )
	if lngCt <> 9 then
		MkSSN = Null
		Exit Function
	end if	
	strIn = CStr( strIn )
	strIn = CStr( Left( strIn, 3 ) & "-" & _
		Mid( strIn, 4, 2 ) & "-" & _
		Right( strIn, 4 ) )
	lngCt = Len( strIn )
	if lngCt <> 11 then
		MkSSN = Null
		Exit Function
	end if	
	If BoolAddSpacers Then
		for i = 1 to len( strIn )
			tmp = Mid(strIn, i, 1)
			strOut = strOut & tmp & " "
		next
	Else
		strOut = strOut & strIn
	End If
	MkSSN = strOut
End Function

'------------------------------------------------------------------------------------------------------------------
'--- SHOW ITEMS IN FORM FUNCTION
'------------------------------------------------------------------------------------------------------------------
Dim item

Sub ShowForm()

	For each item in Request.Form
	Response.Write item & " - "
	Response.Write Request.Form(item) & "<br>"
	Next

End Sub

'------------------------------------------------------------------------------------------------------------------
'--- SHOW ITEMS IN SESSION FUNCTION
'------------------------------------------------------------------------------------------------------------------

Sub ShowSession()

	For each item in Session.Contents
	Response.Write item & " - "
	Response.Write Session(item) & "<br>"
	Next

End Sub

'------------------------------------------------------------------------------------------------------------------
'--- SHOW ITEMS IN SERVER VARIABLES FUNCTION
'------------------------------------------------------------------------------------------------------------------

Sub ShowServervariables()

	For each item in Request.ServerVariables 
	Response.Write item & " - "
	Response.Write Request.ServerVariables(item) & "<br>"

    next 


End Sub
'------------------------------------------------------------------------------------------------------------------
'--- CLOSE DATABASE CONNECTION
'--- Close Connections
'------------------------------------------------------------------------------------------------------------------

Sub CloseConn()

	If IsObject(Conn) Then

		Conn.Close
		SET Conn = Nothing

	End If

End Sub

'------------------------------------------------------------------------------------------------------------------
'--- USED TO SHOW DATA TO A POINT ON THE PAGE THEN INDICATE CODE IS OK THEN END
'--- Debug Code
'------------------------------------------------------------------------------------------------------------------

Sub Debug()
	Response.Write "ok"
	Response.End
End Sub



%>
