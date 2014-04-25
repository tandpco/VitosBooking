<%
'*******************************************************
' Function SendMail
' Sends a mail message. Pass NULL if no attachments
' Returns TRUE or FALSE
'*********************************************************
Function SendMail(ByVal psFrom, ByVal psTo, ByVal psCC, ByVal psBCC, ByVal psSubject, ByVal psBody, ByVal psHTML, ByVal pasAttachments)
	Dim lbRet, cdoMessage, cdoConfig, i
	
	lbRet = FALSE
	
	set cdoMessage = Server.CreateObject("CDO.Message")
	set cdoConfig = Server.CreateObject("CDO.Configuration")
	cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = gsMailSystem
	If Not IsEmpty(gsSMTPUserID) > 0 Then
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = gsSMTPUserID
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = gsSMTPPassword
	End If
	cdoConfig.Fields.Update
	set cdoMessage.Configuration = cdoConfig
	cdoMessage.From =  psFrom
	cdoMessage.ReplyTo = psFrom
	cdoMessage.To = psTo
	If Len(psCC) > 0 Then
		cdoMessage.Cc = psCC
	End If
	If Len(psBCC) > 0 Then
		cdoMessage.Bcc = psBCC
	End If
	cdoMessage.Subject = psSubject
	cdoMessage.TextBody = psBody
	If Len(psHTML) > 0 Then
		cdoMessage.HtmlBody = psHTML
	End If
	If IsArrayInitialized(pasAttachments) Then
		For i = 0 To UBound(pasAttachments)
			If Len(Trim(pasAttachments(i))) > 0 Then
				cdoMessage.AddAttachment Trim(pasAttachments(i))
			End If
		Next
	End If
	on error resume next
	cdoMessage.Send
	if Err.Number = 0 then
		lbRet = TRUE
	end if
	set cdoMessage = Nothing
	set cdoConfig = Nothing
	
	SendMail = lbRet
End Function
%>