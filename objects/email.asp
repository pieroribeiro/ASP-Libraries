<%
Class Email
	'---------------------------------------------
	'		PRIVATE VARS
	'---------------------------------------------
	Private p_EmailFrom, p_EmailTo, p_EmailCc, p_EmailCco, p_EmailSubject, p_EmailBody

	'---------------------------------------------
	'		PUBLIC PROPERTIES
	'---------------------------------------------
	Public Property Let EmailFrom(p):     p_EmailFrom 		= p: End Property
	Public Property Let EmailTo(p):       p_EmailTo 		= p: End Property
	Public Property Let EmailCc(p):       p_EmailCc 		= p: End Property
	Public Property Let EmailCco(p):      p_EmailCco 		= p: End Property
	Public Property Let EmailSubject(p):  p_EmailSubject 		= p: End Property
	Public Property Let EmailBody(p):     p_EmailBody 		= p: End Property

	'---------------------------------------------
	'		CLASS INITIALIZE AND DESTROYER
	'---------------------------------------------
	Public Sub Class_Initialize(): End Sub
	Public Sub Class_Terminate(): End Sub

	'---------------------------------------------
	'		PUBLIC METHODS
	'---------------------------------------------
	Public Function Send()
		On Error Resume Next
		sch = "http://schemas.microsoft.com/cdo/configuration/"
		Set cdoConfig = Server.CreateObject("CDO.Configuration")
		cdoConfig.Fields.Item(sch & "sendusing") 	= 2
		cdoConfig.Fields.Item(sch &"smtpserverport") 	= 587
		cdoConfig.Fields.Item(sch & "smtpserver") 	= "localhost"
		cdoConfig.fields.update
		Set cdoMessage = Server.CreateObject("CDO.Message")
		Set cdoMessage.Configuration = cdoConfig
		cdoMessage.From		    	= p_EmailFrom
		cdoMessage.Subject		= p_EmailSubject
		cdoMessage.To			= p_EmailTo
		cdoMessage.Cc			= p_EmailCc
		cdoMessage.Bcc			= p_EmailCco
		cdoMessage.HTMLBody		= p_EmailBody
		cdoMessage.Send
		Set cdoMessage 			= Nothing
		Set cdoConfig 			= Nothing

		If Err <> 0 Then: Send = false: Else: Send = true: End If
	End Function
End Class
%>
