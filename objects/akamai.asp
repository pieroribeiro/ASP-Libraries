<%
Class Akamai
	Private XMLHttp
	Private AkamaiUser
	Private AkamaiPass
	Private emailNotification
	
	Public Sub Class_Initialize()
		user = "YOUR_AKAMAI_USERNAME"
		pass = "YOUR_AKAMAI_PASSWORD"
		emailNotification = "your_email@your_domain.com"
  		Set XMLHttp 			= Server.CreateObject("MSXML2.ServerXMLHTTP")
		XMLHttp.setTimeouts 30000, 60000, 40000, 40000
	End Sub	
	Public Sub Class_Terminate()
		If IsObject(XMLHttp) Then
			Set XMLHttp 			= Nothing
		End If
	End Sub	
	Public Function Purge(action, typeOfPurge, domain, urls)
		Dim XML, retorno, url
		XML = "<?xml version=""1.0"" encoding=""UTF-8""?>"
		XML = XML &"<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:ns1=""http://www.akamai.com/purge"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" SOAP-ENV:encodingStyle=""http://schemas.xmlsoap.org/soap/encoding/"">"
		XML = XML &"<SOAP-ENV:Body>"
		XML = XML &"<ns1:purgeRequest>"
		XML = XML &"<name xsi:type=""xsd:string"">"& AkamaiUser &"</name>"
		XML = XML &"<pwd xsi:type=""xsd:string"">"& AkamaiPass &"</pwd>"
		XML = XML &"<network xsi:type=""xsd:string""></network>"
		XML = XML &"<opt SOAP-ENC:arrayType=""xsd:string[0]"" xsi:type=""ns1:ArrayOfString"">"
		XML = XML &"<item xsi:type=""xsd:string"">action="& action &"</item>"
		XML = XML &"<item xsi:type=""xsd:string"">type="& typeOfPurge &"</item>"
		XML = XML &"<item xsi:type=""xsd:string"">domain=production</item>"
		XML = XML &"<item xsi:type=""xsd:string"">email-notification="& emailNotification &"</item>"
		XML = XML &"</opt>"
		If urls <> "" Then
			XML = XML &"<uri SOAP-ENC:arrayType=""xsd:string[1]"" xsi:type=""ns1:ArrayOfString"">"
			url = Split(urls,",")
			If IsArray(url) Then
				For a = 0 To UBound(url)
					XML = XML &"<item xsi:type=""xsd:string"">"& url(a) &"</item>"
				Next
			End If
			XML = XML &"</uri>"
		End If
		XML = XML &"</ns1:purgeRequest>"
		XML = XML &"</SOAP-ENV:Body>"
		XML = XML &"</SOAP-ENV:Envelope>"
		
		XMLHttp.Open 				"POST", "https://ccuapi.akamai.com:443/soap/servlet/soap/purge", Null, AkamaiUser, AkamaiPass
		XMLHttp.setRequestHeader	"SOAPAction", "purgeRequest"
		XMLHttp.SetRequestHeader 	"Content-Type", "text/xml;charset=UTF-8"
		XMLHttp.Send 				XML
		retorno						= XMLHttp.ResponseText
		
		Purge = retorno
	End Function
End Class
%>
