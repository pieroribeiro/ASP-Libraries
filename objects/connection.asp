<%
Class Connection
	'---------------------------------------------
	'		VARS
	'---------------------------------------------
	Private ServerDB
	Private ServerUSR
	Private ServerPWD
	Private ServerProvider

	Private tmp_conn
	Private tmp_connString

	Public Property Get Connection()
		Connection 			= tmp_conn
	End Property
	Public Property Get ConnectionString()
		ConnectionString 	= tmp_connString
	End Property

	'---------------------------------------------
	'		CLASS INITIALIZE AND DESTROYER
	'---------------------------------------------
	Private Sub Class_Initialize()
		ServerDB		= "YOUR_SERVER_IP"
		ServerUSR		= "YOUR_SERVER_USER"
		ServerPWD		= "YOUR_SERVER_PASSWORD"
		ServerProvider	= "SQLNCLI11"
		Set tmp_conn 	= Server.CreateObject("ADODB.Connection")
	End Sub

	Private Sub Class_Terminate()
		Set tmp_conn 	= Nothing
		tmp_conn 		= Empty
	End Sub

	Private Function SetConnectionString(Database)
		Dim ConnString
		ConnString 		= "Provider="& ServerProvider &";Persist Security Info=True;Data Source=tcp:"& ServerDB &";Database="& Database &";User ID="& ServerUSR &";PWD="& ServerPWD
		tmp_connString 	= ConnString
		SetConnectionString = ConnString
	End Function

	'---------------------------------------------
	'		PUBLIC METHODS
	'---------------------------------------------
	Public Function Open(Database)
		Dim connStr: connStr = SetConnectionString(Database)
		tmp_conn.Open 	connStr
	End Function

	Public Function Close()
		If (tmp_conn.State = 1) Then
			tmp_conn.Close()
		End If
	End Function

	Public Function Execute(SQL)
		Dim tmp_rs
		If (tmp_conn.State = 1) Then
			Set tmp_rs = tmp_conn.Execute(SQL)
			If tmp_rs.state = 1 Then
				If Not tmp_rs.Eof Then
					Execute = tmp_rs.getRows()
				Else
					Execute = Null
				End If
			Else
				Execute = Null
			End If
			Set tmp_rs = Nothing
		Else
			Execute = Null
		End If
	End Function
End Class
%>
