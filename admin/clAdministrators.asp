<%

Function GetAdministrators()

	Dim rsAdministrators
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsAdministrators = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT [id], [First_Name], [Last_Name], [User], [Password], [Email] " & _
		   "FROM adminInfo " & _
		   "ORDER BY [Last_Name]"

	rsAdministrators.Open sSQL, conn
	
	set GetAdministrators = rsAdministrators
	
End Function

Function GetAdministrator(adminID)

	Dim rsAdministrator
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsAdministrator = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT [id], [First_Name], [Last_Name], [User], [Password], [Email] " & _
		   "FROM adminInfo " & _
		   "WHERE id = " & adminID

	rsAdministrator.Open sSQL, conn
	
	set GetAdministrator = rsAdministrator
	
End Function

Sub AddAdministrator( FirstName, LastName, User, Password, Email )

	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	sSQL = "INSERT INTO adminInfo([First_Name], [Last_Name], [User], [Password], [Email]) " & _
		   "VALUES( @FirstName, @LastName, @sUser, @sPassword, @Email)"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@FirstName", adVarChar, adParamInput, Len(FirstName) , FirstName)
    	.Parameters.Append .CreateParameter("@LastName", adVarChar, adParamInput, Len(LastName), LastName)
    	.Parameters.Append .CreateParameter("@sUser", adVarChar, adParamInput, Len(User), User)
    	.Parameters.Append .CreateParameter("@sPassword", adVarChar, adParamInput, Len(Password), Password)
    	.Parameters.Append .CreateParameter("@Email", adVarChar, adParamInput, Len(Email), Email)
	end with

	oCmd.execute

	'conn.execute sSQL
	conn.Close()

	response.Redirect("administrators.asp")
	
End Sub


Sub UpdateAdministrator( adminID, First_Name, Last_Name, User, Password, Email )

	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	sSQL = "UPDATE adminInfo " & _
		   "SET [First_Name] = '" & First_Name & "', " & _
		   "[Last_Name] = '" & Last_Name & "', " & _
		   "[User] = '" & User & "', " & _
		   "[Password] = '" & Password & "', " & _
		   "[Email] = '" & Email & "' " & _
		   "WHERE id = " & adminID

	conn.execute sSQL
	conn.Close()
	
	response.Redirect("administrators.asp")
	
End Sub

Sub DeleteAdministrator( adminID )

	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	sSQL = "DELETE FROM adminInfo " & _
		   "WHERE id = " & adminID

	conn.execute sSQL
	conn.Close()
	
	response.Redirect("Administrators.asp")
	
End Sub


%>