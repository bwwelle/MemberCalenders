<%

Function GetForums()

	Dim rsForums
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsForums = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT id, [Name], Code, Password, address, city, state, zip, " & _
		   "country, phone1, phone2, fax, email, url, contact_first_name, " & _
		   "contact_last_name, active, approved, date_approved, [notes] " & _
		   "FROM Forums " & _
		   "ORDER BY Name"

	rsForums.Open sSQL, conn
	
	set GetForums = rsForums
	
End Function

Function GetActiveForums()

	Dim rsForums
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsForums = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT id, [Name], Code, Password, address, city, state, zip, " & _
		   "country, phone1, phone2, fax, email, url, contact_first_name, " & _
		   "contact_last_name, active, approved, date_approved, [notes] " & _
		   "FROM Forums " & _
		   "WHERE active = true and approved = true " & _
		   "ORDER BY Name"

	rsForums.Open sSQL, conn
	
	set GetActiveForums = rsForums
	
End Function

Function GetForumCredentials(Email)

	Dim rsForum
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsForum = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT [Name] as username, Password " & _
		   "FROM Forums " & _
		   "WHERE Email = @email_address"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@email_address", adVarChar, adParamInput, Len(Email), Email)
	end with
	
	rsForum.Open oCmd
	
	set GetForumCredentials = rsForum
	
End Function

Function GetForum(fid)

	Dim rsForum
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsForum = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT id, [Name], Code, Password, address, city, state, zip, " & _
		   "country, phone1, phone2, fax, email, url, contact_first_name, " & _
		   "contact_last_name, active, approved, date_approved, [notes] " & _
		   "FROM Forums " & _
		   "WHERE id = @fid"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@fid", adInteger, adParamInput, 4, fid)
	end with
	
	rsForum.Open oCmd
	
	set GetForum = rsForum
	
End Function


Function VerifyForumName(sForumName)

	Dim rsVerifyForumName
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsVerifyForumName = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT count(1) as forumname_count " & _
		   "FROM Forums " & _
		   "WHERE Name = @ForumName"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@ForumName", adVarChar, adParamInput, Len(sForumName), sForumName)
	end with
	
	rsVerifyForumName.Open oCmd
	
	VerifyForumName = rsVerifyForumName("forumname_count")
	
End Function



Sub AddForum( ForumName, Code, Password, Address, City, sState, Zip, Country, _
			  Phone1, Phone2, Fax, Email, Url, ContactFirstName, ContactLastName, _
			  Active, Approved, DateApproved, sNotes, sFrom )

	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	sSQL = "INSERT INTO Forums([Name], [Code], [Password], [Address], [City], [State], [Zip], " & _
		   "[Country], [Phone1], [Phone2], [Fax], [Email], [Url], [Contact_First_Name], " & _
		   "[Contact_Last_Name], [Active], [Approved], [Date_Approved], [Notes]) " & _
		   "VALUES( @ForumName, @ForumCode, @Password, @Address, @City, @State, @Zip, " & _
		   "@Country, @Phone1, @Phone2, @Fax, @Email, @URL, @ContactFirstName, " & _
		   "@ContactLastName, @Active, @Approved, @DateApproved, @Notes )"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@ForumName", adVarChar, adParamInput, Len(ForumName), ForumName)
		if Len(Code) = 0 then
	    	.Parameters.Append .CreateParameter("@ForumCode", adVarChar, adParamInput, 1, null)
		else
	    	.Parameters.Append .CreateParameter("@ForumCode", adVarChar, adParamInput, Len(Code), Code)
		end if
    	.Parameters.Append .CreateParameter("@Password", adVarChar, adParamInput, Len(Password), Password)
    	.Parameters.Append .CreateParameter("@Address", adVarChar, adParamInput, Len(Address), Address)
    	.Parameters.Append .CreateParameter("@City", adVarChar, adParamInput, Len(City), City)
   		.Parameters.Append .CreateParameter("@State", adVarChar, adParamInput, Len(sState), sState)
   		.Parameters.Append .CreateParameter("@Zip", adVarChar, adParamInput, Len(Zip), Zip)

    	.Parameters.Append .CreateParameter("@Country", adVarChar, adParamInput, Len(Country), Country)
    	.Parameters.Append .CreateParameter("@Phone1", adVarChar, adParamInput, Len(Phone1), Phone1)
		if Len(Phone2) = 0 then
    		.Parameters.Append .CreateParameter("@Phone2", adVarChar, adParamInput, 1, null)
		else
    		.Parameters.Append .CreateParameter("@Phone2", adVarChar, adParamInput, Len(Phone2), Phone2)
		end if
		if Len(Fax) = 0 then
	    	.Parameters.Append .CreateParameter("@Fax", adVarChar, adParamInput, 1, null)
		else
	    	.Parameters.Append .CreateParameter("@Fax", adVarChar, adParamInput, Len(Fax), Fax)
		end if
    	.Parameters.Append .CreateParameter("@Email", adVarChar, adParamInput, Len(Email) , Email)
		if Len(url) = 0 then
	    	.Parameters.Append .CreateParameter("@URL", adVarChar, adParamInput, 1, null)
		else
	    	.Parameters.Append .CreateParameter("@URL", adVarChar, adParamInput, Len(url), url)
		end if
    	.Parameters.Append .CreateParameter("@ContactFirstName", adVarChar, adParamInput, Len(ContactFirstName), ContactFirstName)
    	.Parameters.Append .CreateParameter("@ContactLastName", adVarChar, adParamInput, Len(ContactLastName), ContactLastName)
    	.Parameters.Append .CreateParameter("@Active", adBoolean, adParamInput, 4, Active)
    	.Parameters.Append .CreateParameter("@Approved", adBoolean, adParamInput, 4, Approved)
		if DateApproved = "" then
	    	.Parameters.Append .CreateParameter("@DateApproved", adDBTimeStamp, adParamInput, 1, null)
		else
	    	.Parameters.Append .CreateParameter("@DateApproved", adDBTimeStamp, adParamInput, 16, DateApproved)
		end if
		if Len(sNotes) = 0 then
	    	.Parameters.Append .CreateParameter("@Notes", adLongVarChar, adParamInput, 1, null)
		else
	    	.Parameters.Append .CreateParameter("@Notes", adLongVarChar, adParamInput, Len(sNotes), sNotes)
		end if
	end with

	oCmd.execute
		
	Dim rsMembers
	
	set rsMembers = GetMembers("","")
	while not rsMembers.eof
		sSQL = "INSERT INTO MemberForums([MemberID], [ForumName]) " & _
			   "VALUES( @MemberID, @ForumName)"
		Do While (oCmd.Parameters.Count > 0)
        	oCmd.Parameters.Delete 0
	    Loop
		
		with oCmd                                                      
			.ActiveConnection = conn
    		.CommandText = sSQL
	    	.CommandType = adCmdText
	
	    	.Parameters.Append .CreateParameter("@MemberID", adInteger, adParamInput, 4, rsMembers("ID"))
	    	.Parameters.Append .CreateParameter("@ForumName", adVarChar, adParamInput, Len(trim(ForumName)), trim(ForumName))

			.execute
		End With
		rsMembers.MoveNext
	wend

	conn.Close()

	if sFrom <> "Import" then
		if sFrom = "" then
			response.Redirect("Forum.asp")
		else
			response.Redirect("../confirmation.asp?section=forum")
		end if
	end if

	
End Sub


Sub UpdateForum( ForumID, ForumName, Code, Password, Address, City, sState, Zip, Country, _
				 Phone1, Phone2, Fax, Email, Url, ContactFirstName, ContactLastName, _
				 Active, Approved, DateApproved, sNotes, sFrom )

	Dim sSQL
	Dim conn

	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	sSQL = "UPDATE Forums " & _
		   "SET [Name] = @ForumName, " & _
		   "[Code] = @ForumCode, " & _
		   "[Password] = @Password, " & _
		   "[Address] = @Address, " & _
		   "[City] = @City, " & _
		   "[State] = @State, " & _
		   "[Zip] = @Zip, " & _
		   "[Country] = @Country, " & _
		   "[Phone1] = @Phone1, " & _
		   "[Phone2] = @Phone2, " & _
		   "[Fax] = @Fax, " & _
		   "[Email] = @Email, " & _
		   "[Url] = @URL, " & _
		   "[Contact_First_Name] = @ContactFirstName, " & _
		   "[Contact_Last_Name] = @ContactLastName, " & _
		   "[Active] = @Active, " & _
		   "[Approved] = @Approved, " & _
		   "[Date_Approved] = @DateApproved, " & _
		   "[Notes] = @Notes " & _
		   "WHERE id = @ForumID"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@ForumName", adVarChar, adParamInput, Len(ForumName), ForumName)
		if Len(Code) = 0 then
	    	.Parameters.Append .CreateParameter("@ForumCode", adVarChar, adParamInput, 1, null)
		else
	    	.Parameters.Append .CreateParameter("@ForumCode", adVarChar, adParamInput, Len(Code), Code)
		end if
    	.Parameters.Append .CreateParameter("@Password", adVarChar, adParamInput, Len(Password), Password)
    	.Parameters.Append .CreateParameter("@Address", adVarChar, adParamInput, Len(Address), Address)
    	.Parameters.Append .CreateParameter("@City", adVarChar, adParamInput, Len(City), City)
   		.Parameters.Append .CreateParameter("@State", adVarChar, adParamInput, Len(sState), sState)
   		.Parameters.Append .CreateParameter("@Zip", adVarChar, adParamInput, Len(Zip), Zip)

    	.Parameters.Append .CreateParameter("@Country", adVarChar, adParamInput, Len(Country), Country)
    	.Parameters.Append .CreateParameter("@Phone1", adVarChar, adParamInput, Len(Phone1), Phone1)
		if Len(Phone2) = 0 then
    		.Parameters.Append .CreateParameter("@Phone2", adVarChar, adParamInput, 1, null)
		else
    		.Parameters.Append .CreateParameter("@Phone2", adVarChar, adParamInput, Len(Phone2), Phone2)
		end if
		if Len(Fax) = 0 then
	    	.Parameters.Append .CreateParameter("@Fax", adVarChar, adParamInput, 1, null)
		else
	    	.Parameters.Append .CreateParameter("@Fax", adVarChar, adParamInput, Len(Fax), Fax)
		end if
    	.Parameters.Append .CreateParameter("@Email", adVarChar, adParamInput, Len(Email) , Email)
		if Len(url) = 0 then
	    	.Parameters.Append .CreateParameter("@URL", adVarChar, adParamInput, 1, null)
		else
	    	.Parameters.Append .CreateParameter("@URL", adVarChar, adParamInput, Len(url), url)
		end if
    	.Parameters.Append .CreateParameter("@ContactFirstName", adVarChar, adParamInput, Len(ContactFirstName), ContactFirstName)
    	.Parameters.Append .CreateParameter("@ContactLastName", adVarChar, adParamInput, Len(ContactLastName), ContactLastName)
    	.Parameters.Append .CreateParameter("@Active", adBoolean, adParamInput, 4, Active)
    	.Parameters.Append .CreateParameter("@Approved", adBoolean, adParamInput, 4, Approved)
		if DateApproved = "" then
	    	.Parameters.Append .CreateParameter("@DateApproved", adDBTimeStamp, adParamInput, 1, null)
		else
	    	.Parameters.Append .CreateParameter("@DateApproved", adDBTimeStamp, adParamInput, 16, DateApproved)
		end if
		if Len(sNotes) = 0 then
	    	.Parameters.Append .CreateParameter("@Notes", adLongVarChar, adParamInput, 1, null)
		else
	    	.Parameters.Append .CreateParameter("@Notes", adLongVarChar, adParamInput, Len(sNotes), sNotes)
		end if
    	.Parameters.Append .CreateParameter("@ForumID", adInteger, adParamInput, 4, ForumID)
		
	end with

	oCmd.execute
	conn.Close()
	
	if sFrom = "" then
		response.Redirect("Forum.asp")
	else
		response.Redirect("../myforum.asp")
	end if
	
End Sub

Sub DeleteForum( ForumID, sFrom )

	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	sSQL = "DELETE FROM Forums " & _
		   "WHERE id = " & ForumID

	conn.execute sSQL
	conn.Close()
	
	if sFrom = "" then
		response.Redirect("Forum.asp")
	else
		response.Redirect("../confirmation.asp?section=forum")
	end if
	
End Sub


%>