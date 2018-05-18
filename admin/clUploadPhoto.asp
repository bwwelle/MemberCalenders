<%

Sub AddUploadPhoto( MemberID, ProductID, ForumID, sDescription, ImageName, ThumbNailName )

	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	sSQL = "INSERT INTO Photos([photo_id_number], [member_id], [Time_Uploaded], " & _
		   "[Active], [Times_Viewed], [SKU], [Notes]) " & _
		   "VALUES( @photoIDNumber, @MemberID, @TimeUploaded, @Active, @TimesViewed, @SKU, @Notes) " 


	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	time_uploaded = CurrentTime
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@photoIDNumber", adInteger, adParamInput, 4, photo_id_number)
    	.Parameters.Append .CreateParameter("@MemberID", adInteger, adParamInput, 4, member_id)
    	.Parameters.Append .CreateParameter("@TimeUploaded", adDBTimeStamp, adParamInput, 16, time_uploaded)
    	.Parameters.Append .CreateParameter("@Active", adBoolean, adParamInput, 4, Active)
    	.Parameters.Append .CreateParameter("@TimesViewed", adInteger, adParamInput, 4, times_viewed)
		.Parameters.Append .CreateParameter("@SKU", adVarChar, adParamInput, Len(SKU), SKU)
    	.Parameters.Append .CreateParameter("@Notes", adLongVarChar, adParamInput, Len(Notes), Notes)
	end with

	oCmd.execute

	conn.Close()

	response.Redirect("../UploadPhoto.asp")
	
End Sub



%>