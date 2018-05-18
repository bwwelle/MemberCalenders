<!-- #include file ="clMail.asp" -->
<%

Function GetPhotosForVoting()

	Dim rsPhotosForVoting
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsPhotosForVoting = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT id, photo_id_number, member_id, time_uploaded, Active, Times_Viewed,  " & _
		   "SKU, Notes " & _
		   "FROM Photos_For_Voting " & _
		   "AND Active = true " & _
		   "ORDER BY id"

	rsPhotosForVoting.Open sSQL, conn
	
	set GetPhotosForVoting = rsPhotosForVoting
	
End Function

Function GetEmailInformation(PhotoID)

	Dim rsEmailInformation
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsEmailInformation = server.CreateObject("ADODB.Recordset")

	sSQL = "SELECT Members.User, Members.Email, Photos_For_Voting.id, " & _
		   "Vote_Schedule.Round_1_Start, Vote_Schedule.Round_1_End, Vote_Schedule.Round_2_Start," & _
		   "Vote_Schedule.Round_2_End, Vote_Schedule.Round_3_Start, Vote_Schedule.Round_3_End, " & _
		   "Products.Product_Name " & _
		   "FROM ((Members " & _
		   "INNER JOIN Photos_For_Voting ON Members.id = Photos_For_Voting.Member_ID) " & _
		   "INNER JOIN Products ON Photos_For_Voting.Product_ID = Products.id) " & _
		   "INNER JOIN Vote_Schedule ON Products.SKU = Vote_Schedule.SKU " & _
		   "WHERE (((Photos_For_Voting.id) = @photoid))"

	'Response.write(sSQL)
	

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@photoid", adInteger, adParamInput, 4, PhotoID)
	end with
	
	rsEmailInformation.Open oCmd
	
	set GetEmailInformation = rsEmailInformation

End Function

Function GetUnApprovedPhotosCount(ProductID)

	Dim rsUnApprovedPhotosAccount
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsUnApprovedPhotosAccount = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT count(1) as UnApprovedCount " & _
		   "FROM Photos_For_Voting " & _
		   "WHERE product_id = @productid " & _
		   "AND Approved = false " & _
		   "AND Processed = false"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@productid", adInteger, adParamInput, 4, ProductID)
	end with
	
	rsUnApprovedPhotosAccount.Open oCmd
	
	set GetUnApprovedPhotosCount = rsUnApprovedPhotosAccount
	
End Function


Function GetUnApprovedPhotos(ProductID)

	Dim rsUnApprovedPhotos
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsUnApprovedPhotos = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT id, member_id, product_id, forum_id, Description, time_uploaded, " & _
		   "ImageName, ThumbNailName, BlowupName, Approved, Times_Viewed " & _
		   "FROM Photos_For_Voting " & _
		   "WHERE product_id = @productid " & _
		   "AND Approved = false " & _
		   "AND Processed = false " & _		   
		   "ORDER BY ImageName"


	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@productid", adInteger, adParamInput, 4, ProductID)
	end with
	
	rsUnApprovedPhotos.Open oCmd
	
	set GetUnApprovedPhotos = rsUnApprovedPhotos
	
End Function

Function GetApprovedPhotosCount(ProductID)

	Dim rsApprovedPhotosAccount
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsApprovedPhotosAccount = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT count(1) as ApprovedCount " & _
		   "FROM Photos_For_Voting " & _
		   "WHERE product_id = @productid " & _
		   "AND Approved = true " & _
		   "AND Processed = true"		   

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@productid", adInteger, adParamInput, 4, ProductID)
	end with
	
	rsApprovedPhotosAccount.Open oCmd
	
	set GetApprovedPhotosCount = rsApprovedPhotosAccount
	
End Function

Function GetApprovedPhotos(ProductID)

	Dim rsApprovedPhotos
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsApprovedPhotos = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT id, member_id, product_id, forum_id, Description, time_uploaded, " & _
		   "ImageName, ThumbNailName, BlowupName, Approved, Times_Viewed " & _
		   "FROM Photos_For_Voting " & _
		   "WHERE product_id = @productid " & _
		   "AND Approved = true " & _
		   "AND Processed = true " & _ 
		   "ORDER BY ImageName"


	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@productid", adInteger, adParamInput, 4, ProductID)
	end with
	
	rsApprovedPhotos.Open oCmd
	
	set GetApprovedPhotos = rsApprovedPhotos
	
End Function

Function GetUnApprovedProcessedPhotosCount(ProductID)

	Dim rsUnApprovedProcessedPhotosCount
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsUnApprovedProcessedPhotosCount = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT count(1) as UnApprovedProcessedCount " & _
		   "FROM Photos_For_Voting " & _
		   "WHERE product_id = @productid " & _
		   "AND Approved = false " & _
		   "AND Processed = true"		   

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@productid", adInteger, adParamInput, 4, ProductID)
	end with
	
	rsUnApprovedProcessedPhotosCount.Open oCmd
	
	set GetUnApprovedProcessedPhotosCount = rsUnApprovedProcessedPhotosCount
	
End Function

Function GetUnApprovedProcessedPhotos(ProductID)

	Dim rsUnApprovedProcessedPhotos
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsUnApprovedProcessedPhotos = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT id, member_id, product_id, forum_id, Description, time_uploaded, " & _
		   "ImageName, ThumbNailName, BlowupName, Approved, Times_Viewed " & _
		   "FROM Photos_For_Voting " & _
		   "WHERE product_id = @productid " & _
		   "AND Approved = false " & _
		   "AND Processed = true " & _ 
		   "ORDER BY ImageName"


	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@productid", adInteger, adParamInput, 4, ProductID)
	end with
	
	rsUnApprovedProcessedPhotos.Open oCmd
	
	set GetUnApprovedProcessedPhotos = rsUnApprovedProcessedPhotos
	
End Function


Function GetPhotosOfProductByMember(ProductID, MemberID)

	Dim rsPhotosOfProductByMember
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsPhotosOfProductByMember = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT id, member_id, product_id, forum_id, Description, time_uploaded, " & _
		   "ImageName, ThumbNailName, BlowupName, Approved, Times_Viewed " & _
		   "FROM Photos_For_Voting " & _
		   "WHERE product_id = @productid " & _
		   "AND member_id = @memberid " & _
		   "ORDER BY ImageName"


	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@productid", adInteger, adParamInput, 4, ProductID)
    	.Parameters.Append .CreateParameter("@memberid", adInteger, adParamInput, 4, MemberID)
	end with
	
	rsPhotosOfProductByMember.Open oCmd
	
	set GetPhotosOfProductByMember = rsPhotosOfProductByMember
	
End Function

Function GetWinningPhotos(ProductID)

	Dim rsWinningPhotos
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsWinningPhotos = server.CreateObject("ADODB.Recordset")

	sSQL = "SELECT Products.Product_Name, Products.ImagePath, Products.ThumbNailPath, " & _
		   "Products.BlowupPath, Photos_For_Voting.ImageName, Photos_For_Voting.ID, " & _
		   "Photos_For_Voting.ThumbNailName, Photos_For_Voting.BlowupName, Photos_For_Voting.Member_ID," & _
		   "Photos_For_Voting.Times_Viewed, Photos_For_Voting.Round, " & _
		   "Photos_For_Voting.IsWinner, Photos_For_Voting.Product_ID, " & _
		   "Photos_For_Voting.MasterPhotoID " & _
		   "FROM Products " & _
		   "INNER JOIN Photos_For_Voting ON " & _
		   "Products.id = Photos_For_Voting.Product_ID " & _
		   "WHERE (((Photos_For_Voting.Round)=3) " & _
		   "AND ((Photos_For_Voting.IsWinner)=True) " & _
		   "AND ((Photos_For_Voting.Product_ID)=@ProductID)) " & _
		   "ORDER BY Photos_For_Voting.Round"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")


	with oCmd

		Do While (oCmd.Parameters.Count > 0)
        	oCmd.Parameters.Delete 0
	    Loop

		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@ProductID", adInteger, adParamInput, 4, ProductID)
	end with

	rsWinningPhotos.Open oCmd
	
	set GetWinningPhotos = rsWinningPhotos

End Function


Function GetRandomWinningPhotos(ProductID)

	Dim rsWinningPhotos
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsRandomWinningPhotos = server.CreateObject("ADODB.Recordset")

	sSQL = "SELECT Top 5 Photos_For_Voting.ID, Rnd(Photos_For_Voting.ID) as RandomPhoto, " & _
		   "Photos_For_Voting.ImageName, " & _
		   "Photos_For_Voting.ThumbNailName, Photos_For_Voting.BlowupName, " & _
		   "Photos_For_Voting.Times_Viewed, Photos_For_Voting.Round, " & _
		   "Photos_For_Voting.IsWinner, Photos_For_Voting.Product_ID, " & _
		   "Photos_For_Voting.MasterPhotoID " & _
		   "FROM Photos_For_Voting " & _
		   "WHERE Photos_For_Voting.Round = 3 " & _
		   "AND Photos_For_Voting.IsWinner = True " & _
		   "AND Photos_For_Voting.Product_ID = @ProductID " & _
		   "ORDER BY Rnd(Photos_For_Voting.ID)"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")

	with oCmd

		Do While (oCmd.Parameters.Count > 0)
        	oCmd.Parameters.Delete 0
	    Loop

		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@ProductID", adInteger, adParamInput, 4, ProductID)
	end with

	rsRandomWinningPhotos.Open oCmd
	
	set GetRandomWinningPhotos = rsRandomWinningPhotos

End Function


Function GetPhotoForVoting(id)

	Dim rsPhotoForVoting
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsPhotoForVoting = server.CreateObject("ADODB.Recordset")
	
	sSQL = "SELECT id, member_id, product_id, forum_id, Description, time_uploaded, " & _
		   "ImageName, ThumbNailName, BlowupName, Approved, Times_Viewed " & _
		   "FROM Photos_For_Voting " & _
		   "WHERE ID = @id "

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")

	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@id", adInteger, adParamInput, 4, id)
	end with

	rsPhotoForVoting.Open oCmd
	
	set GetPhotoForVoting = rsPhotoForVoting

End Function

Sub AddPhotoForVoting( MemberID, ProductID, ForumID, sDescription, ImageName, ThumbNailName, _
					   BlowupName )

	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	sSQL = "INSERT INTO Photos_For_Voting([Member_ID], [Product_id], [Forum_ID], " & _
		   "[Description], [Time_Uploaded], [ImageName], [ThumbNailName], [BlowupName], " & _
		   "[Approved], [Times_Viewed], [Round]) " & _
		   "VALUES( @MemberID, @ProductID, @ForumID, @Description, @TimeUploaded, " & _
		   "@ImageName, @ThumbNailName, @BlowupName, @Approved, @TimesViewed, @Round) " 


	Dim oCmd
	Dim times_viewed
	set oCmd = Server.CreateObject("ADODB.Command")
	
	time_uploaded = CurrentTime
	times_viewed = 0
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@MemberID", adInteger, adParamInput, 4, MemberID)
    	.Parameters.Append .CreateParameter("@ProductID", adInteger, adParamInput, 4, ProductID)
    	.Parameters.Append .CreateParameter("@ForumID", adInteger, adParamInput, 4, ForumID)
		if Len(sDescription) = 0 then
	    	.Parameters.Append .CreateParameter("@Description", adLongVarChar, adParamInput, 1, null)
		else
	    	.Parameters.Append .CreateParameter("@Description", adLongVarChar, adParamInput, Len(sDescription), sDescription)
		end if
    	.Parameters.Append .CreateParameter("@TimeUploaded", adDBTimeStamp, adParamInput, 16, time_uploaded)
		.Parameters.Append .CreateParameter("@ImageName", adVarChar, adParamInput, Len(ImageName), ImageName)
		.Parameters.Append .CreateParameter("@ThumbNailName", adVarChar, adParamInput, Len(ThumbNailName), ThumbNailName)
		.Parameters.Append .CreateParameter("@BlowupName", adVarChar, adParamInput, Len(BlowUpName), BlowupName)
    	.Parameters.Append .CreateParameter("@Approved", adBoolean, adParamInput, 4, 0)
    	.Parameters.Append .CreateParameter("@TimesViewed", adInteger, adParamInput, 4, times_viewed)
    	.Parameters.Append .CreateParameter("@Round", adInteger, adParamInput, 4, 1)
	end with

	oCmd.execute

	conn.Close()

	response.Redirect("../UploadImage.asp?pid=" & ProductID)
	
End Sub

Function GetProductsWithBannedImages()

	Dim rsProductsWithBannedImages
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsProductsWithBannedImages = server.CreateObject("ADODB.Recordset")
	
	sSQL = "SELECT  Products.Product_Name, Products.id, Count(Banned_Images.id) as BannedCount " & _
		   "FROM (Products " & _
		   "INNER JOIN Photos_For_Voting ON Products.id = Photos_For_Voting.Product_ID) " & _
		   "INNER JOIN Banned_Images ON Photos_For_Voting.id = Banned_Images.Photo_ID " & _
		   "GROUP BY Products.Product_Name, Products.id"
		   
'	sSQL = "SELECT Products.Product_Name, Products.id, Count(Banned_Images.id) AS BannedCount, " & _
'		   "Products.Product_Name, Members.User " & _
'		   "FROM ((Products INNER JOIN Photos_For_Voting ON Products.id = " & _
'		   "Photos_For_Voting.Product_ID) " & _
'		   "INNER JOIN Banned_Images ON Photos_For_Voting.id = Banned_Images.Photo_ID) " & _
'		   "INNER JOIN Members ON Photos_For_Voting.Member_ID = Members.id " & _
'		   "GROUP BY Members.User, Products.Product_Name, Products.id, Products.Product_Name"
		   
		   
	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")

	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

	end with

	rsProductsWithBannedImages.Open oCmd
	
	set GetProductsWithBannedImages = rsProductsWithBannedImages


End Function

Function GetBannedImagesCount(ProductID)

	Dim rsBannedImagesCount
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsBannedImagesCount = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT count(1) as TotalCount " & _
		   "FROM (Banned_Images " & _
		   "INNER JOIN Photos_For_Voting ON Banned_Images.Photo_ID = Photos_For_Voting.id) " & _
		   "INNER JOIN Products ON Photos_For_Voting.Product_ID = Products.id " & _
		   "WHERE Products.id = @ProductID"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")

	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@ProductID", adInteger, adParamInput, 4, ProductID)
	end with

	rsBannedImagesCount.Open oCmd
	
	GetBannedImagesCount = rsBannedImagesCount("TotalCount")

End Function

Function GetBannedImages(ProductID)

	Dim rsBannedImages
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsBannedImages = server.CreateObject("ADODB.Recordset")
	
	
'	sSQL = "SELECT Banned_Images.Photo_ID, Photos_For_Voting.ImageName, " & _
'		   "Photos_For_Voting.ThumbNailName, Products.ImagePath, Products.ThumbNailPath, " & _
'		   "Products.ID, Photos_For_Voting.Approved " & _
'		   "FROM (Banned_Images " & _
'		   "INNER JOIN Photos_For_Voting ON Banned_Images.Photo_ID = Photos_For_Voting.id) " & _
'		   "INNER JOIN Products ON Photos_For_Voting.Product_ID = Products.id " & _
'		   "WHERE products.id = @ProductID"
'
	sSQL = "SELECT Banned_Images.Photo_ID, Photos_For_Voting.ImageName, " & _
		   "Photos_For_Voting.ThumbNailName, Products.ImagePath, Products.ThumbNailPath, " & _
		   "Products.id, Photos_For_Voting.Approved, Members.User, Products.Product_Name " & _
		   "FROM ((Banned_Images " & _
		   "INNER JOIN Photos_For_Voting ON Banned_Images.Photo_ID = Photos_For_Voting.id) " & _
		   "INNER JOIN Products ON Photos_For_Voting.Product_ID = Products.id) " & _
		   "INNER JOIN Members ON Photos_For_Voting.Member_ID = Members.id " & _
		   "WHERE (((Products.id)=[@ProductID]))"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")

	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@ProductID", adInteger, adParamInput, 4, ProductID)
	end with

	rsBannedImages.Open oCmd
	
	set GetBannedImages = rsBannedImages

End Function

Sub AddBannedImage( PhotoID, ReportingMemberID )

	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	sSQL = "INSERT INTO Banned_Images([Photo_ID], [Reporting_Member_ID]) " & _
		   "VALUES( @photoID, @ReportingMemberID) " 


	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@photoID", adInteger, adParamInput, 4, PhotoID)
    	.Parameters.Append .CreateParameter("@ReportingMemberID", adInteger, adParamInput, 4, ReportingMemberID)
	end with

	oCmd.execute

	conn.Close()

End Sub

Sub UpdateWinnersForRound(RoundWinnersNumber, ProductID, CurrentRound)

	Dim sSQL
	Dim conn
	Dim rsTopRoundWinners
	Dim NextRound
	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))

	NextRound = CurrentRound + 1
	'-----------------------------------------------------------------------------------
	' Get the Top X winners for the product and current round
	'-----------------------------------------------------------------------------------
	set rsTopRoundWinners = server.CreateObject("ADODB.Recordset")
	
	sSQL = "SELECT Top " & RoundWinnersNumber & " count(photo_id), Photo_ID " & _
		   "FROM votes_placed " & _
		   "WHERE product_id = @productid " & _
		   "AND Which_Round = @currentround " & _
		   "GROUP BY photo_id " & _
		   "ORDER BY count(photo_id) DESC"

	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@productid", adInteger, adParamInput, 4, ProductID)
    	.Parameters.Append .CreateParameter("@currentround", adInteger, adParamInput, 4, CurrentRound)
	end with
	response.Write("sSQL = " & sSqL & "<BR>")
	rsTopRoundWinners.Open oCmd
	
	'-----------------------------------------------------------------------------------
	' Get the winners in a comma delimited string for the next query
	'-----------------------------------------------------------------------------------
	Dim sWinners
	sWinners = ""
	while not rsTopRoundWinners.eof
		if trim(sWinners) = "" then
			sWinners = rsTopRoundWinners("Photo_ID")
		else
			sWinners = sWinners & "," & rsTopRoundWinners("Photo_ID")
		end if
		rsTopRoundWinners.MoveNext
	Wend

	Dim aWinners
	Dim i
	
	aWinners = Split(sWinners, ",")
	
	response.Write("winners = " & sWinners & "<BR>")
	For i = 0 to UBound(aWinners)
		'-----------------------------------------------------------------------------------
		' Update the wiinners in the database.
		'-----------------------------------------------------------------------------------
		sSQL = "UPDATE Photos_For_Voting " & _
			   "SET IsWinner = true " & _
			   "WHERE ID = @Winner"
	
		response.Write("winners = " & aWinners(i) & "<BR>")

		Do While (oCmd.Parameters.Count > 0)
    	   	oCmd.Parameters.Delete 0
	    Loop

		with oCmd                                                      
			.ActiveConnection = conn
			.CommandText = sSQL
			.CommandType = adCmdText
	
			.Parameters.Append .CreateParameter("@Winner", adVarChar, adParamInput, Len(aWinners(i)), aWinners(i))
		end with
	
		oCmd.execute
	
		if CurrentRound = 1 then
			clEmail_FirstRoundVotingCompleted aWinners(i)
		end if
		if CurrentRound = 2 then
			clEmail_SecondRoundVotingCompleted aWinners(i)
		end if
		
		if CurrentRound = 3 then
			clEmail_FinalRoundVotingResults aWinners(i)
		end if
	
		'-----------------------------------------------------------------------------------
		' Move the winners to the next round
		'-----------------------------------------------------------------------------------
		if CurrentRound < 3 then
			sSQL = "INSERT INTO Photos_For_Voting(Member_ID, Product_ID, Forum_ID, Description, " & _
				   "Time_Uploaded, ImageName, ThumbNailName, BlowupName, Approved, Times_Viewed, " & _
				   "Round, IsWinner, MasterPhotoID) " & _
				   "SELECT Member_ID, Product_ID, Forum_ID, Description, @currenttime, ImageName, " & _
				   "ThumbNailName, BlowupName, 1, 0, " & NextRound & ", 0, @MasterPhotoID " & _
				   "FROM Photos_For_Voting " & _
				   "WHERE ID = @Winner"
		
			Do While (oCmd.Parameters.Count > 0)
		       	oCmd.Parameters.Delete 0
		    Loop
		
			with oCmd                                                      
				.ActiveConnection = conn
				.CommandText = sSQL
				.CommandType = adCmdText
		
		    	.Parameters.Append .CreateParameter("@currenttime", adDBTimeStamp, adParamInput, 16, CurrentTime)
				.Parameters.Append .CreateParameter("@MasterPhotoID", adVarChar, adParamInput, Len(aWinners(i)), aWinners(i))
				.Parameters.Append .CreateParameter("@Winner", adVarChar, adParamInput, Len(aWinners(i)), aWinners(i))
			end with
		
			oCmd.execute
		
		end if
	Next
	
	
	'Get photoID's that did not make it.
	'-----------------------------------------------------------------------------------
	' Get the photos that didn't make the cut
	'-----------------------------------------------------------------------------------
	set rsRoundNotWinners = server.CreateObject("ADODB.Recordset")
	
	sSQL = "SELECT ID " & _
		   "FROM Photos_For_Voting " & _
		   "WHERE product_id = @productid " & _
		   "AND Round = @currentround " & _
		   "AND IsWinner = false "

	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText


		Do While (oCmd.Parameters.Count > 0)
			oCmd.Parameters.Delete 0
		Loop

    	.Parameters.Append .CreateParameter("@productid", adInteger, adParamInput, 4, ProductID)
    	.Parameters.Append .CreateParameter("@currentround", adInteger, adParamInput, 4, CurrentRound)
	end with
	rsRoundNotWinners.Open oCmd
	
	while not rsRoundNotWinners.eof
		clEmail_PhotoDidNotReceiveEnoughVotes rsRoundNotWinners("ID")
		rsRoundNotWinners.MoveNext
	wend
	
	conn.Close()
	
End Sub

Sub SetApprovalForPhoto( photo_id, approval )

	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	

	sSQL = "UPDATE Photos_For_Voting " & _
		   "SET [Approved] = @Approved, " & _
		   "[Processed] = 1 " & _
		   "WHERE id = @ID"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@Approved", adBoolean, adParamInput, 4, Approval)
    	.Parameters.Append .CreateParameter("@ID", adInteger, adParamInput, 4, photo_id)

	end with

	oCmd.execute
	conn.Close()
	
	if Approval = 1 then
		clEmail_PhotoApproval photo_id
	else
		clEmail_PhotoNotApproved photo_id
	end if

End Sub

Sub UpdatePhotoForVoting( PhotosForVotingID,photo_id_number, member_id, time_uploaded, Active, _
					   times_viewed, sku, notes )

	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	

	sSQL = "UPDATE Photos_For_Voting " & _
		   "SET [photo_id_number] = @photoIDNumber, " & _
		   "[member_id] = @MemberID, " & _
		   "[Time_Uploaded] = @TimeUploaded, " & _
		   "[Active] = @Active, " & _
		   "[Times_Viewed] = @TimesViewed, " & _
		   "[SKU] = @SKU, " & _
		   "[Notes] = @Notes " & _
		   "WHERE id = @ID"

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
    	.Parameters.Append .CreateParameter("@ID", adInteger, adParamInput, 4, PhotosForVotingID)

	end with

	oCmd.execute
	conn.Close()
	
	response.Redirect("photos_for_Voting.asp")
	
End Sub


Sub UpdateTimesViewed( id, times_viewed )

	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	times_viewed = times_viewed + 1

	sSQL = "UPDATE Photos_For_Voting " & _
		   "SET [Times_Viewed] = @TimesViewed " & _
		   "WHERE id = @ID"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@TimesViewed",adInteger,adParamInput,4,times_viewed)
    	.Parameters.Append .CreateParameter("@ID",adInteger,adParamInput,4,id)

	end with

	oCmd.execute
	conn.Close()

End Sub

Function GetPhotoVoteCountByRound(CurrentPhotoID, CurrentRound, MasterPhotoID)

	Dim rsPhotoVoteCountByRound
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsPhotoVoteCountByRound = server.CreateObject("ADODB.Recordset")

	if CLng(MasterPhotoID) = CLng(0) then
		sSQL = "SELECT count(1) as photocount " & _	
			   "FROM Votes_Placed " & _
			   "WHERE Photo_ID = @photoid " & _
			   "AND Which_Round = @currentround "
	else
		sSQL = "SELECT count(1) as photocount " & _	
			   "FROM Votes_Placed " & _
			   "WHERE Which_Round = @currentround "	 & _
			   "AND Photo_ID IN (SELECT MasterPhotoID " & _
			   "FROM Photos_For_Voting " & _
			   "WHERE Photos_For_Voting.ID = @photoid " & _
			   "AND Photos_For_Voting.IsWinner = True " & _
			   "AND Photos_For_Voting.Round = 2 )"
	end if
	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@photoid", adInteger, adParamInput, 4, CurrentPhotoID)
    	.Parameters.Append .CreateParameter("@currentround", adVarChar, adParamInput, len(CurrentRound), CurrentRound)
	end with
	
	
	
	rsPhotoVoteCountByRound.Open oCmd
	
	set GetPhotoVoteCountByRound = rsPhotoVoteCountByRound
	
End Function

Function GetMemberVotesByRound(MemberID, ProductID, CurrentRound)

	Dim rsMemberVotesByRound
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsMemberVotesByRound = server.CreateObject("ADODB.Recordset")
	
	sSQL = "SELECT count(1) as photocount " & _	
		   "FROM Votes_Placed " & _
		   "WHERE Product_ID = @ProductID " & _
		   "AND Member_ID = @MemberID " & _
		   "AND Which_Round = @currentround "

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@ProductID", adInteger, adParamInput, 4, ProductID)
    	.Parameters.Append .CreateParameter("@MemberID", adInteger, adParamInput, 4, MemberID)
    	.Parameters.Append .CreateParameter("@currentround", adInteger, adParamInput, 4, CurrentRound)
	end with
	
	rsMemberVotesByRound.Open oCmd
	
	set GetMemberVotesByRound = rsMemberVotesByRound
	
End Function



Sub PlaceVote( PhotoID, MemberID, ProductID, WhichRound )

	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	sSQL = "INSERT INTO Votes_Placed([Photo_ID], [Member_ID], [Product_id], " & _
		   "[Which_Round], [Date_Placed]) " & _
		   "VALUES( @PhotoID, @MemberID, @ProductID, @WhichRound, @DatePlaced ) "


	Dim oCmd
	Dim DatePlaced
	set oCmd = Server.CreateObject("ADODB.Command")
	
	DatePlaced = CurrentTime
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@PhotoID", adInteger, adParamInput, 4, PhotoID)
    	.Parameters.Append .CreateParameter("@MemberID", adInteger, adParamInput, 4, MemberID)
    	.Parameters.Append .CreateParameter("@ProductID", adInteger, adParamInput, 4, ProductID)
    	.Parameters.Append .CreateParameter("@WhichRound", adInteger, adParamInput, 4, WhichRound)
    	.Parameters.Append .CreateParameter("@TimeUploaded", adDBTimeStamp, adParamInput, 16, DatePlaced)
	end with

	oCmd.execute

	conn.Close()

	'response.Redirect("../UploadImages.asp?pid=" & ProductID)
	
End Sub



Sub DeletePhotoForVoting( pid )

	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	sSQL = "DELETE FROM Photos_For_Voting " & _
		   "WHERE id = " & pid

	conn.execute sSQL
	conn.Close()
	
	response.Redirect("photos_for_Voting.asp")
	
End Sub

%>