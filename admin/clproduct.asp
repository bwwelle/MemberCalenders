<%

Function GetProducts()

	Dim rsProducts
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsProducts = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT id, sku, product_type, product_name, description, price,  " & _
		   "thumbnail, large_image, active, related_items, in_stock, forum_id, notes, " & _
		   "DateCreated, ImagePath, ThumbNailPath, BlowupPath " & _
		   "FROM Products " & _
		   "ORDER BY Product_Name"

	rsProducts.Open sSQL, conn
	
	set GetProducts = rsProducts
	
End Function

Function GetProductsAvailableForVotingSchedule(sSKU)

	Dim rsProductsAvailableForVotingSchedule
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsProductsAvailableForVotingSchedule = server.CreateObject("ADODB.Recordset")
	
	if Len(sSKU) = 0 then
		sSQL = "SELECT Products.SKU, Products.Product_Name " & _
			   "FROM Products " & _
			   "WHERE Products.SKU not in (SELECT SKU FROM Vote_Schedule)"
	else
		sSQL = "SELECT Products.SKU, Products.Product_Name " & _
			   "FROM Products " & _
			   "WHERE Products.SKU in (SELECT SKU FROM Vote_Schedule WHERE SKU = @sku)"
	end if
	
	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText
	
		if Len(sSKU) > 0 then
	    	.Parameters.Append .CreateParameter("@sku", adVarChar, adParamInput, len(sSKU), sSKU)
		end if
	end with
	
	rsProductsAvailableForVotingSchedule.Open oCmd
	
	set GetProductsAvailableForVotingSchedule = rsProductsAvailableForVotingSchedule
	
End Function


Function GetActiveProducts()

	Dim rsActiveProducts
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsActiveProducts = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT id, sku, product_type, product_name, description, price,  " & _
		   "thumbnail, large_image, active, related_items, in_stock, forum_id, notes, " & _
		   "DateCreated, ImagePath, ThumbNailPath, BlowupPath " & _
		   "FROM Products " & _
		   "WHERE active = true " & _
		   "ORDER BY Product_Name"

	rsActiveProducts.Open sSQL, conn
	
	set GetActiveProducts = rsActiveProducts
	
End Function


Function GetProduct(pid)

	Dim rsProduct
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsProduct = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT id, sku, product_type, product_name, description, price,  " & _
		   "thumbnail, large_image, active, related_items, in_stock, forum_id, notes, " & _
		   "DateCreated, ImagePath, ThumbNailPath, BlowupPath " & _
		   "FROM Products " & _
		   "WHERE ID = @pid"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@pid", adInteger, adParamInput, 4, pid)
	end with
	
	rsProduct.Open oCmd
	
	set GetProduct = rsProduct
	
End Function

Function GetProductBySku(sku)

	Dim rsProductBySku
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsProductBySku = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT count(id) as skucount " & _
		   "FROM Products " & _
		   "WHERE Sku = @sku"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@sku", adVarChar, adParamInput, len(sku), sku)
	end with
	
	rsProductBySku.Open oCmd
	
	Dim skucount
	
	skucount = rsProductBySku("skucount")
	
	GetProductBySku = skucount
	
End Function

Function GetProductDetailBySku(sku)

	Dim rsProductDetailBySku
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsProductDetailBySku = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT id, sku, product_type, product_name, description, price,  " & _
		   "thumbnail, large_image, active, related_items, in_stock, forum_id, notes, " & _
		   "DateCreated, ImagePath, ThumbNailPath, BlowupPath " & _
		   "FROM Products " & _
		   "WHERE Sku = @sku"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@sku", adVarChar, adParamInput, len(sku), sku)
	end with
	
	rsProductDetailBySku.Open oCmd
	
	set GetProductDetailBySku = rsProductDetailBySku
	
End Function

Function GetProductsToVoteForHomePage()

	Dim rsProductsToVoteForHomePage
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsProductsToVoteForHomePage = server.CreateObject("ADODB.Recordset")
	
	sSQL = "SELECT Top 4 Products.id, Products.Product_Name, Products.Forum_Id, Products.Thumbnail, " & _
		   "Products.Large_Image, Products.ImagePath,  Vote_Schedule.Round_3_End, " & _
		   "Vote_Schedule.Round_3_Completed, Forums.Name " & _
		   "FROM (Products INNER JOIN Vote_Schedule ON Products.SKU = Vote_Schedule.SKU) " & _
		   "INNER JOIN Forums ON Products.Forum_Id = Forums.id " & _
		   "WHERE (((Vote_Schedule.Round_1_Start) < @CurrentTime1) " & _
		   "AND ((Vote_Schedule.Round_3_Completed) is null)) " & _
		   "AND Products.Active = true " & _
		   "ORDER BY Products.id"
		   
	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText
		
    	.Parameters.Append .CreateParameter("@CurrentTime1", adDBTimeStamp, adParamInput, 16, CurrentTime)
	end with
	
	rsProductsToVoteForHomePage.Open oCmd
	
	set GetProductsToVoteForHomePage = rsProductsToVoteForHomePage

End Function


Function GetProductsToSell()

	Dim rsProductsToSell
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsProductsToSell = server.CreateObject("ADODB.Recordset")
	
	sSQL = "SELECT Products.Product_Name, Products.Forum_Id, Products.Thumbnail, " & _
		   "Products.Large_Image, Products.ImagePath, Products.id, Vote_Schedule.Round_3_End, " & _
		   "Products.Price, Products.In_Stock, Vote_Schedule.Round_3_Completed, Forums.Name " & _
		   "FROM (Products INNER JOIN Vote_Schedule ON Products.SKU = Vote_Schedule.SKU) " & _
		   "INNER JOIN Forums ON Products.Forum_Id = Forums.id " & _
		   "ORDER BY Vote_Schedule.Round_3_Completed DESC"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText
		
    	'.Parameters.Append .CreateParameter("@CurrentTime1", adDBTimeStamp, adParamInput, 16, CurrentTime)

	end with
	
	rsProductsToSell.Open oCmd
	
	set GetProductsToSell = rsProductsToSell

End Function


Function GetProductsForHomePage()

	Dim rsProductsForHomePage
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsProductsForHomePage = server.CreateObject("ADODB.Recordset")
	
	sSQL = "SELECT Top 4 Products.id, Products.Product_Name, Products.Forum_Id, Products.Thumbnail, " & _
		   "Products.Large_Image, Products.ImagePath, Vote_Schedule.Round_3_End, " & _
		   "Vote_Schedule.Round_3_Completed, Forums.Name " & _
		   "FROM (Products INNER JOIN Vote_Schedule ON Products.SKU = Vote_Schedule.SKU) " & _
		   "INNER JOIN Forums ON Products.Forum_Id = Forums.id " & _
		   "WHERE (((Vote_Schedule.Round_1_Start) < @CurrentTime1) " & _
		   "AND ((Vote_Schedule.Round_3_Completed) < @CurrentTime2)) " & _
		   "ORDER BY Products.id"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText
		
    	.Parameters.Append .CreateParameter("@CurrentTime1", adDBTimeStamp, adParamInput, 16, CurrentTime)
    	.Parameters.Append .CreateParameter("@CurrentTime2", adDBTimeStamp, adParamInput, 16, CurrentTime)

	end with
	
	rsProductsForHomePage.Open oCmd
	
	set GetProductsForHomePage = rsProductsForHomePage

End Function

Function GetProductsNotInStockForHomePage()

	Dim rsProductsNotInStockForHomePage
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsProductsNotInStockForHomePage = server.CreateObject("ADODB.Recordset")
	
	sSQL = "SELECT Top 4 Products.id, Products.Product_Name, Products.Forum_Id, Products.Thumbnail, Forums.Name,  " & _
		   "Products.Large_Image, Products.ImagePath " & _
		   "FROM Products " & _
		   "INNER JOIN Forums ON Products.Forum_ID = Forums.ID " & _
		   "WHERE In_Stock = false " & _
		   "ORDER BY Products.id"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText
		
    	.Parameters.Append .CreateParameter("@CurrentTime1", adDBTimeStamp, adParamInput, 16, CurrentTime)
    	.Parameters.Append .CreateParameter("@CurrentTime2", adDBTimeStamp, adParamInput, 16, CurrentTime)

	end with
	
	rsProductsNotInStockForHomePage.Open oCmd
	
	set GetProductsNotInStockForHomePage = rsProductsNotInStockForHomePage

End Function

Function GetProductWithForum(pid)

	Dim rsProductWithForum
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsProductWithForum = server.CreateObject("ADODB.Recordset")
	
	sSQL = "SELECT Forums.id, Forums.Name, Forums.Code, Products.id, Products.SKU, " & _
		   "Products.Product_Type, Products.Product_Name, Products.Description, " & _
		   "Products.Price, Products.Active, Products.In_Stock, Products.Forum_Id, " & _
		   "Products.Notes, Products.ImagePath, Products.ThumbNailPath, Products.BlowupPath, " & _ 	
		   "Vote_Schedule.id as voteid, Vote_Schedule.Round_1_Start, " & _
		   "Vote_Schedule.Round_1_End, Vote_Schedule.Round_2_Start, " & _
		   "Vote_Schedule.Round_2_End, Vote_Schedule.Round_3_Start, Vote_Schedule.Round_3_End, " & _
		   "Vote_Schedule.Round_1_Number_of_Votes, Vote_Schedule.Round_2_Number_of_Votes, " & _
		   "Vote_Schedule.Round_3_Number_of_Votes, Vote_Schedule.Round_1_Winners, " & _
		   "Vote_Schedule.Round_2_Winners, Vote_Schedule.Round_3_Winners, " & _
		   "Vote_Schedule.Cut_off_Date " & _
		   "FROM (Products INNER JOIN Forums ON Products.Forum_Id = Forums.id) " & _
		   "INNER JOIN Vote_Schedule ON Products.SKU = Vote_Schedule.SKU " & _
		   "WHERE Products.ID = @pid " & _
		   "AND Approved = true"
		   

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@pid", adInteger, adParamInput, 4, pid)
	end with
	
	rsProductWithForum.Open oCmd
	
	set GetProductWithForum = rsProductWithForum
	
End Function


Function GetProductByForum(ForumID)

	Dim rsProduct
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsProduct = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT id, sku, product_type, product_name, description, price,  " & _
		   "thumbnail, large_image, active, related_items, in_stock, forum_id, notes, " & _
		   "DateCreated " & _
		   "FROM Products " & _
		   "WHERE forum_id = @forumid"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@forumid", adInteger, adParamInput, 4, ForumID)
	end with
	
	rsProduct.Open oCmd
	
	set GetProductByForum = rsProduct
	
End Function

Function GetProductBySku(sSKU)

	Dim rsGetProductID
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsGetProductID = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT id  " & _
		   "FROM Products " & _
		   "WHERE SKU = @SKU"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@SKU", adVarChar, adParamInput, len(sSKU), sSKU)
	end with
	
	rsGetProductID.Open oCmd
	if rsGetProductID.eof then
		GetProductBySku = 0
	else
		GetProductBySku = rsGetProductID("ID")
	end if

End Function


Function GetPhotosForProduct(productid, SortBy, totalrecords, searchcriteria, CurrentRound)

	Dim rsPhotosForProduct
	Dim sSQL
	Dim conn
	
	if trim(SortBy) = "" then
		SortBy = "ORDER BY time_uploaded desc"
	else
	 	SortBy = "ORDER BY Times_Viewed desc"
	end if

	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsPhotosForProduct = server.CreateObject("ADODB.Recordset")
	
	'Set the RecordSet PageSize property
	rsPhotosForProduct.PageSize = totalrecords

	'Set the RecordSet CacheSize property to the
	'number of records that are returned on each page of results
	rsPhotosForProduct.CacheSize = totalrecords
		   
	sSQL = "SELECT Photos_For_Voting.id, Photos_For_Voting.Member_ID, " & _
		   "Photos_For_Voting.Product_ID, Photos_For_Voting.Forum_ID, " & _
		   "Photos_For_Voting.Description AS photo_description,Photos_For_Voting.Time_Uploaded, " & _
		   "Photos_For_Voting.ImageName, Photos_For_Voting.ThumbNailName, " & _
		   "Photos_For_Voting.BlowupName, Photos_For_Voting.Approved, " & _
		   "Photos_For_Voting.Times_Viewed, Forums.Name, Forums.Code, Products.Product_Name, " & _
		   "Products.Description AS product_description, Members.User " & _
		   "FROM ((Forums INNER JOIN Photos_For_Voting ON Forums.id = Photos_For_Voting.Forum_ID) " & _
		   "INNER JOIN Products ON Photos_For_Voting.Product_ID = Products.id) " & _
		   "INNER JOIN Members ON Photos_For_Voting.Member_ID = Members.id " & _
		   "WHERE (((Photos_For_Voting.Product_ID) = " & ProductID & ") " & _
		   "AND ((Photos_For_Voting.Approved) = True) " & _
		   "AND ((Photos_For_Voting.Round) = " & CurrentRound & ")) "
		   
	if searchcriteria <> "" then
		sSQL = sSQL & "AND (Members.User like '%" & searchcriteria & "%' " & _
			   "OR Photos_For_Voting.Description like '%" & searchcriteria & "%') " & SortBy
	else
		sSQL = sSQL & SortBy
	end if

	rsPhotosForProduct.Open sSQL, conn, adOpenStatic, adLockReadOnly

	if not rsPhotosForProduct.eof then
		'Move to the selected page in the record set
		rsPhotosForProduct.AbsolutePage = session("CurrentPage")
	end if
	
	set oCmd = Nothing
	
	set GetPhotosForProduct = rsPhotosForProduct
	
End Function

Function GetPhotosForProductCount(ProductID, sSortBy, iPageSize, sSearchCriteria, CurrentRound)

	Dim rsPhotosForProductCount
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsPhotosForProductCount = server.CreateObject("ADODB.Recordset")
		   
	sSQL = "SELECT Count(Photos_For_Voting.id) as productcount " & _
		   "FROM ((Forums INNER JOIN Photos_For_Voting ON Forums.id = Photos_For_Voting.Forum_ID) " & _
		   "INNER JOIN Products ON Photos_For_Voting.Product_ID = Products.id) " & _
		   "INNER JOIN Members ON Photos_For_Voting.Member_ID = Members.id " & _
		   "WHERE (((Photos_For_Voting.Product_ID) = " & ProductID & ") " & _
		   "AND ((Photos_For_Voting.Approved) = True) " & _
		   "AND ((Photos_For_Voting.Round) = " & CurrentRound & ")) "
		   
	if searchcriteria <> "" then
		sSQL = sSQL & "AND (Members.User like '%" & searchcriteria & "%' " & _
			   "OR Photos_For_Voting.Description like '%" & searchcriteria & "%') "
	end if

	rsPhotosForProductCount.Open sSQL, conn

	set oCmd = Nothing
	
	set GetPhotosForProductCount = rsPhotosForProductCount
	
End Function



Sub AddProduct( SKU, ProductType, ProductName, sDescription, Price, Thumbnail, _
				Large_Image, Active, Related_Items, InStock, ForumID, Notes, _
				sImagePath, sThumbNailPath, sBlowupPath, sFrom )

	Dim sSQL
	Dim conn
	
	if ForumID = "" then
		ForumID = 0
	end if
	
	ProductPath = server.MapPath("/") & "/images/"
	if len(sImagePath) = 0 then
		sImagePath = ProductName & "/images/"
		sThumbNailPath = ProductName & "/thumbnails/"
		sBlowupPath = ProductName & "/blowups/"
	end if

	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	sSQL = "INSERT INTO Products([SKU], [Product_Type], [Product_Name], [Description], [Price], " & _
		   "[Thumbnail], [Large_Image], [Active], [Related_Items], [In_Stock], [Forum_ID], " & _
		   "[Notes], [ImagePath], [ThumbNailPath], [BlowupPath] ) " &  _
		   "VALUES( @SKU, @ProductType, @ProductName, @sDescription, @Price, @ThumbNail, " & _
		   "@LargeImage, @sActive, @RelatedItem, @InStock, @ForumID, @sNotes, " & _
		   "@ImagePath, @ThumbNailPath, @BlowupPath )"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@SKU", adVarChar, adParamInput, Len(SKU), SKU)
    	.Parameters.Append .CreateParameter("@ProductType", adVarChar, adParamInput, Len(ProductType), ProductType)
    	.Parameters.Append .CreateParameter("@ProductName", adVarChar, adParamInput, Len(ProductName), ProductName)
		if Len(sDescription) = 0 then
	    	.Parameters.Append .CreateParameter("@sDescription", adVarChar, adParamInput, 1, null)
		else
	    	.Parameters.Append .CreateParameter("@sDescription", adVarChar, adParamInput, Len(sDescription), sDescription)
		end if
    	.Parameters.Append .CreateParameter("@Price", adNumeric, adParamInput, 4, Price)
		if Len(Thumbnail) = 0 then
    		.Parameters.Append .CreateParameter("@ThumbNail", adVarWChar, adParamInput, 1, null)
		else
    		.Parameters.Append .CreateParameter("@ThumbNail", adVarChar, adParamInput, Len(Thumbnail), Thumbnail)
		end if
		if Len(Large_Image) = 0 then
    		.Parameters.Append .CreateParameter("@LargeImage", adVarWChar, adParamInput, 1, null)
		else
    		.Parameters.Append .CreateParameter("@LargeImage", adVarChar, adParamInput, Len(Large_Image), Large_Image)
		end if
    	.Parameters.Append .CreateParameter("@sActive", adBoolean, adParamInput, 4, Active)
		if Len(Related_Items) = 0 then
	    	.Parameters.Append .CreateParameter("@RelatedItem", adVarWChar, adParamInput, 1, null)
		else
	    	.Parameters.Append .CreateParameter("@RelatedItem", adVarChar, adParamInput, Len(Related_Items), Related_Items)
		end if
    	.Parameters.Append .CreateParameter("@InStock", adBoolean, adParamInput, 4, InStock)
    	.Parameters.Append .CreateParameter("@ForumID", adVarChar, adParamInput, Len(ForumID), ForumID)
		if Len(Notes) = 0 then
	    	.Parameters.Append .CreateParameter("@sNotes", adLongVarChar, adParamInput, 1, null)
		else
	    	.Parameters.Append .CreateParameter("@sNotes", adLongVarChar, adParamInput, Len(Notes) , Notes)
		end if
		.Parameters.Append .CreateParameter("@ImagePath", adVarChar, adParamInput, Len(sImagePath), sImagePath)
    	.Parameters.Append .CreateParameter("@ThumbNailPath", adVarChar, adParamInput, Len(sThumbnailPath), sThumbnailPath)
    	.Parameters.Append .CreateParameter("@BlowupPath", adVarChar, adParamInput, Len(sBlowupPath), sBlowupPath)
	end with

	oCmd.execute
	conn.Close()

	if sFrom <> "Import" then
		if sFrom = "" then
			response.Redirect("product.asp")
		else
			response.Redirect("Product.asp")
		end if
	end if
	
End Sub


Sub UpdateProduct( ProductID, SKU, ProductType, ProductName, sDescription, Price, Thumbnail, _
				   Large_Image, Active, Related_Items, InStock, ForumID, Notes, _
				   sImagePath, sThumbNailPath, sBlowupPath, sFrom )

	Dim sSQL
	Dim conn
	Dim ProductPath
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	if (ForumID) = "" then
		ForumID = 0
	end if
	
	ProductPath = server.MapPath("/") & "/images/"
	if len(sImagePath) = 0 then
		sImagePath = ProductName & "/images/"
		sThumbNailPath = ProductName & "/thumbnails/"
		sBlowupPath = ProductName & "/blowups/"
	end if
	
	if IsNull(Large_Image) then
		Large_Image = ""
	end if
	if IsNull(ThumbNail) then
		Thumbnail = ""
	end if
	sSQL = "UPDATE Products " & _
		   "SET [SKU] = @SKU, " & _
		   "[Product_Type] = @ProductType, " & _
		   "[Product_Name] = @ProductName, " & _
		   "[Description] = @sDescription, " & _
		   "[Price] = @Price, " & _
		   "[Thumbnail] = @ThumbNailImage, " & _
		   "[Large_Image] = @LargeImage, " & _
		   "[Active] = @sActive, " & _
		   "[Related_Items] = @RelatedItem, " & _
		   "[In_Stock] = @InStock, " & _
		   "[Forum_ID] = @ForumID, " & _
		   "[Notes] = @sNotes, " & _
		   "[ImagePath] = @ImagePath, " & _
		   "[ThumbNailPath] = @ThumbNailPath, " & _
		   "[BlowupPath] = @BlowupPath " & _
		   "WHERE id = @productID"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@SKU", adVarChar, adParamInput, Len(SKU), SKU)
    	.Parameters.Append .CreateParameter("@ProductType", adVarChar, adParamInput, Len(ProductType), ProductType)
    	.Parameters.Append .CreateParameter("@ProductName", adVarChar, adParamInput, Len(ProductName), ProductName)
		if Len(sDescription) = 0 then
	    	.Parameters.Append .CreateParameter("@sDescription", adLongVarChar, adParamInput, 1, null)
		else
	    	.Parameters.Append .CreateParameter("@sDescription", adLongVarChar, adParamInput, Len(sDescription), sDescription)
		end if
    	.Parameters.Append .CreateParameter("@Price", adNumeric, adParamInput, 4, Price)
		if len(thumbnail) = 0 then
	    	.Parameters.Append .CreateParameter("@ThumbNailImage", adVarChar, adParamInput, 1, null)
		else
	    	.Parameters.Append .CreateParameter("@ThumbNailImage", adVarChar, adParamInput, Len(Thumbnail), Thumbnail)
		end if
		if Len(Trim(Large_Image)) = 0 then
	    	.Parameters.Append .CreateParameter("@LargeImage", adVarChar, adParamInput, 1, null)
		else
	    	.Parameters.Append .CreateParameter("@LargeImage", adVarChar, adParamInput, Len(Large_Image), Large_Image)
		end if
    	.Parameters.Append .CreateParameter("@sActive", adBoolean, adParamInput, 4, Active)
		if Len(Related_Items) = 0 then
		   	.Parameters.Append .CreateParameter("@RelatedItem", adVarChar, adParamInput, 1, null)
		else
		   	.Parameters.Append .CreateParameter("@RelatedItem", adVarChar, adParamInput, Len(Related_Items), Related_Items)
		end if
    	.Parameters.Append .CreateParameter("@InStock", adBoolean, adParamInput, 4, InStock)
    	.Parameters.Append .CreateParameter("@ForumID", adVarChar, adParamInput, Len(ForumID), ForumID)
		if Len(Notes) = 0 then
	    	.Parameters.Append .CreateParameter("@sNotes", adLongVarChar, adParamInput, 1, null)
		else
	    	.Parameters.Append .CreateParameter("@sNotes", adLongVarChar, adParamInput, Len(Notes) , Notes)
		end if
    	.Parameters.Append .CreateParameter("@ImagePath", adVarChar, adParamInput, Len(sImagePath), sImagePath)
    	.Parameters.Append .CreateParameter("@ThumbNailPath", adVarChar, adParamInput, Len(sThumbnailPath), sThumbnailPath)
    	.Parameters.Append .CreateParameter("@BlowupPath", adVarChar, adParamInput, Len(sBlowupPath), sBlowupPath)
    	.Parameters.Append .CreateParameter("@productID", adInteger, adParamInput, 4, ProductID)
		
	end with

	oCmd.execute
	conn.Close()
	
	'Create the myFSO variable
	Dim myFSO
	'Initialize myFSO as a file system object
	SET myFSO = Server.CreateObject("Scripting.FileSystemObject")

	'---------------------------------------------------------------------
	' Create the image path where all images will be saved
	'---------------------------------------------------------------------
	If NOT myFSO.FolderExists(ProductPath & ProductName) Then 
		myFSO.CreateFolder(ProductPath & ProductName)
	else
	
	End If 

	'---------------------------------------------------------------------
	' Create the thumbnail image path
	'---------------------------------------------------------------------
	If NOT myFSO.FolderExists(ProductPath & sImagePath) Then 
		myFSO.CreateFolder(ProductPath & sImagePath)
	End If 

	'---------------------------------------------------------------------
	' Create the thumbnail image path
	'---------------------------------------------------------------------
	If NOT myFSO.FolderExists(ProductPath & sThumbNailPath) Then 
		myFSO.CreateFolder(ProductPath & sThumbNailPath)
	End If 

	'---------------------------------------------------------------------
	' Create the blowup image path
	'---------------------------------------------------------------------
	If NOT myFSO.FolderExists(ProductPath & sBlowupPath) Then 
		myFSO.CreateFolder(ProductPath & sBlowupPath)
	End If 

	SET myFSO = NOTHING 

	
	if sFrom = "product" then
		response.Redirect("product.asp")
	elseif sFrom = "" then
		response.Redirect("product.asp")
	else
		response.Redirect("product.asp")
	end if
	
End Sub

Sub DeleteProduct( pid )

	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	sSQL = "DELETE FROM Products " & _
		   "WHERE id = " & pid

	conn.execute sSQL
	conn.Close()
	
	response.Redirect("product.asp")
	
End Sub

%>