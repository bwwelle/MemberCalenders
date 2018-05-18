<%

Function GetProductTypes()

	Dim rsProductTypes
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsProductTypes = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT id, product_type, description, [notes] " & _
		   "FROM Product_Types " & _
		   "ORDER BY Product_Type"

	rsProductTypes.Open sSQL, conn
	
	set GetProductTypes = rsProductTypes
	
	
	
	
End Function

Function GetProductType(ptid)

	Dim rsProductType
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsProductType = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT id, product_type, description, notes " & _
		   "FROM Product_Types " & _
		   "WHERE id = @ptid"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@ptid", adInteger, adParamInput, 4, ptid)
	end with
	
	rsProductType.Open oCmd
	
	set GetProductType = rsProductType
	
End Function

Sub AddProductType( productType, sDescription, sNotes )

	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	sSQL = "INSERT INTO Product_Types([product_type], [Description], [Notes]) " & _
		   "VALUES( @productType, @sDescription, @Notes)"


	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@productType", adVarChar, adParamInput, Len(productType), productType)
		if Len(sDescription) = 0 then
	    	.Parameters.Append .CreateParameter("@sDescription", adVarChar, adParamInput, 1, null)
		else
	    	.Parameters.Append .CreateParameter("@sDescription", adVarChar, adParamInput, Len(sDescription), sDescription)
		end if
		if Len(sNotes) = 0 then
	    	.Parameters.Append .CreateParameter("@Notes", adLongVarChar, adParamInput, 1, null)
		else
	    	.Parameters.Append .CreateParameter("@Notes", adLongVarChar, adParamInput, Len(sNotes) , sNotes)
		end if
	end with

	oCmd.execute
	conn.Close()

	response.Redirect("product_type.asp")
	
End Sub


Sub UpdateProductType( productTypeID, productType, sDescription, sNotes )

	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	sSQL = "UPDATE Product_Types " & _
		   "SET product_type = @ProductType, " & _
		   "[Description] = @Description, " & _
		   "[Notes] = @Notes " & _
		   "WHERE id = @ProductTypeID"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@ProductType", adVarChar, adParamInput, Len(productType), productType)
		if Len(sDescription) = 0 then
	    	.Parameters.Append .CreateParameter("@Description", adVarChar, adParamInput, 1, null)
		else
	    	.Parameters.Append .CreateParameter("@Description", adVarChar, adParamInput, Len(sDescription), sDescription)
		end if
		if Len(sNotes) = 0 then
	    	.Parameters.Append .CreateParameter("@Notes", adLongVarChar, adParamInput, 1, null)
		else
	    	.Parameters.Append .CreateParameter("@Notes", adLongVarChar, adParamInput, Len(sNotes) , sNotes)
		end if
    	.Parameters.Append .CreateParameter("@ProductTypeID", adInteger, adParamInput, 4, productTypeID)
	end with

	oCmd.execute
	conn.Close()
	
	response.Redirect("product_type.asp")
	
End Sub

Sub DeleteProductType( productTypeID )

	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	sSQL = "DELETE FROM Product_Types " & _
		   "WHERE id = @productTypeID"


	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@productTypeID", adInteger, adParamInput, 4, productTypeID)
	end with
	
	oCmd.execute
		
	conn.Close()
	
	response.Redirect("product_type.asp")
	
End Sub


%>