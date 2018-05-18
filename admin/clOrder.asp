<%

Function GetOrders()

	Dim rsOrders
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsOrders = server.CreateObject("ADODB.Recordset")	
	
	sSQL = "SELECT Orders.Order_Number, Orders.Customer_Name, Orders.Order_Date, Orders.Order_Total, Orders.Tax, Orders.Shipping, Orders.First_Name, Orders.Last_Name, Orders.Email, Orders.Address1, Orders.Address2, Orders.City, Orders.State, Orders.Zip, Orders.Country, Orders.Day_Phone, Orders.Night_Phone, Orders.Ship_First_Name, Orders.Ship_Last_Name, Orders.Ship_Address1, Orders.Ship_Address2, Orders.Ship_City, Orders.Ship_State, Orders.Ship_Zip, Orders.Ship_Country, Orders.Ship_Day_Phone, Orders.Ship_Night_Phone, Orders.Password, Orders.Order_Status, Products.SKU, Orders_Detail.Quantity, orders.referenceCode " & _
"FROM Orders INNER JOIN (Orders_Detail INNER JOIN Products ON Orders_Detail.item_id=Products.id) ON Orders.Order_Number=Orders_Detail.Order_Number where orders.completed=yes ORDER BY Orders.Order_Number DESC"

	rsOrders.Open sSQL, conn
	
	set GetOrders = rsOrders
	
End Function

Sub UpdateOrder(OrderNumber, OrderStatus, TrackingNumber)

	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	sSQL = "UPDATE Orders " & _
		   "SET Order_Status = @OrderStatus, " & _
		   "TrackingNumber = @TrackingNumber " & _
		   "WHERE Order_Number = @OrderNumber"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@OrderStatus",adVarChar,adParamInput,Len(OrderStatus),OrderStatus)
    	.Parameters.Append .CreateParameter("@TrackingNumber",adVarChar,adParamInput,Len(TrackingNumber),TrackingNumber)
    	.Parameters.Append .CreateParameter("@OrderNumber",adInteger,adParamInput,4,OrderNumber)

	end with

	oCmd.execute
	conn.Close()


End Sub
%>