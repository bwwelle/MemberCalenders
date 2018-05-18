<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!-- #include file ="dbconnections.asp" -->
<!-- #include file ="clforum.asp" -->
<!-- #include file ="clMemberAccounts.asp" -->
<!-- #include file ="clproduct.asp" -->
<!-- #include file ="clOrder.asp" -->
<!-- #include file ="clLookup.asp" -->

<!-- #include file = "adminsession.asp" -->

<% 
Dim ExportType
Dim objFile
Dim sPath
Dim objFSO
Dim rsExport
Dim sColumnHeadings
Dim x
Dim y
Dim strLine

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")


ExportType = Request.QueryString("exporttype")

if ExportType = "forums" then
	sColumnHeadings = "id,Name,Code,Password,Address,City,State,Zip,Country,Phone1,Phone2," & _
	"Fax,Email,Url,Contact First Name,Contact Last Name,Active,Approved,Date Approved,Notes"

	sPath = Server.MapPath("/") & "/admin/Exports/Forums/Forums_" & Replace(FormatDateTime(now(),2),"/","_") & ".csv"

	set rsExport = GetForums()

elseif ExportType = "members" then
	sColumnHeadings = "id,User,Password,Forum UserName,First Name, Last Name," & _ 
					  "Address,City,State,Zip,Country,Phone1,Phone2," & _
				 	  "Fax,Email,Bio,IP_Address, Url,Join Date,Active,,Notes"


	set rsExport = GetMembers("","asc")

	sPath = Server.MapPath("/") & "/admin/Exports/Members/members_" & Replace(FormatDateTime(now(),2),"/","_") & ".csv"
elseif ExportType = "orders" then
sColumnHeadings = "Order Number,Customer Name,Order Date,Order Total,Tax,Shipping," & _
				      "First Name,Last Name,Email,Address1,Address2,City,State,Zip,Country," & _
					  "Day Phone,Night Phone,Ship First Name,Ship Last Name,Ship Address1," & _
					  "Ship Address2,Ship City,Ship State,Ship Zip,Ship Country,Ship Day Phone," & _
					  "Ship Night Phone,Password,Order_Status, SKU, Quantity,Reference Code"


	set rsExport = GetOrders()
	
	sPath = Server.MapPath("/") & "/admin/Exports/Orders/Orders_" & Replace(FormatDateTime(now(),2),"/","_") & ".csv"
elseif ExportType = "products" then

	sColumnHeadings = "id,SKU,Product Type,Product Name,Description,Price,ThumbNail,Large Image," & _
	 				  "Active,RelatedItems,In Stock,Forum ID,Notes,Date Created,Image Path, " & _
					  "ThumbNailPath,BlowupPath"

	set rsExport = GetProducts()

	sPath = Server.MapPath("/") & "/admin/Exports/Products/Products_" & Replace(FormatDateTime(now(),2),"/","_") & ".csv"
end if

Set objFile = objFSO.CreateTextFile(sPath)




objFile.WriteLine sColumnHeadings
Do While Not rsExport.EOF
	strLine = ""
	For j = 0 to rsExport.Fields.Count-1
		If Not IsNull(rsExport(j)) Then 
			strLine = strLine & Chr(34) & Replace(rsExport(j), Chr(34), Chr(34) & Chr(34)) & Chr(34)
		end if
		If j < rsExport.Fields.Count-1 Then 
			strLine = strLine & ","
		end if
	Next
	objFile.WriteLine strLine
	rsExport.MoveNext
Loop

response.Redirect("ImportExportData.asp")
%> 
