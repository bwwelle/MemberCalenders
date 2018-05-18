<%

Function GetWebPages()

	Dim rsWebPages
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsWebPages = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT WebPage, WebPageContent, WebPageDisplayName  " & _
		   "FROM WebPages " & _
		   "ORDER BY WebPageDisplayName"

	rsWebPages.Open sSQL, conn
	
	set GetWebPages = rsWebPages
	
End Function

Function GetWebPage(WebPage)

	Dim rsWebPage
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsWebPage = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT WebPage, WebPageContent, WebPageDisplayName  " & _
		   "FROM WebPages " & _
		   "WHERE WebPage = @WebPage"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	
	with oCmd
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@WebPage", adVarChar, adParamInput, Len(WebPage), WebPage)
	end with
	
	rsWebPage.Open oCmd
	
	set GetWebPage = rsWebPage
	
End Function

Sub UpdateWebPage( WebPageName, WebPageContent)

	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	sSQL = "UPDATE WebPages " & _
		   "SET WebPageContent = @WebPageContent " & _
		   "WHERE WebPage = @WebPage"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	with oCmd                                                      
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

		if Len(WebPageContent) = 0 then
	    	.Parameters.Append .CreateParameter("@WebPageContent", adLongVarChar, adParamInput, 1, null)
		else
	    	.Parameters.Append .CreateParameter("@WebPageContent", adLongVarChar, adParamInput, Len(WebPageContent) , WebPageContent)
		end if
    	.Parameters.Append .CreateParameter("@WebPage", adVarChar, adParamInput, Len(WebPageName), WebPageName)
	end with

	oCmd.execute
	
	response.Redirect("WebPages.asp")
	
End Sub


'		if Len(sNotes) = 0 then
'	    	.Parameters.Append .CreateParameter("@Notes", adLongVarChar, adParamInput, 1, null)
'		else
'	    	.Parameters.Append .CreateParameter("@Notes", adLongVarChar, adParamInput, Len(sNotes), sNotes)
'		end if

%>