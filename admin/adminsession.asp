<%
	if session("adminid") = "" then
	
		response.Redirect("index.asp?errorno=10001")
		
	end if
%>