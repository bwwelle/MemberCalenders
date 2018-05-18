<%

Function GetStates()

	Dim rsStates
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsStates = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT State_Name, State_Abbreviation  " & _
		   "FROM States " & _
		   "ORDER BY State_Name"

	rsStates.Open sSQL, conn
	
	set GetStates = rsStates
	
End Function

Function GetCountries()

	Dim rsCountries
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsCountries = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT country_name,  sort_order " & _
		   "FROM country " & _
		   "ORDER BY sort_order, country_name"
	rsCountries.Open sSQL, conn
	
	set GetCountries = rsCountries
	
End Function

%>