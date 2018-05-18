<%
Dim dbcstrPath
Dim dbcintLoc
Dim dicConnections

Const adLockReadOnly = 1 'Open a RecordSet in read-only mode

'---- CursorType Values ---- 
Const adOpenForwardOnly = 0 
Const adOpenKeyset = 1 
Const adOpenDynamic = 2 
Const adOpenStatic = 3 

'---- CursorLocationEnum Values ---- 
Const adUseServer = 2 
Const adUseClient = 3 

'---- DataTypeEnum Values ---- 
Const adEmpty = 0 
Const adTinyInt = 16 
Const adSmallInt = 2 
Const adInteger = 3 
Const adBigInt = 20 
Const adUnsignedTinyInt = 17 
Const adUnsignedSmallInt = 18 
Const adUnsignedInt = 19 
Const adUnsignedBigInt = 21 
Const adSingle = 4 
Const adDouble = 5 
Const adCurrency = 6 
Const adDecimal = 14 
Const adNumeric = 131 
Const adBoolean = 11 
Const adError = 10 
Const adUserDefined = 132 
Const adVariant = 12 
Const adIDispatch = 9 
Const adIUnknown = 13 
Const adGUID = 72 
Const adDate = 7 
Const adDBDate = 133 
Const adDBTime = 134 
Const adDBTimeStamp = 135 
Const adBSTR = 8 
Const adChar = 129 
Const adVarChar = 200 
Const adLongVarChar = 201 
Const adWChar = 130 
Const adVarWChar = 202 
Const adLongVarWChar = 203 
Const adBinary = 128 
Const adVarBinary = 204 
Const adLongVarBinary = 205 

'---- CommandTypeEnum Values ---- 
Const adCmdUnknown = &H0008 
Const adCmdText = 1
Const adCmdTable = &H0002 
Const adCmdStoredProc = &H0004 

'---- ParameterDirectionEnum Values ----
Const adParamUnknown = &H0000
Const adParamInput = &H0001
Const adParamOutput = &H0002
Const adParamInputOutput = &H0003
Const adParamReturnValue = &H0004



dbcstrPath = Request.ServerVariables("SCRIPT_NAME")
dbcintLoc = instrrev(dbcstrPath, "/")
if dbcintLoc > 2 then
	dbcstrPath = left(dbcstrPath, dbcintLoc - 1)
else
	dbcstrPath = ""
end if
set dicConnections = server.CreateObject("Scripting.Dictionary")
dicConnections.Add "customcal", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & server.mapPath("/" & "database/Customcal.mdb")


sub CloseConnection(objConnection)
	
	if isobject(objConnection) then
		if not objConnection is nothing then
			if objConnection.State = adStateOpen then
				'Response.Write "<!-- Closing Connection -->" & vbcrlf
				objConnection.Close
			end if
			'Response.Write "<!-- Setting Connection to Nothing -->" & vbcrlf
			set objConnection = nothing
		end if
	end if
end sub

%>