<%

Function GetRoundsToCalculate()

	Dim rsGetRoundsToCalculate
	Dim sSQL
	Dim conn

	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsGetRoundsToCalculate = server.CreateObject("ADODB.Recordset")
	
	sSQL = "SELECT ID, SKU, Round_1_Start, Round_1_End, Round_2_Start, Round_2_End, " & _
		   "Round_3_Start, Round_3_End, Round_1_Number_Of_Votes, Round_2_Number_Of_Votes, " & _
		   "Round_3_Number_Of_Votes, Round_1_Winners, Round_2_Winners, Round_3_Winners, " & _
		   "Cut_Off_Date, Round_1_Completed, Round_2_Completed, Round_3_Completed " & _
		   "FROM Vote_Schedule " & _
		   "WHERE ((Round_1_End < @CurrentDate1) OR (Round_2_End < @CurrentDate2) OR (Round_3_End < @CurrentDate3)) " & _
		   "ORDER BY Round_1_End, Round_2_End, Round_3_End"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	with oCmd                                                      
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@CurrentDate1", adDBTimeStamp, adParamInput, 16, CurrentTime)
    	.Parameters.Append .CreateParameter("@CurrentDate2", adDBTimeStamp, adParamInput, 16, CurrentTime)
    	.Parameters.Append .CreateParameter("@CurrentDate3", adDBTimeStamp, adParamInput, 16, CurrentTime)
	end with

	rsGetRoundsToCalculate.Open oCmd
	
	set GetRoundsToCalculate = rsGetRoundsToCalculate


End Function

Function GetVotingSchedules()

	Dim rsVotingSchedules
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsVotingSchedules = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT ID, SKU, Round_1_Start, Round_1_End, Round_2_Start, Round_2_End, " & _
		   "Round_3_Start, Round_3_End, Round_1_Number_Of_Votes, Round_2_Number_Of_Votes, " & _
		   "Round_3_Number_Of_Votes, Round_1_Winners, Round_2_Winners, Round_3_Winners, " & _
		   "Cut_Off_Date " & _
		   "FROM Vote_Schedule " & _
		   "ORDER BY ID"

	rsVotingSchedules.Open sSQL, conn
	
	set GetVotingSchedules = rsVotingSchedules
	
End Function

Function GetVotingSchedule(vsID)

	Dim rsVotingSchedule
	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	set rsVotingSchedule = server.CreateObject("ADODB.Recordset")
	
	
	sSQL = "SELECT ID, SKU, Round_1_Start, Round_1_End, Round_2_Start, Round_2_End, " & _
		   "Round_3_Start, Round_3_End, Round_1_Number_Of_Votes, Round_2_Number_Of_Votes, " & _
		   "Round_3_Number_Of_Votes, Round_1_Winners, Round_2_Winners, Round_3_Winners, " & _
		   "Cut_Off_Date " & _
		   "FROM Vote_Schedule " & _
		   "WHERE id = " & vsID

	rsVotingSchedule.Open sSQL, conn
	
	set GetVotingSchedule = rsVotingSchedule
	
End Function

Sub AddMessageForWinnersScheduler( Message )

	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	
	sSQL = "INSERT INTO DetermineWinnersMessage([Message],[CreateDate]) " & _
		   "VALUES( @Message, @CurrentDateTime )"

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	with oCmd                                                      
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@Message", adVarChar, adParamInput, Len(Message), Message)
    	.Parameters.Append .CreateParameter("@CurrentDateTime", adDBTimeStamp, adParamInput, 16, CurrentTime)
	end with

	oCmd.execute

	conn.Close()

End Sub

Sub AddVotingSchedule( SKU, Round1Start, Round1End, Round2Start, Round2End, _
					   Round3Start, Round3End, Round1NumberOfVotes, Round2NumberOfVotes, _
					   Round3NumberOfVotes, Round1Winners, Round2Winners, Round3Winners, _
					   CutOffDate)

	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	
	sSQL = "INSERT INTO Vote_Schedule([SKU], [Round_1_Start], [Round_1_End], [Round_2_Start], " & _
		   "[Round_2_End], [Round_3_Start], [Round_3_End], [Round_1_Number_Of_Votes], " & _
		   "[Round_2_Number_Of_Votes], [Round_3_Number_Of_Votes], [Round_1_Winners], " & _ 
		   "[Round_2_Winners], [Round_3_Winners], [Cut_Off_Date] )" & _
		   "VALUES( '" & SKU & "', #" & Round1Start & "#, #" & Round1End & "#, #" & _
		   Round2Start & "#, #" & Round2End & "#, #" & Round3Start & "#, #" & _
		   Round3End & "#, " & Round1NumberOfVotes & ", " & Round1NumberOfVotes & ", " & _
		   Round3NumberOfVotes & ", " & Round1Winners & ", " & Round2Winners & ", " & _
		   Round3Winners & ", #" & CutOffDate & "#)"

	conn.execute sSQL
	conn.Close()

	response.Redirect("VotingSchedule.asp")
	
End Sub

Sub UpdateVoteScheduleCompletedDate(Sku, CurrentRound)

	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	if CurrentRound = 1 then
		sSQL = "UPDATE Vote_Schedule " & _
			   "SET Round_1_Completed = @CurrentTime " & _
			   "WHERE Sku = @Sku"
	elseif CurrentRound = 2 then
		sSQL = "UPDATE Vote_Schedule " & _
			   "SET Round_2_Completed = @CurrentTime " & _
			   "WHERE Sku = @Sku"
	else
		sSQL = "UPDATE Vote_Schedule " & _
			   "SET Round_3_Completed = @CurrentTime " & _
			   "WHERE Sku = @Sku"
	end if
	

	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	with oCmd                                                      
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@CurrentTime", adDBTimeStamp, adParamInput, 16, CurrentTime)
    	.Parameters.Append .CreateParameter("@Sku", adVarChar, adParamInput, Len(sku), sku)
	end with

	oCmd.execute

	conn.Close()

End Sub


Sub UpdateVotingSchedule( VotingScheduleID, SKU, Round1Start, Round1End, Round2Start, Round2End, _
						  Round3Start, Round3End, Round1NumberOfVotes, Round2NumberOfVotes, _
						  Round3NumberOfVotes, Round1Winners, Round2Winners, Round3Winners, _
						  CutOffDate)

	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	sSQL = "UPDATE Vote_Schedule " & _
		   "SET [SKU] = @SKU, " & _
		   "[Round_1_Start] = @Round1Start, " & _
		   "[Round_1_End] = @Round1End, " & _
		   "[Round_2_Start] = @Round2Start, " & _
		   "[Round_2_End] = @Round2End, " & _
		   "[Round_3_Start] = @Round3Start, " & _
		   "[Round_3_End] = @Round3End, " & _
		   "[Round_1_Number_Of_Votes] = @Round1NumberOfVotes, " & _
		   "[Round_2_Number_Of_Votes] = @Round2NumberOfVotes, " & _
		   "[Round_3_Number_Of_Votes] = @Round3NumberOfVotes, " & _
		   "[Round_1_Winners] = @Round1Winners, " & _
		   "[Round_2_Winners] = @Round2Winners, " & _
		   "[Round_3_Winners] = @Round3Winners, " & _
		   "[Cut_Off_Date] = @CutOffDate " & _
		   "WHERE id = @VotingScheduleID"



	Dim oCmd
	set oCmd = Server.CreateObject("ADODB.Command")
	with oCmd                                                      
		.ActiveConnection = conn
    	.CommandText = sSQL
    	.CommandType = adCmdText

    	.Parameters.Append .CreateParameter("@Sku", adVarChar, adParamInput, Len(sku), sku)
    	.Parameters.Append .CreateParameter("@Round1Start", adDBTimeStamp, adParamInput, 16, Round1Start)
    	.Parameters.Append .CreateParameter("@Round1End", adDBTimeStamp, adParamInput, 16, Round1End)
    	.Parameters.Append .CreateParameter("@Round2Start", adDBTimeStamp, adParamInput, 16, Round2Start)
    	.Parameters.Append .CreateParameter("@Round2End", adDBTimeStamp, adParamInput, 16, Round2End)
    	.Parameters.Append .CreateParameter("@Round3Start", adDBTimeStamp, adParamInput, 16, Round3Start)
    	.Parameters.Append .CreateParameter("@Round3End", adDBTimeStamp, adParamInput, 16, Round3End)
    	.Parameters.Append .CreateParameter("@Round1NumberOfVotes", adInteger, adParamInput, 4, Round1NumberOfVotes)
    	.Parameters.Append .CreateParameter("@Round2NumberOfVotes", adInteger, adParamInput, 4, Round2NumberOfVotes)
    	.Parameters.Append .CreateParameter("@Round3NumberOfVotes", adInteger, adParamInput, 4, Round3NumberOfVotes)
    	.Parameters.Append .CreateParameter("@Round1Winners", adInteger, adParamInput, 4, Round1Winners)
    	.Parameters.Append .CreateParameter("@Round2Winners", adInteger, adParamInput, 4, Round2Winners)
    	.Parameters.Append .CreateParameter("@Round3Winners", adInteger, adParamInput, 4, Round3Winners)
    	.Parameters.Append .CreateParameter("@CutOffDate", adDBTimeStamp, adParamInput, 16, CutOffDate)
    	.Parameters.Append .CreateParameter("@VotingScheduleID", adInteger, adParamInput, 4, VotingScheduleID)
	end with

	oCmd.execute

	conn.Close()
	
	response.Redirect("VotingSchedule.asp")
	
End Sub

Sub DeleteVotingSchedule( VotingScheduleID )

	Dim sSQL
	Dim conn
	
	set conn = server.CreateObject("ADODB.Connection")
	conn.Open(dicConnections("customcal"))
	
	sSQL = "DELETE FROM Vote_Schedule " & _
		   "WHERE id = " & VotingScheduleID

	conn.execute sSQL
	conn.Close()
	
	response.Redirect("VotingSchedule.asp")
	
End Sub

%>