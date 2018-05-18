<!-- #include file ="dbconnections.asp" -->
<!-- #include file ="clPhotosForVoting.asp" -->
<!-- #include file ="clVotingSchedule.asp" -->
<!-- #include file ="clproduct.asp" -->
<!-- #include virtual ="/constants.asp" -->
<%


Dim rsRoundsToCalculate
Dim ProcessingRound
Dim CurrentRound
Dim NextRound
Dim RoundWinnersNumber
Dim ProductID
Dim ProcessRound
Dim CurrentMessage

set rsRoundsToCalculate = GetRoundsToCalculate()

AddMessageForWinnersScheduler "Determine Winners Started" 		


if not rsRoundsToCalculate.eof then
	While not rsRoundsToCalculate.eof 

		CurrentRound = 0
		ProcessingRound = 0
		NextRound = 0
		RoundWinnersNumber = ""
		ProductID = 0
		ProcessRound = 0
		CurrentMessage = ""
	
		if (rsRoundsToCalculate("Round_1_End") < CurrentTime and (isnull(rsRoundsToCalculate("Round_1_Completed")) or trim(rsRoundsToCalculate("Round_1_Completed")) = "")) then
			CurrentRound = 1
			ProcessingRound = 1
			NextRound = 2
			RoundWinnersNumber = rsRoundsToCalculate("Round_1_Winners")
			ProductID = GetProductBySku(rsRoundsToCalculate("SKU"))
			ProcessRound = 1
		End if
		if (rsRoundsToCalculate("Round_2_End") < CurrentTime and (isnull(rsRoundsToCalculate("Round_2_Completed")) or rsRoundsToCalculate("Round_2_Completed") = "")) then
			CurrentRound = 2
			ProcessingRound = 2
			NextRound = 3
			RoundWinnersNumber = rsRoundsToCalculate("Round_2_Winners")
			ProductID = GetProductBySku(rsRoundsToCalculate("SKU"))
			ProcessRound = 1
		End if
		if (rsRoundsToCalculate("Round_3_End") < CurrentTime and (isnull(rsRoundsToCalculate("Round_3_Completed")) or rsRoundsToCalculate("Round_3_Completed") = "")) then
			CurrentRound = 3
			ProcessingRound = 3
			NextRound = 4
			RoundWinnersNumber = rsRoundsToCalculate("Round_3_Winners")
			ProductID = GetProductBySku(rsRoundsToCalculate("SKU"))
			ProcessRound = 1
		End if

		if productID > 0 then
			if ProcessRound = 1 then
				UpdateWinnersForRound RoundWinnersNumber, ProductID, ProcessingRound
				UpdateVoteScheduleCompletedDate rsRoundsToCalculate("sku"), ProcessingRound
			end if
	
			CurrentMessage = "Processing ProductID = " & ProductID & " for round " & _
							 ProcessingRound & ".  Winning photo ids are " & RoundWinnersNumber & "."
			
			AddMessageForWinnersScheduler CurrentMessage 

		end if
		rsRoundsToCalculate.MoveNext
	Wend
	AddMessageForWinnersScheduler "The Winners for all expired rounds have been processed"
else
	AddMessageForWinnersScheduler "Nothing was ready for processing" 		
end if


AddMessageForWinnersScheduler "Determine Winners Ended" 		

%>

