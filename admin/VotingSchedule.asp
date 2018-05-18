<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!-- #include file ="dbconnections.asp" -->
<!-- #include file ="clVotingSchedule.asp" -->
<!-- #include file ="clproduct.asp" -->
<!-- #include Virtual ="/constants.asp" -->

<!-- #include file = "adminsession.asp" -->

<%


Dim rsVotingSchedules
Dim rsProducts
Dim sAction
Dim editID
Dim editSku
Dim editRound1Start
Dim editRound1End
Dim editRound1StartTime
Dim editRound1EndTime
Dim editRound2Start
Dim editRound2End
Dim editRound2StartTime
Dim editRound2EndTime
Dim editRound3Start
Dim editRound3End
Dim editRound3StartTime
Dim editRound3EndTime
Dim editRound1NumberOfVotes
Dim editRound2NumberOfVotes
Dim editRound3NumberOfVotes
Dim editRound1Winners
Dim editRound2Winners
Dim editRound3Winners
Dim editCutOffDate

sAction = "view"
if Request.QueryString("action") <> "" then
	sAction = Request.QueryString("action")
end if

if sAction = "view" then
	set rsVotingSchedules = GetVotingSchedules()
	HeaderText = "Voting Schedules"
elseif sAction = "add" then
	HeaderText = "Add Voting Schedules"
	set rsProducts = GetProductsAvailableForVotingSchedule("")
elseif sAction = "edit" then
	HeaderText = "Update Voting Schedules"
	editID = Request.QueryString("id")
	set rsVotingSchedule = GetVotingSchedule(editID)
	if not rsVotingSchedule.eof then
		set rsProducts = GetProductsAvailableForVotingSchedule(rsVotingSchedule("SKU"))
		editSku = rsVotingSchedule("SKU")
		editRound1Start = FormatDateTime(rsVotingSchedule("Round_1_Start"),2)
		editRound1End = FormatDateTime(rsVotingSchedule("Round_1_End"),2)
		editRound1StartTime = Trim(FormatDateTime(rsVotingSchedule("Round_1_Start"),3))
		editRound1EndTime = FormatDateTime(rsVotingSchedule("Round_1_End"),3)
		editRound2Start = FormatDateTime(rsVotingSchedule("Round_2_Start"),2)
		editRound2End = FormatDateTime(rsVotingSchedule("Round_2_End"),2)
		editRound2StartTime = FormatDateTime(rsVotingSchedule("Round_2_Start"),3)
		editRound2EndTime = FormatDateTime(rsVotingSchedule("Round_2_End"),3)
		editRound3Start = FormatDateTime(rsVotingSchedule("Round_3_Start"),2)
		editRound3End = FormatDateTIme(rsVotingSchedule("Round_3_End"),2)
		editRound3StartTime = FormatDateTime(rsVotingSchedule("Round_3_Start"),3)
		editRound3EndTime = FormatDateTime(rsVotingSchedule("Round_3_End"),3)
		editRound1NumberOfVotes = rsVotingSchedule("Round_1_Number_Of_Votes")
		editRound2NumberOfVotes = rsVotingSchedule("Round_2_Number_Of_Votes")
		editRound3NumberOfVotes = rsVotingSchedule("Round_3_Number_Of_Votes")
		editRound1Winners = rsVotingSchedule("Round_1_Winners")
		editRound2Winners = rsVotingSchedule("Round_2_Winners")
		editRound3Winners = rsVotingSchedule("Round_3_Winners")
		editCutOffDate = FormatDateTime(rsVotingSchedule("Cut_Off_Date"),2)
		editCutoffStartTime = FormatDateTime(rsVotingSchedule("Cut_Off_Date"),3)
	end if
end if


%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Voting Schedules</title>
<LINK href="../css/styles.css" rel="stylesheet" type="text/css">
<link href="../css/style.css" rel="stylesheet" type="text/css" />
<LINK href="css/calendar.css" rel="stylesheet" type="text/css">
<script language="javascript" src="js/calendar_us.js"></script>
<script language="javascript">
<!--
function validate() {

	var w = document.frmUpdate.ddlSKU.selectedIndex;
	if (w == 0) {
		alert("Please select a SKU.");
		return false;
	}
	
	if (document.frmUpdate.txtRound1StartDate.value == "") {
		alert("Please enter in the Round 1 Start Date.");
		return false;
	}
	if (document.frmUpdate.txtRound1EndDate.value == "") {
		alert("Please enter in the Round 1 End Date.");
		return false;
	}
	if (document.frmUpdate.txtRound2StartDate.value == "") {
		alert("Please enter in the Round 2 Start Date.");
		return false;
	}
	if (document.frmUpdate.txtRound2EndDate.value == "") {
		alert("Please enter in the Round 2 End Date.");
		return false;
	}
	if (document.frmUpdate.txtRound3StartDate.value == "") {
		alert("Please enter in the Round 3 Start Date.");
		return false;
	}
	if (document.frmUpdate.txtRound3EndDate.value == "") {
		alert("Please enter in the Round 3 End Date.");
		return false;
	}
	
	if (document.frmUpdate.txtRound1NumberOfVotes.value == "") {
		alert("Please enter in the Round 1 Number of Votes.");
		return false;
	}
	if (document.frmUpdate.txtRound2NumberOfVotes.value == "") {
		alert("Please enter in the Round 2 Number of Votes.");
		return false;
	}
	if (document.frmUpdate.txtRound3NumberOfVotes.value == "") {
		alert("Please enter in the Round 3 Number of Votes.");
		return false;
	}


	if (document.frmUpdate.txtRound1Winners.value == "") {
		alert("Please enter in the Round 1 Winners.");
		return false;
	}
	if (document.frmUpdate.txtRound2Winners.value == "") {
		alert("Please enter in the Round 2 Winners.");
		return false;
	}
	if (document.frmUpdate.txtRound3Winners.value == "") {
		alert("Please enter in the Round 3 Winners.");
		return false;
	}

	if (document.frmUpdate.txtCutOffDate.value == "") {
		alert("Please enter in the cut off date.");
		return false;
	}
}

function confirmDelete(id) {

	var answer = confirm('Do you want to delete this item?');
	
	if (answer) {
		document.frmUpdate.hdnAction.value = "delete";
		document.frmUpdate.hdnID.value = id;
		
		document.frmUpdate.submit();
	}
}
//-->
</script>
</head>

<body><br />

<table width="999" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
 <!--header starts--> 
  <!--#include Virtual="header.asp"-->  
 <!--header ends-->
  <tr>
    <td colspan="2" align="left" valign="top">
	
	<table width="100%" style="margin-bottom:5px;"> 
	<tr>
	<td width="136" class="leftnav" valign="top">
	
	<!--left nav starts-->
	<!--#include file="sidenav.asp"-->  
	<!--Left Nav ends--> 
	</td>
  	<td width="12"  align="left" valign="top">&nbsp;</td>
	<td width="848"  align="left" valign="top" bgcolor="#ffffff">
    <!---------------------------Main Body -------------------------------------->
<form action="save.asp" method="post" name="frmUpdate" id="frmUpdate" onsubmit="return validate();">
<input type="hidden" name="hdnSection" id="hdnSection" value="votingschedule" />
<input type="hidden" name="hdnAction" id="hdnAction" value="<%=sAction%>" />
<input type="hidden" name="hdnID" id="hdnID" value="<%=editID%>" />

<% if sAction = "view" then %>
<table width="100%" cellpadding="2" cellspacing="0" border="0" bgcolor="#ffffff">
	<tr>
		<td colspan="6">
			<font style="font-size: large;"><%=HeaderText%></font><br /><br />
			<a href="VotingSchedule.asp?action=add" style="background-color: #ffffff;">Add Voting Schedule</a><br /><br />
		</td>
	</tr>
	<tr style="background-color: gray;">
		<td style="width: 15%;" class="whitetext">SKU</td>
		<td style="width: 20%;" class="whitetext">Product Name</td>
		<td style="width: 30%;" class="whitetext">Cut Off Date</td>
		<td style="width: 5%;"><br /></td>
		<td style="width: 5%;"><br /></td>
		<td style="width: 25%;"><br /></td>
	</tr>
	<% while not rsVotingSchedules.eof 
			set rsProductDetailBySku = GetProductDetailBySku(rsVotingSchedules("SKU"))
			if rsProductDetailBySku.eof then
				productName = ""
			else
				productName = rsProductDetailBySku("product_name")
			end if
			
			
	%>
	<tr>
		<td valign="top"><%=rsVotingSchedules("SKU")%></td>
		<td valign="top"><%=productName%></td>
		<td valign="top"><%=rsVotingSchedules("Cut_Off_Date")%></td>
		<td valign="top"><a href="VotingSchedule.asp?action=edit&id=<%=rsVotingSchedules("id")%>">edit</a></td>
		<td valign="top"><a href="javascript:confirmDelete(<%=rsVotingSchedules("id")%>)">delete</a></td>
		<td><br /></td>
		<% rsVotingSchedules.MoveNext %>
	</tr>
	<% wend %>
</table>
<% end if %>

<% if sAction = "edit" or sAction = "add" then %>
<table width="100%" cellpadding="2" cellspacing="0" border="0" bgcolor="#ffffff">
	<tr>
		<td colspan="6" align="center"><font style="font-size: large;"><%=HeaderText%></font></td>
	</tr>
	<tr>
		<td colspan="6"><br /></td>
	</tr>
	<tr>
		<td style="width: 20%;" align="right" class="requiredfields">SKU:</td>
		<td colspan="3">
				<select name="ddlSKU" id="ddlSKU" style="background-color: #ffffff;">
				<option value="">-- Select Product/SKU --</option>
				<% while not rsProducts.eof %>
					<option value="<%=rsProducts("sku")%>" <% if rsProducts("sku") = editsku then%>selected<%end if%>><%=rsProducts("product_name")%> - <%=rsProducts("SKU")%></option>
				<% 	rsProducts.MoveNext 
				   Wend %>
			</select>
		</td>
	</tr>
	<tr>
		<td style="width: 20%;" align="right" class="requiredfields">Round 1 Start Date:</td>
		<td>
			<input type="text" name="txtRound1StartDate" id="txtRound1StartDate" maxlength="10" value="<%=editRound1Start%>" style="background-color: #ffffff;" size="12" />
	<script language="javascript">
	new tcal ({
		// form name
		'formname': 'frmUpdate',
		// input name
		'controlname': 'txtRound1StartDate'
	});

	</script>
			<select name="ddlRound1StartTime" id="ddlRound1StartTime">
				<option value="12:00:00 AM" <%if editRound1StartTime = "12:00:00 AM" then%>selected<%end if%>>12:00:00 AM</option>
				<option value="1:00:00 AM" <%if editRound1StartTime = "1:00:00 AM" then%>selected<%end if%>>1:00:00 AM</option>
				<option value="2:00:00 AM" <%if editRound1StartTime = "2:00:00 AM" then%>selected<%end if%>>2:00:00 AM</option>
				<option value="3:00:00 AM" <%if editRound1StartTime = "3:00:00 AM" then%>selected<%end if%>>3:00:00 AM</option>
				<option value="4:00:00 AM" <%if editRound1StartTime = "4:00:00 AM" then%>selected<%end if%>>4:00:00 AM</option>
				<option value="5:00:00 AM" <%if editRound1StartTime = "5:00:00 AM" then%>selected<%end if%>>5:00:00 AM</option>
				<option value="6:00:00 AM" <%if editRound1StartTime = "6:00:00 AM" then%>selected<%end if%>>6:00:00 AM</option>
				<option value="7:00:00 AM" <%if editRound1StartTime = "7:00:00 AM" then%>selected<%end if%>>7:00:00 AM</option>
				<option value="8:00:00 AM" <%if editRound1StartTime = "8:00:00 AM" then%>selected<%end if%>>8:00:00 AM</option>
				<option value="9:00:00 AM" <%if editRound1StartTime = "9:00:00 AM" then%>selected<%end if%>>9:00:00 AM</option>
				<option value="10:00:00 AM" <%if editRound1StartTime = "10:00:00 AM" then%>selected<%end if%>>10:00:00 AM</option>
				<option value="11:00:00 AM" <%if editRound1StartTime = "11:00:00 AM" then%>selected<%end if%>>11:00:00 AM</option>
				<option value="12:00:00 PM" <%if editRound1StartTime = "12:00:00 PM" then%>selected<%end if%>>12:00:00 PM</option>
				<option value="1:00:00 PM" <%if editRound1StartTime = "1:00:00 PM" then%>selected<%end if%>>1:00:00 PM</option>
				<option value="2:00:00 PM" <%if editRound1StartTime = "2:00:00 PM" then%>selected<%end if%>>2:00:00 PM</option>
				<option value="3:00:00 PM" <%if editRound1StartTime = "3:00:00 PM" then%>selected<%end if%>>3:00:00 PM</option>
				<option value="4:00:00 PM" <%if editRound1StartTime = "4:00:00 PM" then%>selected<%end if%>>4:00:00 PM</option>
				<option value="5:00:00 PM" <%if editRound1StartTime = "5:00:00 PM" then%>selected<%end if%>>5:00:00 PM</option>
				<option value="6:00:00 PM" <%if editRound1StartTime = "6:00:00 PM" then%>selected<%end if%>>6:00:00 PM</option>
				<option value="7:00:00 PM" <%if editRound1StartTime = "7:00:00 PM" then%>selected<%end if%>>7:00:00 PM</option>
				<option value="8:00:00 PM" <%if editRound1StartTime = "8:00:00 PM" then%>selected<%end if%>>8:00:00 PM</option>
				<option value="9:00:00 PM" <%if editRound1StartTime = "9:00:00 PM" then%>selected<%end if%>>9:00:00 PM</option>
				<option value="10:00:00 PM" <%if editRound1StartTime = "10:00:00 PM" then%>selected<%end if%>>10:00:00 PM</option>
				<option value="11:00:00 PM" <%if editRound1StartTime = "11:00:00 PM" then%>selected<%end if%>>11:00:00 PM</option>
			</select>
		</td>
		<td style="width: 20%;" align="right" class="requiredfields">Round 1 End Date:</td>
		<td>
			<input type="text" name="txtRound1EndDate" id="txtRound1EndDate" maxlength="10" value="<%=editRound1End%>" style="background-color: #ffffff;" size="12" />
	<script language="javascript">
	new tcal ({
		// form name
		'formname': 'frmUpdate',
		// input name
		'controlname': 'txtRound1EndDate'
	});

	</script>
			<select name="ddlRound1EndTime" id="ddlRound1EndTime">
				<option value="12:00:00 AM" <%if editRound1EndTime = "12:00:00 AM" then%>selected<%end if%>>12:00:00 AM</option>
				<option value="1:00:00 AM" <%if editRound1EndTime = "1:00:00 AM" then%>selected<%end if%>>1:00:00 AM</option>
				<option value="2:00:00 AM" <%if editRound1EndTime = "2:00:00 AM" then%>selected<%end if%>>2:00:00 AM</option>
				<option value="3:00:00 AM" <%if editRound1EndTime = "3:00:00 AM" then%>selected<%end if%>>3:00:00 AM</option>
				<option value="4:00:00 AM" <%if editRound1EndTime = "4:00:00 AM" then%>selected<%end if%>>4:00:00 AM</option>
				<option value="5:00:00 AM" <%if editRound1EndTime = "5:00:00 AM" then%>selected<%end if%>>5:00:00 AM</option>
				<option value="6:00:00 AM" <%if editRound1EndTime = "6:00:00 AM" then%>selected<%end if%>>6:00:00 AM</option>
				<option value="7:00:00 AM" <%if editRound1EndTime = "7:00:00 AM" then%>selected<%end if%>>7:00:00 AM</option>
				<option value="8:00:00 AM" <%if editRound1EndTime = "8:00:00 AM" then%>selected<%end if%>>8:00:00 AM</option>
				<option value="9:00:00 AM" <%if editRound1EndTime = "9:00:00 AM" then%>selected<%end if%>>9:00:00 AM</option>
				<option value="10:00:00 AM" <%if editRound1EndTime = "10:00:00 AM" then%>selected<%end if%>>10:00:00 AM</option>
				<option value="11:00:00 AM" <%if editRound1EndTime = "11:00:00 AM" then%>selected<%end if%>>11:00:00 AM</option>
				<option value="12:00:00 PM" <%if editRound1EndTime = "12:00:00 PM" then%>selected<%end if%>>12:00:00 PM</option>
				<option value="1:00:00 PM" <%if editRound1EndTime = "1:00:00 PM" then%>selected<%end if%>>1:00:00 PM</option>
				<option value="2:00:00 PM" <%if editRound1EndTime = "2:00:00 PM" then%>selected<%end if%>>2:00:00 PM</option>
				<option value="3:00:00 PM" <%if editRound1EndTime = "3:00:00 PM" then%>selected<%end if%>>3:00:00 PM</option>
				<option value="4:00:00 PM" <%if editRound1EndTime = "4:00:00 PM" then%>selected<%end if%>>4:00:00 PM</option>
				<option value="5:00:00 PM" <%if editRound1EndTime = "5:00:00 PM" then%>selected<%end if%>>5:00:00 PM</option>
				<option value="6:00:00 PM" <%if editRound1EndTime = "6:00:00 PM" then%>selected<%end if%>>6:00:00 PM</option>
				<option value="7:00:00 PM" <%if editRound1EndTime = "7:00:00 PM" then%>selected<%end if%>>7:00:00 PM</option>
				<option value="8:00:00 PM" <%if editRound1EndTime = "8:00:00 PM" then%>selected<%end if%>>8:00:00 PM</option>
				<option value="9:00:00 PM" <%if editRound1EndTime = "9:00:00 PM" then%>selected<%end if%>>9:00:00 PM</option>
				<option value="10:00:00 PM" <%if editRound1EndTime = "10:00:00 PM" then%>selected<%end if%>>10:00:00 PM</option>
				<option value="11:00:00 PM" <%if editRound1EndTime = "11:00:00 PM" then%>selected<%end if%>>11:00:00 PM</option>
			</select>
			
		</td>
	</tr>
	<tr>
		<td style="width: 20%;" align="right" class="requiredfields">Round 1 Number of Votes:</td>
		<td><input type="text" name="txtRound1NumberOfVotes" id="txtRound1NumberOfVotes" maxlength="50" value="<%=editRound1NumberOfVotes%>" style="background-color: #ffffff;" /></td>
		<td style="width: 20;" align="right" class="requiredfields">Round 1 Winners:</td>
		<td><input type="text" name="txtRound1Winners" id="txtRound1Winners" maxlength="50" value="<%=editRound1Winners%>" style="background-color: #ffffff;" /></td>
	</tr>
	<tr>
		<td style="width: 20%;" align="right" class="requiredfields">Round 2 Start Date:</td>
		<td>
			<input type="text" name="txtRound2StartDate" id="txtRound2StartDate" maxlength="10" value="<%=editRound2Start%>" style="background-color: #ffffff;" size="12" />
	<script language="javascript">
	new tcal ({
		// form name
		'formname': 'frmUpdate',
		// input name
		'controlname': 'txtRound2StartDate'
	});
	</script>
			<select name="ddlRound2StartTime" id="ddlRound2StartTime">
				<option value="12:00:00 AM" <%if editRound2StartTime = "12:00:00 AM" then%>selected<%end if%>>12:00:00 AM</option>
				<option value="1:00:00 AM" <%if editRound2StartTime = "1:00:00 AM" then%>selected<%end if%>>1:00:00 AM</option>
				<option value="2:00:00 AM" <%if editRound2StartTime = "2:00:00 AM" then%>selected<%end if%>>2:00:00 AM</option>
				<option value="3:00:00 AM" <%if editRound2StartTime = "3:00:00 AM" then%>selected<%end if%>>3:00:00 AM</option>
				<option value="4:00:00 AM" <%if editRound2StartTime = "4:00:00 AM" then%>selected<%end if%>>4:00:00 AM</option>
				<option value="5:00:00 AM" <%if editRound2StartTime = "5:00:00 AM" then%>selected<%end if%>>5:00:00 AM</option>
				<option value="6:00:00 AM" <%if editRound2StartTime = "6:00:00 AM" then%>selected<%end if%>>6:00:00 AM</option>
				<option value="7:00:00 AM" <%if editRound2StartTime = "7:00:00 AM" then%>selected<%end if%>>7:00:00 AM</option>
				<option value="8:00:00 AM" <%if editRound2StartTime = "8:00:00 AM" then%>selected<%end if%>>8:00:00 AM</option>
				<option value="9:00:00 AM" <%if editRound2StartTime = "9:00:00 AM" then%>selected<%end if%>>9:00:00 AM</option>
				<option value="10:00:00 AM" <%if editRound2StartTime = "10:00:00 AM" then%>selected<%end if%>>10:00:00 AM</option>
				<option value="11:00:00 AM" <%if editRound2StartTime = "11:00:00 AM" then%>selected<%end if%>>11:00:00 AM</option>
				<option value="12:00:00 PM" <%if editRound2StartTime = "12:00:00 PM" then%>selected<%end if%>>12:00:00 PM</option>
				<option value="1:00:00 PM" <%if editRound2StartTime = "1:00:00 PM" then%>selected<%end if%>>1:00:00 PM</option>
				<option value="2:00:00 PM" <%if editRound2StartTime = "2:00:00 PM" then%>selected<%end if%>>2:00:00 PM</option>
				<option value="3:00:00 PM" <%if editRound2StartTime = "3:00:00 PM" then%>selected<%end if%>>3:00:00 PM</option>
				<option value="4:00:00 PM" <%if editRound2StartTime = "4:00:00 PM" then%>selected<%end if%>>4:00:00 PM</option>
				<option value="5:00:00 PM" <%if editRound2StartTime = "5:00:00 PM" then%>selected<%end if%>>5:00:00 PM</option>
				<option value="6:00:00 PM" <%if editRound2StartTime = "6:00:00 PM" then%>selected<%end if%>>6:00:00 PM</option>
				<option value="7:00:00 PM" <%if editRound2StartTime = "7:00:00 PM" then%>selected<%end if%>>7:00:00 PM</option>
				<option value="8:00:00 PM" <%if editRound2StartTime = "8:00:00 PM" then%>selected<%end if%>>8:00:00 PM</option>
				<option value="9:00:00 PM" <%if editRound2StartTime = "9:00:00 PM" then%>selected<%end if%>>9:00:00 PM</option>
				<option value="10:00:00 PM" <%if editRound2StartTime = "10:00:00 PM" then%>selected<%end if%>>10:00:00 PM</option>
				<option value="11:00:00 PM" <%if editRound2StartTime = "11:00:00 PM" then%>selected<%end if%>>11:00:00 PM</option>
			</select>

		</td>
		<td style="width: 20%;" align="right" class="requiredfields">Round 2 End Date:</td>
		<td>
			<input type="text" name="txtRound2EndDate" id="txtRound2EndDate" maxlength="10" value="<%=editRound2End%>" style="background-color: #ffffff;" size="12" />
	<script language="javascript">
	new tcal ({
		// form name
		'formname': 'frmUpdate',
		// input name
		'controlname': 'txtRound2EndDate'
	});
	</script>
			<select name="ddlRound2EndTime" id="ddlRound2EndTime">
				<option value="12:00:00 AM" <%if editRound2EndTime = "12:00:00 AM" then%>selected<%end if%>>12:00:00 AM</option>
				<option value="1:00:00 AM" <%if editRound2EndTime = "1:00:00 AM" then%>selected<%end if%>>1:00:00 AM</option>
				<option value="2:00:00 AM" <%if editRound2EndTime = "2:00:00 AM" then%>selected<%end if%>>2:00:00 AM</option>
				<option value="3:00:00 AM" <%if editRound2EndTime = "3:00:00 AM" then%>selected<%end if%>>3:00:00 AM</option>
				<option value="4:00:00 AM" <%if editRound2EndTime = "4:00:00 AM" then%>selected<%end if%>>4:00:00 AM</option>
				<option value="5:00:00 AM" <%if editRound2EndTime = "5:00:00 AM" then%>selected<%end if%>>5:00:00 AM</option>
				<option value="6:00:00 AM" <%if editRound2EndTime = "6:00:00 AM" then%>selected<%end if%>>6:00:00 AM</option>
				<option value="7:00:00 AM" <%if editRound2EndTime = "7:00:00 AM" then%>selected<%end if%>>7:00:00 AM</option>
				<option value="8:00:00 AM" <%if editRound2EndTime = "8:00:00 AM" then%>selected<%end if%>>8:00:00 AM</option>
				<option value="9:00:00 AM" <%if editRound2EndTime = "9:00:00 AM" then%>selected<%end if%>>9:00:00 AM</option>
				<option value="10:00:00 AM" <%if editRound2EndTime = "10:00:00 AM" then%>selected<%end if%>>10:00:00 AM</option>
				<option value="11:00:00 AM" <%if editRound2EndTime = "11:00:00 AM" then%>selected<%end if%>>11:00:00 AM</option>
				<option value="12:00:00 PM" <%if editRound2EndTime = "12:00:00 PM" then%>selected<%end if%>>12:00:00 PM</option>
				<option value="1:00:00 PM" <%if editRound2EndTime = "1:00:00 PM" then%>selected<%end if%>>1:00:00 PM</option>
				<option value="2:00:00 PM" <%if editRound2EndTime = "2:00:00 PM" then%>selected<%end if%>>2:00:00 PM</option>
				<option value="3:00:00 PM" <%if editRound2EndTime = "3:00:00 PM" then%>selected<%end if%>>3:00:00 PM</option>
				<option value="4:00:00 PM" <%if editRound2EndTime = "4:00:00 PM" then%>selected<%end if%>>4:00:00 PM</option>
				<option value="5:00:00 PM" <%if editRound2EndTime = "5:00:00 PM" then%>selected<%end if%>>5:00:00 PM</option>
				<option value="6:00:00 PM" <%if editRound2EndTime = "6:00:00 PM" then%>selected<%end if%>>6:00:00 PM</option>
				<option value="7:00:00 PM" <%if editRound2EndTime = "7:00:00 PM" then%>selected<%end if%>>7:00:00 PM</option>
				<option value="8:00:00 PM" <%if editRound2EndTime = "8:00:00 PM" then%>selected<%end if%>>8:00:00 PM</option>
				<option value="9:00:00 PM" <%if editRound2EndTime = "9:00:00 PM" then%>selected<%end if%>>9:00:00 PM</option>
				<option value="10:00:00 PM" <%if editRound2EndTime = "10:00:00 PM" then%>selected<%end if%>>10:00:00 PM</option>
				<option value="11:00:00 PM" <%if editRound2EndTime = "11:00:00 PM" then%>selected<%end if%>>11:00:00 PM</option>
			</select>

		</td>
	</tr>
	<tr>
		<td style="width: 20%;" align="right" class="requiredfields">Round 2 Number of Votes:</td>
		<td><input type="text" name="txtRound2NumberOfVotes" id="txtRound2NumberOfVotes" maxlength="50" value="<%=editRound2NumberOfVotes%>" style="background-color: #ffffff;" /></td>
		<td style="width: 20%;" align="right" class="requiredfields">Round 2 Winners:</td>
		<td><input type="text" name="txtRound2Winners" id="txtRound2Winners" maxlength="50" value="<%=editRound2Winners%>" style="background-color: #ffffff;" /></td>
	</tr>
	<tr>
		<td style="width: 20%;" align="right" class="requiredfields">Round 3 Start Date:</td>
		<td>
			<input type="text" name="txtRound3StartDate" id="txtRound3StartDate" maxlength="10" value="<%=editRound3Start%>" style="background-color: #ffffff;" size="12" />
	<script language="javascript">
	new tcal ({
		// form name
		'formname': 'frmUpdate',
		// input name
		'controlname': 'txtRound3StartDate'
	});
	</script>

			<select name="ddlRound3StartTime" id="ddlRound3StartTime">
				<option value="12:00:00 AM" <%if editRound3StartTime = "12:00:00 AM" then%>selected<%end if%>>12:00:00 AM</option>
				<option value="1:00:00 AM" <%if editRound3StartTime = "1:00:00 AM" then%>selected<%end if%>>1:00:00 AM</option>
				<option value="2:00:00 AM" <%if editRound3StartTime = "2:00:00 AM" then%>selected<%end if%>>2:00:00 AM</option>
				<option value="3:00:00 AM" <%if editRound3StartTime = "3:00:00 AM" then%>selected<%end if%>>3:00:00 AM</option>
				<option value="4:00:00 AM" <%if editRound3StartTime = "4:00:00 AM" then%>selected<%end if%>>4:00:00 AM</option>
				<option value="5:00:00 AM" <%if editRound3StartTime = "5:00:00 AM" then%>selected<%end if%>>5:00:00 AM</option>
				<option value="6:00:00 AM" <%if editRound3StartTime = "6:00:00 AM" then%>selected<%end if%>>6:00:00 AM</option>
				<option value="7:00:00 AM" <%if editRound3StartTime = "7:00:00 AM" then%>selected<%end if%>>7:00:00 AM</option>
				<option value="8:00:00 AM" <%if editRound3StartTime = "8:00:00 AM" then%>selected<%end if%>>8:00:00 AM</option>
				<option value="9:00:00 AM" <%if editRound3StartTime = "9:00:00 AM" then%>selected<%end if%>>9:00:00 AM</option>
				<option value="10:00:00 AM" <%if editRound3StartTime = "10:00:00 AM" then%>selected<%end if%>>10:00:00 AM</option>
				<option value="11:00:00 AM" <%if editRound3StartTime = "11:00:00 AM" then%>selected<%end if%>>11:00:00 AM</option>
				<option value="12:00:00 PM" <%if editRound3StartTime = "12:00:00 PM" then%>selected<%end if%>>12:00:00 PM</option>
				<option value="1:00:00 PM" <%if editRound3StartTime = "1:00:00 PM" then%>selected<%end if%>>1:00:00 PM</option>
				<option value="2:00:00 PM" <%if editRound3StartTime = "2:00:00 PM" then%>selected<%end if%>>2:00:00 PM</option>
				<option value="3:00:00 PM" <%if editRound3StartTime = "3:00:00 PM" then%>selected<%end if%>>3:00:00 PM</option>
				<option value="4:00:00 PM" <%if editRound3StartTime = "4:00:00 PM" then%>selected<%end if%>>4:00:00 PM</option>
				<option value="5:00:00 PM" <%if editRound3StartTime = "5:00:00 PM" then%>selected<%end if%>>5:00:00 PM</option>
				<option value="6:00:00 PM" <%if editRound3StartTime = "6:00:00 PM" then%>selected<%end if%>>6:00:00 PM</option>
				<option value="7:00:00 PM" <%if editRound3StartTime = "7:00:00 PM" then%>selected<%end if%>>7:00:00 PM</option>
				<option value="8:00:00 PM" <%if editRound3StartTime = "8:00:00 PM" then%>selected<%end if%>>8:00:00 PM</option>
				<option value="9:00:00 PM" <%if editRound3StartTime = "9:00:00 PM" then%>selected<%end if%>>9:00:00 PM</option>
				<option value="10:00:00 PM" <%if editRound3StartTime = "10:00:00 PM" then%>selected<%end if%>>10:00:00 PM</option>
				<option value="11:00:00 PM" <%if editRound3StartTime = "11:00:00 PM" then%>selected<%end if%>>11:00:00 PM</option>
			</select>

		</td>
		<td style="width: 20%;" align="right" class="requiredfields">Round 3 End Date:</td>
		<td>
			<input type="text" name="txtRound3EndDate" id="txtRound3EndDate" maxlength="10" value="<%=editRound3End%>" style="background-color: #ffffff;" size="12" />
	<script language="javascript">
	new tcal ({
		// form name
		'formname': 'frmUpdate',
		// input name
		'controlname': 'txtRound3EndDate'
	});
	</script>

			<select name="ddlRound3EndTime" id="ddlRound3EndTime">
				<option value="12:00:00 AM" <%if editRound3EndTime = "12:00:00 AM" then%>selected<%end if%>>12:00:00 AM</option>
				<option value="1:00:00 AM" <%if editRound3EndTime = "1:00:00 AM" then%>selected<%end if%>>1:00:00 AM</option>
				<option value="2:00:00 AM" <%if editRound3EndTime = "2:00:00 AM" then%>selected<%end if%>>2:00:00 AM</option>
				<option value="3:00:00 AM" <%if editRound3EndTime = "3:00:00 AM" then%>selected<%end if%>>3:00:00 AM</option>
				<option value="4:00:00 AM" <%if editRound3EndTime = "4:00:00 AM" then%>selected<%end if%>>4:00:00 AM</option>
				<option value="5:00:00 AM" <%if editRound3EndTime = "5:00:00 AM" then%>selected<%end if%>>5:00:00 AM</option>
				<option value="6:00:00 AM" <%if editRound3EndTime = "6:00:00 AM" then%>selected<%end if%>>6:00:00 AM</option>
				<option value="7:00:00 AM" <%if editRound3EndTime = "7:00:00 AM" then%>selected<%end if%>>7:00:00 AM</option>
				<option value="8:00:00 AM" <%if editRound3EndTime = "8:00:00 AM" then%>selected<%end if%>>8:00:00 AM</option>
				<option value="9:00:00 AM" <%if editRound3EndTime = "9:00:00 AM" then%>selected<%end if%>>9:00:00 AM</option>
				<option value="10:00:00 AM" <%if editRound3EndTime = "10:00:00 AM" then%>selected<%end if%>>10:00:00 AM</option>
				<option value="11:00:00 AM" <%if editRound3EndTime = "11:00:00 AM" then%>selected<%end if%>>11:00:00 AM</option>
				<option value="12:00:00 PM" <%if editRound3EndTime = "12:00:00 PM" then%>selected<%end if%>>12:00:00 PM</option>
				<option value="1:00:00 PM" <%if editRound3EndTime = "1:00:00 PM" then%>selected<%end if%>>1:00:00 PM</option>
				<option value="2:00:00 PM" <%if editRound3EndTime = "2:00:00 PM" then%>selected<%end if%>>2:00:00 PM</option>
				<option value="3:00:00 PM" <%if editRound3EndTime = "3:00:00 PM" then%>selected<%end if%>>3:00:00 PM</option>
				<option value="4:00:00 PM" <%if editRound3EndTime = "4:00:00 PM" then%>selected<%end if%>>4:00:00 PM</option>
				<option value="5:00:00 PM" <%if editRound3EndTime = "5:00:00 PM" then%>selected<%end if%>>5:00:00 PM</option>
				<option value="6:00:00 PM" <%if editRound3EndTime = "6:00:00 PM" then%>selected<%end if%>>6:00:00 PM</option>
				<option value="7:00:00 PM" <%if editRound3EndTime = "7:00:00 PM" then%>selected<%end if%>>7:00:00 PM</option>
				<option value="8:00:00 PM" <%if editRound3EndTime = "8:00:00 PM" then%>selected<%end if%>>8:00:00 PM</option>
				<option value="9:00:00 PM" <%if editRound3EndTime = "9:00:00 PM" then%>selected<%end if%>>9:00:00 PM</option>
				<option value="10:00:00 PM" <%if editRound3EndTime = "10:00:00 PM" then%>selected<%end if%>>10:00:00 PM</option>
				<option value="11:00:00 PM" <%if editRound3EndTime = "11:00:00 PM" then%>selected<%end if%>>11:00:00 PM</option>
			</select>

		</td>
	</tr>
	<tr>
		<td style="width: 20%;" align="right" class="requiredfields">Round 3 Number of Votes:</td>
		<td><input type="text" name="txtRound3NumberOfVotes" id="txtRound3NumberOfVotes" maxlength="50" value="<%=editRound3NumberOfVotes%>" style="background-color: #ffffff;" /></td>
		<td style="width: 20%;" align="right" class="requiredfields">Round 3 Winners:</td>
		<td><input type="text" name="txtRound3Winners" id="txtRound3Winners" maxlength="50" value="<%=editRound3Winners%>" style="background-color: #ffffff;" /></td>
	</tr>
	<tr>
		<td style="width: 20%;" align="right" class="requiredfields">Upload Image Cut Off Date:</td>
		<td colspan="3">
			<input type="text" name="txtCutOffDate" id="txtCutOffDate" maxlength="10" value="<%=editCutOffDate%>" style="background-color: #ffffff;" />
	<script language="javascript">
	new tcal ({
		// form name
		'formname': 'frmUpdate',
		// input name
		'controlname': 'txtCutOffDate'
	});
	</script>
			<select name="ddlCutoffStartTime" id="ddlCutoffStartTime">
				<option value="12:00:00 AM" <%if editCutoffStartTime = "12:00:00 AM" then%>selected<%end if%>>12:00:00 AM</option>
				<option value="1:00:00 AM" <%if editCutoffStartTime = "1:00:00 AM" then%>selected<%end if%>>1:00:00 AM</option>
				<option value="2:00:00 AM" <%if editCutoffStartTime = "2:00:00 AM" then%>selected<%end if%>>2:00:00 AM</option>
				<option value="3:00:00 AM" <%if editCutoffStartTime = "3:00:00 AM" then%>selected<%end if%>>3:00:00 AM</option>
				<option value="4:00:00 AM" <%if editCutoffStartTime = "4:00:00 AM" then%>selected<%end if%>>4:00:00 AM</option>
				<option value="5:00:00 AM" <%if editCutoffStartTime = "5:00:00 AM" then%>selected<%end if%>>5:00:00 AM</option>
				<option value="6:00:00 AM" <%if editCutoffStartTime = "6:00:00 AM" then%>selected<%end if%>>6:00:00 AM</option>
				<option value="7:00:00 AM" <%if editCutoffStartTime = "7:00:00 AM" then%>selected<%end if%>>7:00:00 AM</option>
				<option value="8:00:00 AM" <%if editCutoffStartTime = "8:00:00 AM" then%>selected<%end if%>>8:00:00 AM</option>
				<option value="9:00:00 AM" <%if editCutoffStartTime = "9:00:00 AM" then%>selected<%end if%>>9:00:00 AM</option>
				<option value="10:00:00 AM" <%if editCutoffStartTime = "10:00:00 AM" then%>selected<%end if%>>10:00:00 AM</option>
				<option value="11:00:00 AM" <%if editCutoffStartTime = "11:00:00 AM" then%>selected<%end if%>>11:00:00 AM</option>
				<option value="12:00:00 PM" <%if editCutoffStartTime = "12:00:00 PM" then%>selected<%end if%>>12:00:00 PM</option>
				<option value="1:00:00 PM" <%if editCutoffStartTime = "1:00:00 PM" then%>selected<%end if%>>1:00:00 PM</option>
				<option value="2:00:00 PM" <%if editCutoffStartTime = "2:00:00 PM" then%>selected<%end if%>>2:00:00 PM</option>
				<option value="3:00:00 PM" <%if editCutoffStartTime = "3:00:00 PM" then%>selected<%end if%>>3:00:00 PM</option>
				<option value="4:00:00 PM" <%if editCutoffStartTime = "4:00:00 PM" then%>selected<%end if%>>4:00:00 PM</option>
				<option value="5:00:00 PM" <%if editCutoffStartTime = "5:00:00 PM" then%>selected<%end if%>>5:00:00 PM</option>
				<option value="6:00:00 PM" <%if editCutoffStartTime = "6:00:00 PM" then%>selected<%end if%>>6:00:00 PM</option>
				<option value="7:00:00 PM" <%if editCutoffStartTime = "7:00:00 PM" then%>selected<%end if%>>7:00:00 PM</option>
				<option value="8:00:00 PM" <%if editCutoffStartTime = "8:00:00 PM" then%>selected<%end if%>>8:00:00 PM</option>
				<option value="9:00:00 PM" <%if editCutoffStartTime = "9:00:00 PM" then%>selected<%end if%>>9:00:00 PM</option>
				<option value="10:00:00 PM" <%if editCutoffStartTime = "10:00:00 PM" then%>selected<%end if%>>10:00:00 PM</option>
				<option value="11:00:00 PM" <%if editCutoffStartTime = "11:00:00 PM" then%>selected<%end if%>>11:00:00 PM</option>
			</select>

		</td>
	</tr>
	<tr>
		<td><input type="submit" name="btnSave" id="btnSave" value="Save" style="background-color: silver;" /></td>
		<td colspan="3"><br /></td>
	</tr>
</table>
<% end if %>
</form>
    <!-----------------------End Main Body -------------------------------------->
    </td>	
      </tr>
 </table></tr>
<!-- footer start -->
 <!--#include Virtual="/footer.asp"-->
  <!-- footer end -->
</table>
</body>
</html>
