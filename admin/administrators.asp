<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!-- #include file ="dbconnections.asp" -->
<!-- #include file ="clAdministrators.asp" -->
<!-- #include Virtual ="/constants.asp" -->

<!-- #include file = "adminsession.asp" -->

<%

Dim rsAdministrators
Dim rsAdministrator
Dim sAction
Dim editID
Dim editFirstName
Dim editLastName
Dim editUser
Dim editPassword
Dim editEmail

sAction = "view"
if Request.QueryString("action") <> "" then
	sAction = Request.QueryString("action")
end if

if sAction = "view" then
	set rsAdministrators = GetAdministrators()
	HeaderText = "Administrator Accounts"
elseif sAction = "add" then
	HeaderText = "Add Administrator Accounts"
elseif sAction = "edit" then
	HeaderText = "Update Administrator Accounts"
	editID = Request.QueryString("id")
	set rsAdministrator = getAdministrator(editID)
	if not rsAdministrator.eof then
		editFirstName = rsAdministrator("First_Name")
		editLastName = rsAdministrator("Last_Name")
		editUser = rsAdministrator("User")
		editPassword = rsAdministrator("Password")
		editEmail = rsAdministrator("Email")
	end if
end if


%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Member Sign-up</title>
<link href="../css/style.css" rel="stylesheet" type="text/css" />
<LINK href="../css/styles.css" rel="stylesheet" type="text/css">

<script language="javascript">
<!--
function validate() {

	if (document.frmUpdate.txtFirstName.value == "") {
		alert("Please enter in the administrators's first name.");
		return false;
	}
	if (document.frmUpdate.txtLastName.value == "") {
		alert("Please enter in the administrator's last name.");
		return false;
	}
	if (document.frmUpdate.txtUser.value == "") {
		alert("Please enter in the administrator's User Name.");
		return false;
	}
	if (document.frmUpdate.txtPassword.value == "") {
		alert("Please enter in the administrator's password.");
		return false;
	}
	if (document.frmUpdate.txtEmail.value == "") {
		alert("Please enter in the administrator's email.");
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
  <!--#include Virtual="/header.asp"-->  
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
	<td width="848"  align="left" valign="top" bgcolor="#ffffff" class="bdytxt">
    <!---------------------------Main Body -------------------------------------->
<form action="save.asp" method="post" name="frmUpdate" id="frmUpdate">
<input type="hidden" name="hdnSection" id="hdnSection" value="admin" />
<input type="hidden" name="hdnAction" id="hdnAction" value="<%=sAction%>" />
<input type="hidden" name="hdnID" id="hdnID" value="<%=editID%>" />

<% if sAction = "view" then %>
<table width="100%" cellpadding="2" cellspacing="0" border="0" bgcolor="#ffffff">
	<tr>
		<td colspan="5">
			<font style="font-size: large;"><%=HeaderText%></font><br /><br />
			<a href="administrators.asp?action=add" style="background-color: #ffffff;">Add Administrator Account</a>
		</td>
	</tr>
	<tr style="background-color: gray;">
		<td style="width: 15%;">Administrator</td>
		<td style="width: 15%;">User</td>
		<td style="width: 25%;">Email</td>
		<td style="width: 5%;"><br /></td>
		<td style="width: 5%;"><br /></td>
	</tr>
	<% while not rsAdministrators.eof %>
	<tr>
		<td><%=rsAdministrators("First_Name")%>&nbsp;<%=rsAdministrators("Last_Name")%></td>
		<td><%=rsAdministrators("User")%></td>
		<td><%=rsAdministrators("Email")%></td>
		<td><a href="administrators.asp?action=edit&id=<%=rsAdministrators("id")%>">edit</a></td>
		<td><a href="javascript:confirmDelete(<%=rsAdministrators("id")%>)">delete</a></td>
		<% rsAdministrators.MoveNext %>
	</tr>
	<% wend %>
</table>
<% end if %>

<% if sAction = "edit" or sAction = "add" then %>
<table width="100%" cellpadding="2" cellspacing="0" border="0" bgcolor="#ffffff">
	<tr>
		<td colspan="2" align="center"><font style="font-size: large;"><%=HeaderText%></font></td>
	</tr>
	<tr>
		<td colspan="2"><br /></td>
	</tr>
	<tr>
		<td style="width: 20%;" align="right" class="requiredfields">First Name:</td>
		<td><input type="text" name="txtFirstName" id="txtFirstName" maxlength="50" width="50" value="<%=editFirstName%>" style="background-color: #ffffff;" /></td>
	</tr>
	<tr>
		<td style="width: 20%;" align="right" class="requiredfields">Last Name:</td>
		<td><input type="text" name="txtLastName" id="txtLastName" maxlength="50" width="50" value="<%=editLastName%>" style="background-color: #ffffff;" /></td>
	</tr>
	<tr>
		<td style="width: 20%;" align="right" class="requiredfields">User:</td>
		<td><input type="text" name="txtUser" id="txtUser" maxlength="15" width="15" value="<%=editUser%>" style="background-color: #ffffff;" /></td>
	</tr>
	<tr>
		<td style="width: 20%;" align="right" class="requiredfields">Password:</td>
		<td><input type="password" name="txtPassword" id="txtPassword" maxlength="15" width="15" value="<%=editPassword%>" style="background-color: #ffffff;" /></td>
	</tr>
	<tr>
		<td style="width: 20%;" align="right" class="requiredfields">Email:</td>
		<td><input type="text" name="txtEmail" id="txtEmail" maxlength="50" width="50" value="<%=editEmail%>" style="background-color: #ffffff;" /></td>
	</tr>
	<tr>
		<td align="right"><input type="submit" name="btnSave" id="btnSave" value="Save" style="background-color: silver;" /></td>
		<td><br /></td>
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
