<%





Sub SendEmail( sRecipientEmail, sRecipientName, sSubject, sBody)
		
	'------------------------------------------------------------------------------------
	'create an Asp mail component.	
	'------------------------------------------------------------------------------------
	Set Mail = CreateObject("CDO.Message")
	Mail.From = "donotreply@membercalendars.com"

 
 
	'------------------------------------------------------------------------------------
	' The mail server you have to use with Asp Mail
	'------------------------------------------------------------------------------------
	Mail.To= sRecipientEmail
	Mail.Subject = sSubject
	Mail.HTMLBody = sBody

	'On Error Resume Next 


	Mail.Send
	If Err <> 0 Then
  		'Response.Write "An error occurred: " & Err.Description
	End If 

	set Mail = Nothing
End Sub

Sub clEmail_ValidateUserEmail(UserName, EmailAddress)

Dim sBody
Dim sSubject

sFrom = "donotreply@membercalendars.com"
sSubject = "Your membercalendars.com Membership is Complete!"


sBody = "Dear " & UserName & ",<br><br>" & _
"Your membercalendars.com membership is now active!  Log in to membercalendars.com , to start voting and upload a photo to be " & _
"voted on by your peers!<br><br>" & _
"There will be 13 winning photos (1 cover picture + 12 monthly pictures) chosen to be published in an officially licensed 2011 calendar! " & _
"The days of buying stock image calendars off the shelves at your local grocery store or retailer are over.<br><br>" & _
"We thank you for choosing membercalendars.com and realizing that you have a voice when it comes to the calendar hanging on your wall " & _
"365 days a year!<br><br>" & _
"Regards,<br><br>" & _
"Membercalendars.com"

SendEmail EmailAddress, UserName, sSubject, sBody

End Sub

Sub clEmail_NewUserRegistration(UserName, EmailAddress)

Dim sBody
Dim sSubject

sFrom = "donotreply@membercalendars.com"
sSubject = "Validate Your membercalendars.com Membership!"

sBody = "Dear " & UserName & ",<br><br>" & _
"Thank you for registering @ membercalendars.com, the only website that allows car enthusiasts to participate in a national competition for a chance to be published in an officially licensed calendar.  Click this link, <a href='http://membercalendars.com/ValidateMembership.asp?User=" & UserName & "&email=" & EmailAddress & "'> <a href='http://membercalendars.com/ValidateMembership.asp?User=" & UserName & "&email=" & EmailAddress & "'>http://membercalendars.com/ValidateMembership.asp?User=" & Username  &"&email=" & EmailAddress &"</a> to complete your membership registration!<br><br>" & _
"Regards,<br><br>" & _
"Membercalendars.com"

SendEmail EmailAddress, UserName, sSubject, sBody

End Sub

Sub clEmail_PhotoApproval(PhotoID)

Dim sBody
Dim sSubject
Dim UserName
Dim EmailAddress
Dim Round1Start

set rsGetEmailInformation = GetEmailInformation(PhotoID)

UserName = rsGetEmailInformation("User")
EmailAddress = rsGetEmailInformation("Email")
Round1Start = FormatDateTime(rsGetEmailInformation("Round_1_Start"),1)

sFrom = "donotreply@membercalendars.com"
sSubject = "Your photo has been approved!"

sBody = "Dear " & UserName & ",<br><br>" & _
"Congratulations, your photo has been approved!  Please log in to " & _
"http://membercalendars.com/login.asp to view your photo and others in the competition! Remember that you can only upload 1 photo per calendar competition. " & _
"Voting for this calendar is underway and you will be notified if your photo is a winner!<br><br>" & _
"Thanks,<br><br>" & _
"Membercalendars.com"



SendEmail EmailAddress, UserName, sSubject, sBody


End Sub

Sub clEmail_PhotoNotApproved(PhotoID)

Dim sBody
Dim sSubject
Dim UserName
Dim EmailAddress

set rsGetEmailInformation = GetEmailInformation(PhotoID)

UserName = rsGetEmailInformation("User")
EmailAddress = rsGetEmailInformation("Email")
sSubject = "Your uploaded photo was not approved"

sBody = "Dear " & UserName & ",<br><br>" & _
"We're sorry but your photo was not approved for display in our voting contest. Possible reasons may include but not be limited to the following:  inappropriate, profane, defamatory, obscene, indecent, or unlawful material or content.<br><br>All photos must also be entered into the appropriate contest (i.e.  Mustang photos cannot be entered into the Corvette contest)<br><br>If you have any questions of concerns in regards to this incident, please contact us at membercalendars.com/contact<br><br>Membercalendars.com reserves the right to remove any photo without cause or explanation.<br><br>Thanks,<br><br>MemberCalendars.com"

SendEmail EmailAddress, UserName, sSubject, sBody

End Sub


Sub clEmail_PurchaseInfo(MemberName, ConfirmationNumber, OrderID)

Dim sSubject
Dim sBody

sFrom = "confirmation@membercalendars.com"
sSubject = "Your order with membercalendars.com is confirmed"

sBody = "Dear " & MemberName & "," & _
"Thank you for ordering from membercalendars.com!  Your confirmation number is " & ConfirmationNumber & "." 

'Quantity:   2 			
'Product:   G8board.com 2011 calendar		
'Price:    $14.95

'Shipping & Handling:   $2.00	
'Tax:  0.00
'Total:  	$16.95	
'Delivery Estimate:  Dec. 15, 2010

sBody = sBody & "Thanks again for choosing membercalendars.com"

End Sub

Sub clEmail_OrderHasShipped()

Dim sSubject
Dim sBody 

sFrom = "shipping@membercalendars.com"
sSubject = "Your “G8board.com.com 2011 Wall Calendar” order has been shipped!"

sBody = "Dear " & MemberName & "," & _
"Membercalendars.com has shipped your “G8board.com.com 2011 Wall Calendar." 

'Revised delivery estimate:  Dec. 21st

'Shipment was sent to: 

'“Gene Wooden
'204 Hartford Court
'Oswego, IL 60543”

'This shipment was delivered by USPS.  Track your order now by clicking on the link below.

'http://www.usps.com/shipping/trackandconfirm.htm 

'Carrier tracking ID:  00000014456789524364565

'No more changes can be made to this order.  If you have any questions about this order, please visit www.membercalendars.com/contactus
'Thanks,

'Membercalendars.com

End Sub


Sub clEmail_PhotoHasBeenReported(PhotoID)

Dim sSubject
Dim sBody
Dim rsGetEmailInformation
Dim UserName
Dim EmailAddress

set rsGetEmailInformation = GetEmailInformation(PhotoID)

UserName = rsGetEmailInformation("User")
EmailAddress = rsGetEmailInformation("Email")

sFrom = "donotreply@membercalendars.com"
sSubject = "Your photo on membercalendars.com has been reported and removed"

sBody = "Dear " & UserName & ",<br><br>" & _
"The photo that you have uploaded to membercalendars.com has been reported.   Your photo is currently under review.  You will be notified by email whether Membercalendars.com has approved or removed your photo.<br><br>" & _
"If you have any questions or concerns in regards to this incident, please contact us at membercalendars.com/contact<br><br> " & _
"Membercalendars.com reserves the right to remove any photo without cause or explanation.<br><br>" & _
"Thanks,<br><br>" & _
"Membercalendars.com"

SendEmail EmailAddress, UserName, sSubject, sBody


End Sub


' ==============================================================================================
' Send an email when the user requests to know their username and password
' ==============================================================================================
Sub clEmail_ForgotPassword(UserName, Password, EmailAddress)

Dim sSubject
Dim sBody

sFrom = "donotreply@membercalendars.com"
sSubject = "Membercalendars.com Username and Password Reminder"

sBody = "Dear " & UserName & ",<br><br>" & _
"In response to your recent request to be reminded of your membercalendars.com login information, please see below the information we have on file for you.<br><br>" & _
"Username: " & UserName & "<br>" & _
"Password: " & Password & "<br><br>" & _
"Thank you,<br><br>" & _
"Membercalendars.com"

SendEmail EmailAddress, UserName, sSubject, sBody


End Sub

Sub clEmail_FirstRoundVotingCompleted(PhotoID)

Dim sBody
Dim sSubject
Dim UserName
Dim EmailAddress
Dim Round2Start
Dim Round2End

set rsGetEmailInformation = GetEmailInformation(PhotoID)

if rsGetEmailInformation.RecordCount > 0 then
UserName = rsGetEmailInformation("User")
EmailAddress = rsGetEmailInformation("Email")
Round2Start = FormatDateTime(rsGetEmailInformation("Round_2_Start"),1)
Round2End = FormatDateTime(rsGetEmailInformation("Round_2_End"),1)

sFrom = "donotreply@membercalendars.com"
sSubject = "Congratulations!  Your photo has made it to the 2nd Round of Voting"

sBody = "Dear " & UserName & ",<br><br>" & _
"Congratulations, your photo has received enough votes to enter the 2nd Round of voting.   Voting will commence " & Round2Start & " and run through " & Round2End & ".<br><br>" & _
"Don’t miss your chance to win a spot in an officially licensed calendar!!!   Make sure that your friends and family go to http://membercalendars.com/login.asp to vote for your photo!<br><br>" & _
"Thanks,<br><br>" & _
"Membercalendars.com"

SendEmail EmailAddress, UserName, sSubject, sBody
End if

End Sub

Sub clEmail_SecondRoundVotingCompleted(PhotoID)

Dim sSubject
Dim sBody
Dim UserName
Dim EmailAddress
Dim Round3Start
Dim Round3End

set rsGetEmailInformation = GetEmailInformation(PhotoID)

if not rsGetEmailInformation.eof then
	UserName = rsGetEmailInformation("User")
	EmailAddress = rsGetEmailInformation("Email")
	Round3Start = FormatDateTime(rsGetEmailInformation("Round_3_Start"),1)
	Round3End = FormatDateTime(rsGetEmailInformation("Round_3_End"),1)
	
	sFrom = "donotreply@membercalendars.com"
	sSubject = "Congratulations!  Your photo has made it to the Final Round of Voting" 
	
	sBody = "Dear " & UserName & ",<br><br>" & _
	"Congratulations, your photo has received enough votes to enter into the FINAL round of voting.   Voting will commence " & Round3Start & " and run through " & Round3End & ".<br><br>" & _
	"Don’t miss your chance to win a spot in an officially licensed calendar!!!   Make sure that your friends and family go to http://membercalendars.com/login.asp to vote for your photo!<br><br>" & _
	"Thanks,<br><br>" & _
	"Membercalendars.com"
	
	SendEmail EmailAddress, UserName, sSubject, sBody
	
end if

End Sub

Sub clEmail_FinalRoundVotingResults(PhotoID)

Dim sSubject
Dim sBody
Dim UserName
Dim EmailAddress
Dim Round3Start
Dim Round3End
Dim ProductName
Dim CutOffDate

set rsGetEmailInformation = GetEmailInformation(PhotoID)

UserName = rsGetEmailInformation("User")
EmailAddress = rsGetEmailInformation("Email")
Round3Start = FormatDateTime(rsGetEmailInformation("Round_3_Start"),1)
Round3End = FormatDateTime(rsGetEmailInformation("Round_3_End"),1)
ProductName = rsGetEmailInformation("Product_Name")
CutOffDate = FormatDateTime(DateAdd("d", 20, now()),1)

sFrom = "donotreply@membercalendars.com"
sSubject = "Congratulations!!!   Your photo has been selected to be in the " & ProductName & _
" !!"

sBody = "Dear " & UserName & ",<br><br>" & _
"Congratulations, your photo has WON!!!  Your photo received enough votes to be placed in " & _
"this year’s " & ProductName & ".<br><br>" & _
"Your photo is currently pending review for final approval.<br><br>A separate email will be sent out in the near future to obtain your shipping information to receive your 2 free calendars! <br><br>" & _
"Once again congratulations!<br><br>" & _
"Thanks,<br><br>" & _
"Membercalendars.com"

SendEmail EmailAddress, UserName, sSubject, sBody


End Sub


Sub clEmail_PhotoDidNotReceiveEnoughVotes( PhotoID )

Dim sSubject
Dim sBody
Dim UserName
Dim EmailAddress
Dim Round3Start
Dim Round3End
Dim ProductName

set rsGetEmailInformation = GetEmailInformation(PhotoID)
if rsGetEmailInformation.RecordCount > 0 then
response.write("PhotoID= " & PhotoID)
UserName = rsGetEmailInformation("User")
EmailAddress = rsGetEmailInformation("Email")
Round3Start = FormatDateTime(rsGetEmailInformation("Round_3_Start"),1)
Round3End = FormatDateTime(rsGetEmailInformation("Round_3_End"),1)
ProductName = rsGetEmailInformation("Product_Name")

sFrom = "donotreply@membercalendars.com"
sSubject = "Sorry, your photo did not receive enough votes to advance"

sBody = "Dear " & UserName & ",<br><br>" & _
"Sorry, your photo did not receive enough votes to win a place in this year’s " & ProductName & _ 
" .<br><br>" & _
"We appreciate your choice to compete and proudly display your vehicle.  Please visit our storefront at membercalendars.com/store to see the finished product.<br><br>" & _
"Membercalendars.com encourages you to start shooting new photos now in preparation for next year’s calendar.  Utilize the change of seasons throughout the year, for a variety in the choices of backgrounds for your photos.<br><br>" & _
"If you have any feedback, please email info@membercalendars.com.<br><br>" & _
"Thanks again for choosing membercalendars.com, your source for customized calendars!"

SendEmail EmailAddress, UserName, sSubject, sBody

End if
End Sub

Sub clEmail_CalendarsReadyForPurchase(MemberName)

Dim sSubject 
Dim sBody

sSubject = "Membercalendars.com calendars are ready for purchase!"
sBody = "Dear " & UserName & "," & _
"Please visit membercalendars.com/store to see all available officially licensed calendar titles for sale.  With our broad selection of calendars, we’re sure you’ll find the perfect gift for your loved ones.  These calendars were made BY car enthusiasts FOR car enthusiasts.  We encourage you to share our passion 365 days a year!!!    " & _
"Thanks," & _
"Membercalendars.com"

SendEmail EmailAddress, UserName, sSubject, sBody

End Sub
%>