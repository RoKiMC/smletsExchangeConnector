//Optional - Hide "Send Outlook Meeting" from End Users
//app.custom.formTasks.add('Incident', null, function (formObj, viewModel) { formObj.boundReady(function () 
//{ if (!session.user.Analyst) { $( ".taskmenu li:contains('Send Outlook Meeting')" ).hide() } }); }); 

//Email address and Display name (that you'll see in the CC: of the email address) of your SCSM Workflow Account to receive the meeting
var scsmWFAccount = "account@domain.tld"
var scsmWFAccountDisplayName = "Service Manager"

//need to make these dynamic
var stateDate = "20171031T050000Z"
var endDate = "20171031T050000Z"
var icsCreateDate = "20171031T050000Z"

//Send Outlook Meeting, Incident
app.custom.formTasks.add('Incident', "Send Outlook Meeting", function (formObj, viewModel){ 
	console.log(formObj);
	
	var meetingSubject = "[" + pageForm.viewModel.Id + "]" + " " + pageForm.viewModel.Title
	var calMeetingInvitee = pageForm.viewModel.RequestedWorkItem.UPN
	
	//Create calendar meeting based on Affected User
	if (calMeetingInvitee)
	{
		var calMeetingInviteeDisplayName = pageForm.viewModel.RequestedWorkItem.DisplayName
		var icsMSG = "BEGIN:VCALENDAR\nVERSION:2.0\nMETHOD:PUBLISH\nBEGIN:VEVENT\nATTENDEE;CN="+ calMeetingInviteeDisplayName +";RSVP=TRUE:mailto:" + calMeetingInvitee + "\nATTENDEE;CN="+ scsmWFAccountDisplayName +";RSVP=TRUE:mailto:" + scsmWFAccount + "\nDTSTART:" + stateDate + "+\nDTEND:" + endDate + "\nTRANSP:OPAQUE\nSEQUENCE:0\nUID:" + pageForm.viewModel.Id + "\nDTSTAMP:" + icsCreateDate + "\nSUMMARY:" + meetingSubject + "\nPRIORITY:5\nCLASS:PUBLIC\nBEGIN:VALARM\nTRIGGER:PT15M\nACTION:DISPLAY\nDESCRIPTION:Reminder\nEND:VALARM\nEND:VEVENT\nEND:VCALENDAR"
	}
	else
	{
		var icsMSG = "BEGIN:VCALENDAR\nVERSION:2.0\nMETHOD:PUBLISH\nBEGIN:VEVENT\nATTENDEE;CN=" + scsmWFAccountDisplayName + ";RSVP=TRUE:mailto:" + scsmWFAccount +"\nDTSTART:" + stateDate + "\nDTEND:" + endDate + "\nTRANSP:OPAQUE\nSEQUENCE:0\nUID:" + pageForm.viewModel.Id +"\nDTSTAMP:" + icsCreateDate + "\nSUMMARY:"+ meetingSubject +"\nPRIORITY:5\nCLASS:PUBLIC\nBEGIN:VALARM\nTRIGGER:PT15M\nACTION:DISPLAY\nDESCRIPTION:Reminder\nEND:VALARM\nEND:VEVENT\nEND:VCALENDAR"
	}
	
	//Open the the ICS file
	window.open( "data:text/calendar;charset=utf8," + escape(icsMSG));
});