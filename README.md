# SCSM Exchange Connector via SMlets
This PowerShell script leverages the [SMlets module](https://www.powershellgallery.com/packages/smlets/0.5.0.1) to build an open and flexible Exchange Connector for controlling Microsoft System Center Service Manager 2012+


## So what is this for?
The stock Exchange Connector is a seperate download for SCSM 2012+ that enables SCSM deployments to leverage an Exchange mailbox to process updates to work items. While incredibly useful, some feel limited by its inability to be customized given its nature as a sealed management pack. This PowerShell script replicates all functionality of [Exchange Connector 3.1](https://www.microsoft.com/en-ca/download/details.aspx?id=45291) introduces several new features, and most importantly enables SCSM Administrators to customize the solution to their needs.

## Who is this for?
This is aimed at SCSM administrators looking to further push the automation limits of what their SCSM deployment can do with inbound email processing. As such, you should be comfortable with PowerShell and navigating SCSM via SMlets.

## What new things can it do?
Digitally Signed and Encrypted Emails (v1.3)
- Leveraging the open source [MimeKit](https://github.com/jstedfast/MimeKit) project by Jeffrey Stedfast, the connector can now process digitally signed or encrypted emails just like regular mail. This requires an appropriate certificate in either the user's personal cert store or the local machine's personal cert store.

System Center Operations Manager (SCOM) Integration (v1.3)
- Using a configurable [keyword], authorized users (as defined individually or through an Active Directory group) can request the overall Health and Alert counts of a Distributed Application.

Merge Replies from Related Users instead of Creating New Default Work Items (v1.2)
- If a user emails the SCSM Workflow Account and also adds additional users to the To/CC lines those related users are automatically added to the Related Items tab of a New Work Item. However in these scenarios, it's possible that one of those users could reply within the same processing loop of the Exchange Connector. As a result, they will queue more emails to be turned into New Default Work Items. This feature aims to address the scenario by querying Exchange Inbox/Deleted Items for matching Conversation Topics and ConversationIDs, finding the original item in the thread, searching for the Work Item that already exists in SCSM, and then appending their Reply to its Action Log

Schedule Work Items (v1.2)
- It's now possible to interact with SCSM via Outlook Calendar Appointments! When a Calendar Appointment is sent, the Scheduled Start Date and Scheduled End Date will be set on the Work Item based on the start/end times of the appointment. If the work item cannot be found/does not exist, a new default work item is created and it's scheduled start/end tiems set accordingly. Upon success, the appointment will be accepted onto the workflow account's calendar and the requester will receive confirmation of the booking. This introduces the possibility of leveraging the workflow's calendar as a central place to see all Scheduled Work Items. Using [Cireson's Outlook Plugin](http://cireson.com/apps/outlook-console/)? When setting a Work Item reminder in Outlook, you can now CC your workflow account to update values on the Work Item.

Minimum File Attachment Size
- You can set a minimum size in KB. In doing so, files less than the defined size will not be added to the work item (i.e. corporate signature graphics won't be added)

File Attachment "Added by"
- When an email is sent with attachments, the "File Attachment Added By User" relationship will be set based on the Sender if the user is found in the CMDB

Incident, Service Request, Change Request, Problem
- [Take] - When emailing your workflow account, it will assign the Incident, Service Request, Change Request, or Problem to you (from address) when this keyword is featured in the body of the email.

Incident
- [Reactivate] - When submitted to a Resolved Incident, it will be reactivated. When submitted to a Closed Incident, a New Incident will be created and the two related to one another.

Change Request
- [Hold] - Place the Change Request On-Hold when this keyword is featured in the body of the email
- [Cancel] - Cancel the Change Request when this keyword is featured in the body of the email

Manual Activity
- [Skipped] - Skip the activity when this keyword is featured in the body of the email
- Misc - Anyone who is not the implementer will have their email appended to the "Notes" area of the MA
- Misc - If the Implementer leaves a comment that is not [Skipped] or [Completed] the comment is added to the highest level Parent Work Item

Review Activity
- Any reviewer who leaves a comment that doesn't contain [Approved] or [Rejected] will have their comment added to the highest level Parent Work Item. This addresses a scenario where users not familiar with SCSM (i.e. departments outside of IT) respond back to the email thinking someone is reading the message on your workflow account. Now their comments aren't simply lost, but instead given the visibility they deserve!

Incident and Service Request
- #private - When the message is attached to the action log, it will be marked as private if #private is featured in the body of the message.

Assigned To/Affected User relationships on the Action Log
- When someone who isn't the Assigned To/Affected User leaves a comment on the Action Log the comment's "IsPrivate" flag is marked as null (this is a bug in the EC v3.0 and v3.1 that has yet to be addressed by Microsoft). As such Cireson's Action Log Notify has no qualifier to go of off. With this script, the same functionality is present but now can be altered to get in line with SCSM and Cireson's MP.

Search Cireson HTML Knowledge Base
- If you're a customer of Cireson and this feature is enabled, your respective Cireson Portal HTML KB will be searched when a New Work item is generated using its title and description. The Sender will be sent a summarized HTML email with links directly to those knowledge articles about their recently created Work Item using the Exchange EWS API defined therein. As an example email, I've included an email body that features a [Resolved] and [Cancelled] link should the Affected User wish to mark their Incident/Service Request accordingly in the event the KB addresses their request. It should be noted, this is using the Cireson Web API to get KB through a now deprecated function. While this works, it goes without saying if Cireson drops this in coming versions it would cease to work. It has been tested and confirmed working with v7.x of their SCSM HTML KB portal.
