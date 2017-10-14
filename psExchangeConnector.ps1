<#
.SYNOPSIS
based on the smletsExchangeConnector V 1.2 this PS Script provides SCSM Exchange Connector functionality through PowerShell with native SCSM lets

.DESCRIPTION
This PowerShell script/runbook aims to address shortcomings and wants in the
out of box SCSM Exchange Connector as well as help enable new functionality around
Work Item scenarios, runbook automation, 3rd party customizations, and
enabling other organizational level processes via email

.NOTES 
Adjustments for native SCSM lets & some other small modifications: Roland Kind
Original Author: Adam Dzyacky

.Modifications based to original V1.2
* securestring password handling
* SR/IR prefix handling based on Customer SCSM definitions
* cleanup of some code for comment functions
* check if mail body contains some text 

#>


   
Set-StrictMode -Version 2

## how to convert password strings to secure string and create encrypted string from secure string
##
## $key= (12,34,45,78,90,101,123,222,13,16,32,123,242,154,33,233,1,34,212,72,62,51,135,143)
## $sc=ConvertTo-SecureString 'yourPlainTextPasswordString' -AsPlainText -Force 
## $st=ConvertFrom-SecureString -SecureString $sc -Key $key
##
   
# Test Environment Key    
$key= (12,34,45,78,90,101,123,222,13,16,32,123,242,154,33,233,1,34,212,72,62,51,135,143) # 192 bit key
    
# wf User (mxs)
$secString_wf="76492dANAAxADEANQAxADAAYwBlAGMANwAwADcANQA3AGQAYgBmAGYAMQBkAGIAOAA...... =" ##stringouput from ConvertFrom-SecureString ...
$password_wf = ConvertTo-SecureString -String $secString_wf -key $key
$username_wf = "<username>"
$un_wf = New-Object System.Management.Automation.PSCredential($username_wf,$password_wf)
   
# SCSM User - currently not in use maybe in the future if the workflow account (scorch) have no rights on scsm
#$secString_scsm= "76492d1.... "
#$password_scsm = ConvertTo-SecureString -String $secString_scsm -key $key
#$username_scsm = "<username>"    
#$un_scsm = New-Object System.Management.Automation.PSCredential($username_scsm,$password_scsm)




#region #### initialization

# SCSM2016 Core cmdLets & Core DLLs
import-Module 'C:\Program Files\Microsoft System Center\Service Manager\Powershell\System.Center.Service.Manager.psd1' 

[Reflection.Assembly]::LoadFile("C:\Program Files\Microsoft System Center\Service Manager\SDK\Microsoft.EnterpriseManagement.Core.dll") | Out-Null
[Reflection.Assembly]::LoadFile("C:\Program Files\Microsoft System Center\Service Manager\SDK\Microsoft.EnterpriseManagement.ServiceManager.dll") | Out-Null

#endregion


#region #### Configuration ####
#define the SCSM management server, this could be a remote name or localhost
$scsmMGMTServer = "<hostname>"

#define SCSM EnterpriseManagementGroupConnection and EnterpriseManagementGroup Object
$emgCS = New-object -TypeName "Microsoft.EnterpriseManagement.EnterpriseManagementConnectionSettings" "x"
$emgCS.ServerName = $scsmMGMTServer
$emg = New-Object -TypeName "Microsoft.EnterpriseManagement.EnterpriseManagementGroup" $emgCS

#define Exchange Server
$mxsServer = "<hostname>"

#define/use SCSM WF credentials
#$exchangeAuthenticationType - "windows" or "impersonation" are valid inputs here.
    #Windows will use the credentials that start this script in order to authenticate to Exchange and retrieve messages
        #choosing this option only requires the $workflowEmailAddress variable to be defined
        #this is ideal if you'll be using Task Manager or SMA to initiate this
    #Impersonation will use the credentials that are defined here to connect to Exchange and retrieve messages
        #choosing this option requires the $workflowEmailAddress, $username, $password, and $domain variables to be defined
$exchangeAuthenticationType = "impersonation"
$workflowEmailAddress = "<email>"
$username = $username_wf
$password = $password_wf
$domain = "<domain>"



#defaultNewWorkItem = set to either "ir" or "sr"
#minFileSizeInKB = Set the minimum file size in kilobytes to be attached to work items
#createUsersNotInCMDB = If someone from outside your org emails into SCSM this allows you to take that email and create a User in your CMDB
#includeWholeEmail = If long chains get forwarded into SCSM, you can choose to write the whole email to a single action log entry OR the beginning to the first finding of "From:"
#attachEmailToWorkItem = If $true, attach email as an *.eml to each work item. Additionally, write the Exchange Conversation ID into the Description of the Attachment object
#fromKeyword = If $includeWholeEmail is set to true, messages will be parsed UNTIL they find this word
$defaultNewWorkItem = "ir"
$defaultIRTemplate = Get-SCSMObjectTemplate -Name "DefaultIncidentTemplate" -computername $scsmMGMTServer
$defaultSRTemplate = Get-SCSMObjectTemplate -Name "ServiceManager.ServiceRequest.Library.Template.DefaultServiceRequest" -computername $scsmMGMTServer
$minFileSizeInKB = "25"
$createUsersNotInCMDB = $false 
$includeWholeEmail = $false
$attachEmailToWorkItem = $true 
$attachSingleAttachmentsToWorkItem = $false
$fromKeyword = "From"


#processCalendarAppointment = If $true, scheduling appointments with the Workflow Inbox where a [WorkItemID] is in the Subject will
    #set the Scheduled Start and End Dates on the Work Item per the Start/End Times of the calendar appointment
#mergeReplies = If $true, emails that are Replies (signified by RE: in the subject) will attempt to be matched to a Work Item in SCSM by their
    #Exchange Conversation ID and will also override $attachEmailToWorkItem to be $true if set to $false
$processCalendarAppointment = $false
$mergeReplies = $false

#optional, enable KB search of your Cireson HTML KB
#this uses the now depricated Cireson KB API Search by Text, it works as of v7.x but should be noted it could be entirely removed in future portals
#$ciresonPortalServer = URL that will be used to search for KB articles via invoke-webrequest. Make sure to leave the "/" after your tld!
#$ciresonAccountGUID = This is the GUID/ID of a User within the ServiceManagement DB that will be used to search the knowledge base from the CI$User table
#$ciresonPortalWindowsAuth = how invoke-webrequest should attempt to auth to your portal server.
    #Leave true if your portal uses Windows Auth, change to False for Forms authentication.
    #If using forms, you'll need to set the ciresonPortalUsername and Password variables. For ease, you could set this equal to the username/password defined above
$searchCiresonHTMLKB = $false
$ciresonPortalServer = "https://portalserver.domain.tld/"
$ciresonAccountGUID = "11111111-2222-3333-4444-AABBCCDDEEFF"
$ciresonKBLanguageCode = "ENU"
$ciresonPortalWindowsAuth = $true
$ciresonPortalUsername = ""
$ciresonPortalPassword = ""

#define SCSM Work Item keywords to be used
$acknowledgedKeyword = "Acknowledged"
$reactivateKeyword = "Reactivate"
$resolvedKeyword = "Resolved"
$closedKeyword = "Closed"
$holdKeyword = "Hold"
$cancelledKeyword = "Cancelled"
$takeKeyword = "Take"
$completedKeyword = "Completed"
$skipKeyword = "Skipped"
$approvedKeyword = "Approved"
$rejectedKeyword = "Rejected"

#define the path to the Exchange Web Services API. the following is the default install directory for EWS API
$exchangeEWSAPIPath = "C:\Program Files\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll"

#endregion


##
## Helper Funktions to replace SMLets
##
function rk.get-scsmenumeration ($name)
{
 $ret = $null;
 $enumEntities= $emg.EntityTypes.GetEnumerations()
 foreach ($enumEntity in $enumEntities)
  {
   if ($enumEntity.name -eq $name)
    {$ret=$enumEntity}
   if ($enumEntity.displayname -eq $name)
    {$ret=$enumEntity}

  }
  return $ret
}

function rk.get-userSMTPAddress ($user)
{
 $UserEmail=$null
 $SCSRelUserObjects = Get-SCSMRelationshipInstance -SourceInstance $user -computername $scsmMGMTServer

        foreach ($SCSRelUserObject in $SCSRelUserObjects)
        {
              $smtpFound=$false
              foreach ($smtpprop in $SCSRelUserObject.targetObject.Values)
              {
      
                if ($smtpprop.Type.Name -eq 'ChannelName' -and $smtpprop.Value -eq 'SMTP')
                {
                    $smtpFound=$true            
                }


                if ($smtpprop.Type.Name -eq 'TargetAddress' -and $smtpFound -eq $true)
                {
                    $UserEmail=$smtpprop            
                }
              }  
        }
 return $UserEmail
}

##  Set-SCSMObject -SMObject $workItem -Property FirstResponseDate -Value $message.DateTimeSent.ToUniversalTime() -computername $scsmMGMTServer}
function rk.Set-SCSMObject {
param (
        [parameter(Mandatory=$True,Position=0)]$SMObject,
        [parameter(Mandatory=$True,Position=1)]$Property,
        [parameter(Mandatory=$True,Position=2)]$Value
        
    )
    $SMobject."$Property"=$Value;
    $xx=Update-SCSMClassInstance $SMObject -PassThru -computername $scsmMGMTServer
}

##   Set-SCSMObject -SMObject $reviewer -PropertyHashtable @{"Decision" = "DecisionEnum.Approved$"; "DecisionDate" = $message.DateTimeSent.ToUniversalTime(); "Comments" = $commentToAdd} 
function rk.Set-SCSMObjectHT {
param (
        [parameter(Mandatory=$True,Position=0)]$SMObject,
        [parameter(Mandatory=$True,Position=1)]$PropertyHashTable
      )
    
    foreach ($key in $PropertyHashTable.keys)
    {
        $SMobject."$key"=$PropertyHashTable["$key"]
    
    }
    
    $xx=Update-SCSMClassInstance $SMObject -PassThru -computername $scsmMGMTServer
}


#region #### SCSM Classes ####
$irClass = get-scsmclass -name "System.WorkItem.Incident" -computername $scsmMGMTServer
$srClass = get-scsmclass -name "System.WorkItem.ServiceRequest" -computername $scsmMGMTServer
$prClass = get-scsmclass -name "System.WorkItem.Problem" -computername $scsmMGMTServer
$crClass = get-scsmclass -name "System.Workitem.ChangeRequest" -computername $scsmMGMTServer
$rrClass = get-scsmclass -name "System.Workitem.ReleaseRecord" -computername $scsmMGMTServer
$maClass = get-scsmclass -name "System.WorkItem.Activity.ManualActivity" -computername $scsmMGMTServer
$raClass = get-scsmclass -name "System.WorkItem.Activity.ReviewActivity" -computername $scsmMGMTServer
$paClass = get-scsmclass -name "System.WorkItem.Activity.ParallelActivity" -computername $scsmMGMTServer
$saClass = get-scsmclass -name "System.WorkItem.Activity.SequentialActivity" -computername $scsmMGMTServer
$daClass = get-scsmclass -name "System.WorkItem.Activity.DependentActivity" -computername $scsmMGMTServer

$raHasReviewerRelClass = Get-SCSMRelationship -name "System.ReviewActivityHasReviewer" -computername $scsmMGMTServer
$raReviewerIsUserRelClass = Get-SCSMRelationship -name "System.ReviewerIsUser" -computername $scsmMGMTServer
$raVotedByUserRelClass = Get-SCSMRelationship -name "System.ReviewerVotedByUser" -computername $scsmMGMTServer

$userClass = get-scsmclass -name "System.User" -computername $scsmMGMTServer
$domainUserClass = get-scsmclass -name "System.Domain.User" -computername $scsmMGMTServer
$notificationClass = get-scsmclass -name "System.Notification.Endpoint" -computername $scsmMGMTServer

$irLowImpact = rk.Get-SCSMEnumeration "System.WorkItem.TroubleTicket.ImpactEnum.Low" 
#$irLowImpact = "8f1a713e-53aa-9d8a-31b9-a9540074f305"

$irLowUrgency = rk.Get-SCSMEnumeration "System.WorkItem.TroubleTicket.UrgencyEnum.Low" 
#$irLowUrgency = "725a4cad-088c-4f55-a845-000db8872e01"

$irActiveStatus = rk.Get-SCSMEnumeration "IncidentStatusEnum.Active" 
#$irActiveStatus = "5e2d3932-ca6d-1515-7310-6f58584df73e" 

$affectedUserRelClass       = get-scsmrelationship -name "System.WorkItemAffectedUser" -computername $scsmMGMTServer
$assignedToUserRelClass     = Get-SCSMRelationship -name "System.WorkItemAssignedToUser" -computername $scsmMGMTServer
$createdByUserRelClass      = Get-SCSMRelationship -name "System.WorkItemCreatedByUser" -computername $scsmMGMTServer
$workResolvedByUserRelClass = Get-SCSMRelationship -name "System.WorkItem.TroubleTicketResolvedByUser" -computername $scsmMGMTServer
$wiRelatesToCIRelClass      = Get-SCSMRelationship -name "System.WorkItemRelatesToConfigItem" -computername $scsmMGMTServer
$wiRelatesToWIRelClass      = Get-SCSMRelationship -name "System.WorkItemRelatesToWorkItem" -computername $scsmMGMTServer
$wiContainsActivityRelClass = Get-SCSMRelationship -name "System.WorkItemContainsActivity" -computername $scsmMGMTServer
$sysUserHasPrefRelClass     = Get-SCSMRelationship -name "System.UserHasPreference" -ComputerName $scsmMGMTServer

$fileAttachmentClass       = Get-SCSMClass -Name "System.FileAttachment" -computername $scsmMGMTServer
$fileAttachmentRelClass    = Get-SCSMRelationship -name "System.WorkItemHasFileAttachment" -computername $scsmMGMTServer
$fileAddedByUserRelClass   = Get-SCSMRelationship -name "System.FileAttachmentAddedByUser" -ComputerName $scsmMGMTServer


#$irTypeProjection = Get-SCSMTypeProjection -name "system.workitem.incident.projectiontype" -computername $scsmMGMTServer
#$srTypeProjection = Get-SCSMTypeProjection -name "system.workitem.servicerequestprojection" -computername $scsmMGMTServer
#$userHasPrefProjection = Get-SCSMTypeProjection -name  "System.User.Preferences.Projection" -computername $scsmMGMTServer
 
$ttAnalystCommentLogRel= Get-SCSMRelationship -name "System.WorkItem.TroubleTicketHasAnalystComment" -ComputerName $scsmMGMTServer
$ttUserCommentLogRel = Get-SCSMRelationship -name "System.WorkItem.TroubleTicketHasUserComment" -ComputerName $scsmMGMTServer
$ttAnalystCommentLogClass = get-scsmclass -name "System.WorkItem.TroubleTicket.AnalystCommentLog" -ComputerName $scsmMGMTServer
$ttUserCommentLogClass = get-scsmclass -name "System.WorkItem.TroubleTicket.UserCommentLog" -ComputerName $scsmMGMTServer


#endregion


#region #### Exchange Connector Functions ####
function New-WorkItem ($message, $wiType, $returnWIBool) 
{
    $from = $message.From
    $to = $message.To
    $cced = $message.CC
    $title = $message.subject
    $description = $message.body

    #if the message is longer than 4000 characters take only the first 4000.
    if ($description.length -ge "4000")
    {
        $description = $description.substring(0,4000)
    }

    #find Affected User from the From Address
    $relatedUsers = @()
    $userSMTPNotification = Get-SCSMclassinstance -Class $notificationClass -Filter "TargetAddress -eq $from"  -computername $scsmMGMTServer
    if ($userSMTPNotification) 
    { 
        $affectedUser = get-scsmclassinstance -id (Get-SCSMRelationshipinstance -Target $userSMTPNotification -computername $scsmMGMTServer).sourceObject.id -computername $scsmMGMTServer
    }
    else
    {
        if ($createUsersNotInCMDB -eq $true)
        {
            $affectedUser = create-userincmdb $from
        }
    }

    #find Related Users (To)       
    # if ($to -gt $0)
    if ($to -ne "")
    {
        $x = 0
        while ($x -lt $to.count)
        {
            $ToSMTP = $to[$x]
            $userToSMTPNotification = Get-SCSMclassinstance -Class $notificationClass -Filter "TargetAddress -eq $($ToSMTP.address)"  -computername $scsmMGMTServer
            if ($userToSMTPNotification) 
            { 
                $relatedUser = (Get-SCRelationshipInstance -Target $userToSMTPNotification -computername $scsmMGMTServer).sourceObject 
                $relatedUsers += $relatedUser
            }
            else
            {
                if ($createUsersNotInCMDB -eq $true)
                {
                    $newUser = create-userincmdb $to[$x]
                    $relatedUsers += $relatedUser
                }
            }
            $x++
        }
    }
    
    #find Related Users (Cc)         
    # if ($cced -gt $0)
    if ($cced -ne "")
    {
        $x = 0
        while ($x -lt $cced.count)
        {
            $ccSMTP = $cced[$x]
            $userCCSMTPNotification = Get-SCSMclassinstance -Class $notificationClass -Filter "TargetAddress -eq $($ccSMTP.address)" -computername $scsmMGMTServer
            if ($userCCSMTPNotification) 
            { 
                $relatedUser = (Get-SCSMRelationshipInstance -Target $userCCSMTPNotification -computername $scsmMGMTServer).sourceObject 
                $relatedUsers += $relatedUser
            }
            else
            {
                if ($createUsersNotInCMDB -eq $true)
                {
                    $newUser = create-userincmdb $cced[$x]
                    $relatedUsers += $relatedUser
                }
            }
            $x++
        }
    }

    #create the Work Item based on the globally defined Work Item type and Template
    switch ($defaultNewWorkItem) 
    {
        "ir" {
                    $prefix = (Get-SCSMClassInstance -computername $scsmMGMTServer -class (Get-SCSMClass -Name "System.WorkItem.Incident.GeneralSetting" -computername $scsmMGMTServer)).PrefixForId 
                    $newWorkItem = New-SCSMclassinstance -Class $irClass -Property @{Id = $prefix+"{0}"; Status = $irActiveStatus; Title = $title; Description = $description; Classification = $null; Impact = $irLowImpact; Urgency = $irLowUrgency; Source=(rk.get-scsmenumeration "IncidentSourceEnum.Email").id} -PassThru -computername $scsmMGMTServer
                    
                    # $irProjection = Get-SCSMObjectProjection -ProjectionName $irTypeProjection.Name -Filter "Name -eq $($newWorkItem.Name)" -computername $scsmMGMTServer
                                    
                    if ($attachSingleAttachmentsToWorkItem -eq $true) {if($message.Attachments){Attach-FileToWorkItem $message $newWorkItem.ID}}
                    if ($attachEmailToWorkItem -eq $true){Attach-EmailToWorkItem $message $newWorkItem.ID}

                    # Set-SCSMObjectTemplate -Projection $irProjection -Template $defaultIRTemplate -computername $scsmMGMTServer
                    
                    if ($affectedUser)
                    {
                        New-SCRelationshipinstance -Relationship $createdByUserRelClass -Source $newWorkItem -Target $affectedUser -computername $scsmMGMTServer
                        New-SCRelationshipinstance -Relationship $affectedUserRelClass -Source $newWorkItem -Target $affectedUser  -computername $scsmMGMTServer
                    }
                    if ($relatedUsers)
                    {
                        foreach ($relatedUser in $relatedUsers)
                        {
                            New-SCRelationshipinstance -Relationship $wiRelatesToCIRelClass -Source $newWorkItem -Target $relatedUser -computername $scsmMGMTServer
                        }
                    }
                    if ($searchCiresonHTMLKB -eq $true)
                    {
                        Search-CiresonKnowledgeBase $message $newWorkItem
                    }
                }
        "sr" {
                    $prefix = (Get-SCSMClassInstance -class (Get-SCSMClass -Name System.GlobalSetting.ServiceRequestSettings)).ServiceRequestPrefix
                    $newWorkItem = New-SCSMClassInstance -Class $srClass -Property @{Id = $prefix+"{0}"; Title = $title; Description = $description; Status= (rk.get-scsmenumeration "ServiceRequestStatusEnum.Submitted").id} -PassThru -computername $scsmMGMTServer
                   
                    # $srProjection = Get-SCSMObjectProjection -ProjectionName $srTypeProjection.Name -Filter "Name -eq $($newWorkItem.Name)" -computername $scsmMGMTServer
                   
                    if ($attachSingleAttachmentsToWorkItem -eq $true) {if($message.Attachments){Attach-FileToWorkItem $message $newWorkItem.ID}}
                    if ($attachEmailToWorkItem -eq $true){Attach-EmailToWorkItem $message $newWorkItem.ID}
                   
                    # Set-SCSMObjectTemplate -projection $srProjection -Template $defaultSRTemplate -computername $scsmMGMTServer
                   
                    if ($affectedUser)
                    {
                        New-SCRelationshipInstance -Relationship $createdByUserRelClass -Source $newWorkItem -Target $affectedUser -computername $scsmMGMTServer
                        New-SCRelationshipInstance -Relationship $affectedUserRelClass -Source $newWorkItem -Target $affectedUser -computername $scsmMGMTServer
                    }
                    if ($relatedUsers)
                    {
                        foreach ($relatedUser in $relatedUsers)
                        {
                            New-SCRelationshipInstance -Relationship $wiRelatesToCIRelClass -Source $newWorkItem -Target $relatedUser -computername $scsmMGMTServer
                        }
                    }
                    if ($searchCiresonHTMLKB -eq $true)
                    {
                        Search-CiresonKnowledgeBase $message $newWorkItem
                    }
                } 
    }

    if ($returnWIBool -eq $true)
    {
        return $newWorkItem
    }
}

function Update-WorkItem ($message, $wiType, $workItemID) 
{

    # check if message.body contains some text ...

    if ($message.body)
    {
        #determine the comment to add and ensure it's less than 4000 characters
        if ($includeWholeEmail -eq $true)
        {
            $commentToAdd = $message.body
            if ($commentToAdd.length -ge "4000")
            {
                $commentToAdd.substring(0, 4000)
            }
        }
        else
        {
            $fromKeywordPosition = $message.Body.IndexOf("$fromKeyword" + ":")
            if (($fromKeywordPosition -eq $null) -or ($fromKeywordPosition -eq -1))
            {
                $commentToAdd = $message.body
                if ($commentToAdd.length -ge "4000")
                {
                    $commentToAdd.substring(0, 4000)
                }
            }
            else
            {
                $commentToAdd = $message.Body.substring(0, $fromKeywordPosition)
                if ($commentToAdd.length -ge "4000")
                {
                    $commentToAdd.substring(0, 4000)
                }
            }
        }
    }
    #determine who left the comment
    $userSMTPNotification = Get-SCSMClassInstance -Class $notificationClass -Filter "TargetAddress -eq $($message.From)" -computername $scsmMGMTServer
    if ($userSMTPNotification) 
    { 
        # $commentLeftBy = get-scsmclassInstance -id (Get-SCSMRelationshipInstance -TargetInstance $userSMTPNotification -computername $scsmMGMTServer).sourceObject.id -computername $scsmMGMTServer
        $commentLeftBy = get-scsmuser -id (Get-SCSMRelationshipInstance -TargetInstance $userSMTPNotification -computername $scsmMGMTServer).sourceObject.id -computername $scsmMGMTServer
    }
    else
    {
        if ($createUsersNotInCMDB -eq $true)
        {
            $commentLeftBy = create-userincmdb $from
        }
    }

    #add any attachments
    if ($message.Attachments)
    {
        Attach-FileToWorkItem $message $workItemID
    }

    #update the work item with the comment and/or action
    switch ($wiType) 
    {
        #### primary work item types ####
        "ir" {
                    $workItem = get-scsmclassinstance -class $irClass -filter "Name -eq $workItemID" -computername $scsmMGMTServer
                    
                    try {
                            $affectedUser = get-scsmclassinstance -computername $scsmMGMTServer -id (Get-SCSMRelationshipinstance -sourceInstance $workItem -computername $scsmMGMTServer | Where-Object{$_.relationshipid -eq $affectedUserRelClass.id}).targetobject.id 
                        } catch {$affectedUser=$null}
                    if($affectedUser){$affectedUserSMTP =rk.get-userSMTPAddress $affectedUser}

                    try {
                            $assignedTo = get-scsmclassinstance -computername $scsmMGMTServer -id (Get-SCSMRelationshipinstance -sourceInstance $workItem -computername $scsmMGMTServer | Where-Object{$_.relationshipid -eq $assignedToUserRelClass.id}).targetobject.id 
                        } catch {$assignedTo=$null}
                    if($assignedTo){$assignedToSMTP =rk.get-userSMTPAddress $assignedTo} else {$assignedToSMTP = @{Value = $null}}
                    
                    #write to the Action log only if body contains some text
                    if ($message.body)
                    {
                        switch ($message.From)
                        {
                            $affectedUserSMTP.Value {rk.Add-WIComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -AnalystComment $false -isPrivate $false}
                            $assignedToSMTP.Value {if($commentToAdd -match "#private"){$isPrivateBool = $true}else{$isPrivateBool = $false};rk.Add-WIComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -AnalystComment $true -isPrivate $isPrivateBool}
                            default {if($commentToAdd -match "#private"){$isPrivateBool = $true}else{$isPrivateBool = $null};rk.Add-WIComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $true -isPrivate $isPrivateBool}
                        }
                    
                    
                        #take action on the Work Item if neccesary
                        switch -Regex ($commentToAdd)
                        {
                            "\[$acknowledgedKeyword]" {if ($workItem.FirstResponseDate -eq $null){rk.Set-SCSMObject -SMObject $workItem -Property FirstResponseDate -Value $message.DateTimeSent.ToUniversalTime() }}
                            "\[$resolvedKeyword]" {New-SCRelationshipInstance -Relationshipclass $workResolvedByUserRelClass -Source $workItem -Target $commentLeftBy -computername $scsmMGMTServer;rk.Set-SCSMObject -SMObject $workItem -Property Status -Value "IncidentStatusEnum.Resolved"; }
                            "\[$closedKeyword]" {rk.Set-SCSMObject -SMObject $workItem -Property Status -Value "IncidentStatusEnum.Closed"}
                            "\[$takeKeyword]" {New-SCRelationshipinstance -Relationshipclass $assignedToUserRelClass -Source $workItem -Target $commentLeftBy -computername $scsmMGMTServer }
                            "\[$reactivateKeyword]" {if ($workItem.Status.Name -eq "IncidentStatusEnum.Resolved") {rk.Set-SCSMObject -SMObject $workItem -Property Status -Value "IncidentStatusEnum.Active" }}
                            "\[$reactivateKeyword]" {if ($workItem.Status.Name -eq "IncidentStatusEnum.Closed") 
                                                        {if($message.Subject -match "[I][R][0-9]+")
                                                            {$message.subject = $message.Subject.Replace("[" + $Matches[0] + "]", "")};
                                                            $returnedWorkItem = New-WorkItem $message "ir" $true; 
                                                            try{New-SCRelationshipinstance -Relationshipclass $wiRelatesToWIRelClass -Source $workItem -Target $returnedWorkItem -computername $scsmMGMTServer}
                                                                catch
                                                                {}}
                                                                }
                        }


                    }
                    #relate the user to the work item
                    New-SCRelationshipinstance -Relationshipclass $wiRelatesToCIRelClass -Source $workItem -Target $commentLeftBy -computername $scsmMGMTServer
                    
                    #add any new attachments
                    if ($attachEmailToWorkItem -eq $true){Attach-EmailToWorkItem $message $WorkItem.ID}
                    
                    
                } 
        "sr" {
                    $workItem = get-scsmclassinstance -class $srClass -filter "Name -eq $workItemID" -computername $scsmMGMTServer                    

                    try {
                            $affectedUser = get-scsmclassinstance -id (Get-SCSMRelationshipinstance -sourceInstance $workItem | Where-Object{$_.relationshipid -eq $affectedUserRelClass.id}).targetobject.id -computername $scsmMGMTServer
                        } catch {}
                    if($affectedUser){$affectedUserSMTP =rk.get-userSMTPAddress $affectedUser}

                    try {
                            $assignedTo = get-scsmclassinstance -id (Get-SCSMRelationshipinstance -sourceInstance $workItem | Where-Object{$_.relationshipid -eq $assignedToUserRelClass.id}).targetobject.id -computername $scsmMGMTServer                            
                        } catch {$assignedTo=$null}
                    if($assignedTo){$assignedToSMTP =rk.get-userSMTPAddress $assignedTo} else {$assignedToSMTP = @{Value = $null}}
                    
                    #write to the Action log
                    if ($message.body)
                    {
                        switch ($message.From)
                        {
                            $affectedUserSMTP.Value {rk.Add-WIComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -AnalystComment $false -isPrivate $false}
                            $assignedToSMTP.Value {if($commentToAdd -match "#private"){$isPrivateBool = $true}else{$isPrivateBool = $false};rk.Add-WIComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -AnalystComment $true -isPrivate $isPrivateBool}
                            default {if($commentToAdd -match "#private"){$isPrivateBool = $true}else{$isPrivateBool = $null};rk.Add-WIComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $true -isPrivate $isPrivateBool}
                        }
                   
                   
                        switch -Regex ($commentToAdd)
                        {
                            "\[$completedKeyword]" {rk.Set-SCSMObject -SMObject $workItem -Property Status -Value "ServiceRequestStatusEnum.Completed" }
                            "\[$cancelledKeyword]" {rk.Set-SCSMObject -SMObject $workItem -Property Status -Value "ServiceRequestStatusEnum.Cancelled" }
                            "\[$closedKeyword]"    {rk.Set-SCSMObject -SMObject $workItem -Property Status -Value "ServiceRequestStatusEnum.Closed" }
                        }
                    }
                    
                    #relate the user to the work item
                    New-SCRelationshipInstalnce -Relationshipclass $wiRelatesToCIRelClass -Source $workItem -Target $commentLeftBy -computername $scsmMGMTServer
                    
                    #add any new attachments
                    if ($attachEmailToWorkItem -eq $true){Attach-EmailToWorkItem $message $newWorkItem.ID}
                } 
        "pr" {
                    $workItem = get-scsmclassinstance -class $prClass -filter "Name -eq $workItemID" -computername $scsmMGMTServer                    
                    
                    try {
                            $assignedTo = get-scsmclassinstance -id (Get-SCSMRelationshipinstance -sourceInstance $workItem | Where-Object{$_.relationshipid -eq $assignedToUserRelClass.id}).targetobject.id -computername $scsmMGMTServer                            
                        } catch {$assignedTo=$null}
                    if($assignedTo){$assignedToSMTP =rk.get-userSMTPAddress $assignedTo} else {$assignedToSMTP = @{Value = $null}}
                    
                    #write to the Action log
                    if ($message.body)
                    {
                        switch ($message.From)
                        {
                            $assignedToSMTP.Value {rk.Add-WIComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -AnalystComment $true -isPrivate $false}
                            default {rk.Add-WIComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $true -isPrivate $null}
                        }
                    
                        #take action on the Work Item if neccesary
                        switch -Regex ($commentToAdd)
                        {
                             "\[$resolvedKeyword]" {rk.Set-SCSMObject -SMObject $workItem -Property Status -Value "ProblemStatusEnum.Resolved" ; New-SCRelationshipIntance -Relationshipclass $workResolvedByUserRelClass -Source $workItem -Target $commentLeftBy -computername $scsmMGMTServer }
                             "\[$closedKeyword]" {rk.Set-SCSMObject -SMObject $workItem -Property Status -Value "ProblemStatusEnum.Closed" }
                             "\[$takeKeyword]" {New-SCRelationshipInstance -relationshipclass $assignedToUserRelClass -Source $workItem -Target $commentLeftBy -computername $scsmMGMTServer }
                             "\[$reactivateKeyword]" {if ($workItem.Status.Name -eq "ProblemStatusEnum.Resolved") 
                                                        {rk.Set-SCSMObject -SMObject $workItem -Property Status -Value "ProblemStatusEnum.Active$" }
                                                     }
                        }
                    }

                    #relate the user to the work item
                    New-SCRelationshipInstance -Relationshipclass $wiRelatesToCIRelClass -Source $workItem -Target $commentLeftBy -computername $scsmMGMTServer
                    
                    #add any new attachments
                    if ($attachEmailToWorkItem -eq $true){Attach-EmailToWorkItem $message $newWorkItem.ID}
                }
        "cr" {
                    $workItem = get-scsmclassinstance -class $crClass -filter "Name -eq $workItemID" -computername $scsmMGMTServer                    
                    
                    try {
                            $assignedTo = get-scsmclassinstance -id (Get-SCSMRelationshipinstance -sourceInstance $workItem | Where-Object{$_.relationshipid -eq $assignedToUserRelClass.id}).targetobject.id -computername $scsmMGMTServer
                        } catch {$assignedTo=$null}
                    if($assignedTo){$assignedToSMTP =rk.get-userSMTPAddress $assignedTo} else {$assignedToSMTP = @{Value = $null}}
                    
                    #write to the Action log
                    if ($message.body)
                    {
                        switch ($message.From)
                        {
                            $assignedToSMTP.Value {rk.Add-WIComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -AnalystComment $true -isPrivate $false}
                            default {rk.Add-WIComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $true -isPrivate $null}
                        }
                    
                        #take action on the Work Item if neccesary
                        switch -Regex ($commentToAdd)
                        {
                            "\[$holdKeyword]" {rk.Set-SCSMObject -SMObject $workItem -Property Status -Value "ChangeStatusEnum.OnHold" }
                            "\[$cancelledKeyword]" {rk.Set-SCSMObject -SMObject $workItem -Property Status -Value "ChangeStatusEnum.Cancelled" }
                            "\[$takeKeyword]" {New-SCRelationshipInstance -relationshipclass $assignedToUserRelClass -Source $workItem -Target $commentLeftBy -computername $scsmMGMTServer }
                        }
                    }
                    #relate the user to the work item
                    New-SCRelationshipInstance -Relationshipclass $wiRelatesToCIRelClass -Source $workItem -Target $commentLeftBy -computername $scsmMGMTServer
                    
                    #add any new attachments
                    if ($attachEmailToWorkItem -eq $true){Attach-EmailToWorkItem $message $newWorkItem.ID}
                }
        
        #### activities ####
        "ra" {
                    $workItem = get-scsmclassinstance -class $raClass -filter "Name -eq $workItemID" -computername $scsmMGMTServer                    
                    $reviewers = Get-SCSMRelationshipinstance -sourceinstance $workItem -computername $scsmMGMTServer | where-object {$_.relationshipid -eq $raHasReviewerRelClass.id} 
                     
                    if ($message.body)
                    {
                        if ($commentToAdd.length -ge "50")
                        {
                            $commentToAdd = $commentToAdd.substring(0,50)
                        }

                        foreach ($reviewer in $reviewers)
                        {
                            $reviewingUser = Get-SCSMClassInstance -id (Get-SCSMRelationshipInstance -SourceInstance (get-scsmclassinstance -id $reviewer.targetobject.id)).targetobject.id                        
                            $reviewingUserSMTP = rk.get-userSMTPAddress $reviewingUser
                        
                            #approved
                            if (($reviewingUserSMTP.Value -eq $message.From) -and ($commentToAdd -match "\[$approvedKeyword]"))
                            {
                                $reviewObject=Get-SCSMClassInstance -id $reviewer.targetobject.id
                                rk.Set-SCSMObjectHT -SMObject $reviewObject -PropertyHashtable @{"Decision" = "DecisionEnum.Approved"; "DecisionDate" = $message.DateTimeSent.ToUniversalTime(); "Comments" = $commentToAdd} 
                                New-SCRelationshipInstance -Relationshipclass $raVotedByUserRelClass -Source $reviewObject -Target $reviewingUser -computername $scsmMGMTServer
                            }
                            #rejected
                            elseif (($reviewingUserSMTP.Value -eq $message.From) -and ($commentToAdd -match "\[$rejectedKeyword]"))
                            {
                                $reviewObject=Get-SCSMClassInstance -id $reviewer.targetobject.id
                                rk.Set-SCSMObjectHT -SMObject $reviewObject -PropertyHashtable @{"Decision" = "DecisionEnum.Rejected"; "DecisionDate" = $message.DateTimeSent.ToUniversalTime(); "Comments" = $commentToAdd} 
                                New-SCRelationshipInstance -Relationshipclass $raVotedByUserRelClass -Source $reviewObject -Target $reviewingUser -computername $scsmMGMTServer
                            }
                            #no keyword, add a comment to parent work item
                            elseif (($reviewingUserSMTP.Value -eq $message.From) -and (($commentToAdd -notmatch "\[$approvedKeyword]") -or ($commentToAdd -notmatch "\[$rejectedKeyword]")))
                            {
                                $parentWorkItem = Get-SCSMWorkItemParent $workItem.Get_Id().Guid
                                switch ($parentWorkItem.Classname)
                                {
                                    "System.WorkItem.ChangeRequest"  {rk.Add-WIComment -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $false -IsPrivate $false}
                                    "System.WorkItem.ServiceRequest" {rk.Add-WIComment -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $false -IsPrivate $false}
                                    "System.WorkItem.Incident"       {rk.Add-WIComment -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $false -IsPrivate $false}
                                }
                            
                            }
                        }
                    }
                }
        "ma" {
                    $workItem = get-scsmclassinstance -class $maClass -filter "Name -eq $workItemID" -computername $scsmMGMTServer  
                    if ($message.body)
                    {
                        if ($commentToAdd.length -ge "50")
                        {
                            $commentToAdd = $commentToAdd.substring(0,50)
                        }
                    
                        try {$activityImplementer = get-scsmclassinstance -id (Get-SCSMRelationshipinstance -Sourceinstance $workItem -computername $scsmMGMTServer | where-object {$_.Relationshipid -eq $assignedToUserRelClass.id}).targetobject.id} catch {}
                        if ($activityImplementer){$activityImplementerSMTP = rk.get-userSMTPAddress $activityImplementer}
                    
                        #completed
                        if (($activityImplementerSMTP.Value -eq $message.From) -and ($commentToAdd -match "\[$completedKeyword]"))
                        {
                            rk.Set-SCSMObjectHT -SMObject $workItem -PropertyHashtable @{"Status" = "ActivityStatusEnum.Completed"; "ActualEndDate" = (get-date).ToUniversalTime(); "Notes" = $commentToAdd } 
                        }
                        #skipped
                        elseif (($activityImplementerSMTP.Value -eq $message.From) -and ($commentToAdd -match "\[$skipKeyword]"))
                        {
                            rk.Set-SCSMObjectHT -SMObject $workItem -PropertyHashtable @{"Status" = "ActivityStatusEnum.Skipped"; "ActualEndDate" = (get-date).ToUniversalTime(); "Notes" = "$($workItem.Notes)$($activityImplementer.Name) @ $(get-date): $commentToAdd `n"} 
                        }
                        #not from the Activity Implementer, add to the MA Notes
                        elseif (($activityImplementerSMTP.value -ne $message.From))
                        {
                            rk.Set-SCSMObjectHT -SMObject $workItem -PropertyHashtable @{"Notes" = "$($workItem.Notes)$($activityImplementer.Name) @ $(get-date): $commentToAdd `n"} 
                        }                    
                        #no keywords, add to the Parent Work Item
                        elseif (($activityImplementerSMTP.Value -eq $message.From) -and (($commentToAdd -notmatch "\[$completedKeyword]") -or ($commentToAdd -notmatch "\[$skipKeyword]")))
                        {
                            $parentWorkItem = Get-SCSMWorkItemParent $workItem.Get_Id().Guid
                            switch ($parentWorkItem.Classname)
                            {
                                "System.WorkItem.ChangeRequest"  {rk.Add-WIComment -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $false -IsPrivate $false}
                                "System.WorkItem.ServiceRequest" {rk.Add-WIComment -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $false -IsPrivate $false}
                                "System.WorkItem.Incident"       {rk.Add-WIComment -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $false -IsPrivate $false}
                            }
                            
                        }
                    }
                }
    } 
}

function Attach-EmailToWorkItem ($message, $workItemID)
{
    $messageMime = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($exchangeService,$message.id,$mimeContentSchema)
    $MemoryStream = New-Object System.IO.MemoryStream($messageMime.MimeContent.Content,0,$messageMime.MimeContent.Content.Length)    
    
    #convert for native scsm cmdlets
    [system.io.stream] $SysStream = $MemoryStream;
   
    $fileAttachmentProperties = @{
        Id = [Guid]::NewGuid().ToString()
        DisplayName = "message.eml"
        Description = "ExchangeConversationID:$($message.ConversationID)"
        Extension =   "eml"
        Size=        $MemoryStream.Length
        AddedDate=   [DateTime]::Now.ToUniversalTime()
        Content =     $SysStream
    };

    $workItemInstance = Get-SCSMClassInstance -name $workItemID -computername $scsmMGMTServer
    $emailAttachment=new-screlationshipinstance -relationshipclass $fileAttachmentRelClass -Source $workItemInstance -TargetClass $fileAttachmentClass -TargetProperty $fileAttachmentProperties -computername $scsmMGMTServer -PassThru
    $emailAttachment1=Get-SCSMclassInstance -id $emailAttachment.targetobject.id -computername $scsmMGMTServer
        
    #create the Attached By relationship if possible
    $userSMTPNotification = Get-SCSMclassinstance -Class $notificationClass -Filter "TargetAddress -eq $($message.from)" -computername $scsmMGMTServer
    if ($userSMTPNotification) 
    { 
        $attachedByUser = get-scsmclassinstance -id (Get-SCSMRelationshipinstance -Target $userSMTPNotification -computername $scsmMGMTServer).sourceObject.id -computername $scsmMGMTServer
        New-SCRelationshipinstance -Source $emailAttachment1 -Relationshipclass $fileAddedByUserRelClass -Target $attachedByUser -computername $scsmMGMTServer
    }
}

#inspired and modified from Stefan Roth here - https://stefanroth.net/2015/03/28/scsm-passing-attachments-via-web-service-e-g-sma-web-service/
function Attach-FileToWorkItem ($message, $workItemId)
{
    foreach ($attachment in $message.Attachments)
    {
        $attachment.Load()
        $base64attachment = [System.Convert]::ToBase64String($attachment.Content)

        #Convert the Base64String back to bytes
        $AttachmentContent = [convert]::FromBase64String($base64attachment)

        #Create a new MemoryStream object out of the attachment data
        $MemoryStream = New-Object System.IO.MemoryStream($AttachmentContent,0,$AttachmentContent.length)
       
        #convert for native scsm cmdlets
        [system.io.stream] $SysStream = $MemoryStream;

        if (([int]$MemoryStream.Length * 1000) -gt $minFileSizeInKB)
        {
            $fileAttachmentProperties = @{
                Id = [Guid]::NewGuid().ToString()
                DisplayName = "message.eml"
                Description = "ExchangeConversationID:$($message.ConversationID)"
                Extension =   "eml"
                Size=        $MemoryStream.Length
                AddedDate=   [DateTime]::Now.ToUniversalTime()
                Content =     $SysStream
            };

            $workItemInstance = Get-SCSMClassInstance -name $workItemID -computername $scsmMGMTServer
            $emailAttachment=new-screlationshipinstance -relationshipclass $fileAttachmentRelClass -Source $workItemInstance -TargetClass $fileAttachmentClass -TargetProperty $fileAttachmentProperties -PassThru -computername $scsmMGMTServer
            $emailAttachment1=Get-SCSMclassInstance -id $emailAttachment.targetobject.id -computername $scsmMGMTServer
            
            #create the Attached By relationship if possible
            $userSMTPNotification = Get-SCSMclassinstance -Class $notificationClass -Filter "TargetAddress -eq $($message.from)" -computername $scsmMGMTServer
            if ($userSMTPNotification) 
            { 
                $attachedByUser = get-scsmclassinstance -id (Get-SCSMRelationshipinstance -Target $userSMTPNotification -computername $scsmMGMTServer).sourceObject.id -computername $scsmMGMTServer
                New-SCRelationshipinstance -Source $emailAttachment1 -Relationshipclass $fileAddedByUserRelClass -Target $attachedByUser -computername $scsmMGMTServer
            }
        }
    }
}

function Get-WorkItem ($workItemID, $workItemClass)
{
    #removes [] surrounding a Work Item ID if neccesary
    if ($workitemID.StartsWith("[") -and $workitemID.EndsWith("]"))
    {
        $workitemID = $workitemID.TrimStart("[").TrimEnd("]")
    }

    #get the work item
    $wi = get-scsmclassinstance -Class $workItemClass -Filter "Name -eq $workItemID" -computername $scsmMGMTServer
    return $wi
}

#courtesy of Leigh Kilday. Modified.
function Get-SCSMWorkItemParent
{
    [CmdLetBinding()]
    PARAM (
        [Parameter(ParameterSetName = 'GUID', Mandatory=$True)]
        [Alias('ID')]
        $WorkItemGUID
    )
    PROCESS
    {
        TRY
        {
            If ($PSBoundParameters['WorkItemGUID'])
            {
                Write-Verbose -Message "[PROCESS] Retrieving WI with GUID"
                $ActivityObject = Get-SCSMObject -Id $WorkItemGUID -computername $scsmMGMTServer
            }
        
            #Retrieve Parent
            Write-Verbose -Message "[PROCESS] Activity: $($ActivityObject.Name)"
            Write-Verbose -Message "[PROCESS] Retrieving WI Parent"
            $ParentRelatedObject = Get-SCSMRelationshipObject -ByTarget $ActivityObject -computername $scsmMGMTServer | ?{$_.RelationshipID -eq $wiContainsActivityRelClass.id.Guid}
            $ParentObject = $ParentRelatedObject.SourceObject

            Write-Verbose -Message "[PROCESS] Activity: $($ActivityObject.Name) - Parent: $($ParentObject.Name)"

            If ($ParentObject.ClassName -eq 'System.WorkItem.ServiceRequest' `
            -or $ParentObject.ClassName -eq 'System.WorkItem.ChangeRequest' `
            -or $ParentObject.ClassName -eq 'System.WorkItem.ReleaseRecord' `
            -or $ParentObject.ClassName -eq 'System.WorkItem.Incident' `
            -or $ParentObject.ClassName -eq 'System.WorkItem.Problem')
            {
                Write-Verbose -Message "[PROCESS] This is the top level parent"
                
                #return parent object Work Item
                Return $ParentObject
            }
            Else
            {
                Write-Verbose -Message "[PROCESS] Not the top level parent. Running against this object"
                Get-SCSMWorkItemParent -WorkItemGUID $ParentObject.Id.GUID -computername $scsmMGMTServer
            }
        }
        CATCH
        {
            Write-Error -Message $Error[0].Exception.Message
        }
    }
}

#inspired and modified from Travis Wright here - https://blogs.technet.microsoft.com/servicemanager/2013/01/16/creating-membership-and-hosting-objectsrelationships-using-new-scsmobjectprojection-in-smlets/
function Create-UserInCMDB ($userEmail)
{
    #The ID for external users appears to be a GUID, but it can't be identified by get-scsmobject
    #The ID for internal domain users takes the form of domain_username_SMTP
    #It's unclear how this ID should be generated. Opted to take the form of an internal domain for the ID
    #By using the internal domain style (_SMTP) this means New/Update Work Item tasks will understand how to find these new external users going forward
    $username = $userEmail.Split("@")[0]
    $domainAndTLD = $userEmail.Split("@")[1]
    $domain = $domainAndTLD.Split(".")[0]
    $newID = $domain + "_" + $username + "_SMTP"

    #create the new user
    $newUser = New-SCSMclassinstance -computername $scsmMGMTServer -Class $domainUserClass -Property @{"domain" = "$domainAndTLD"; "username" = "$username"; "displayname" = "$userEmail"} -PassThru

    $UserProperties = @{
                Id = $NewID
                TargetAddress = $userEmail
                DisplayName = "e-Mail Address"
                ChannelName = "SMTP"
            };
    
    #create the user's email notification channel
    $UserRelObject=new-screlationshipinstance -computername $scsmMGMTServer -relationshipclass (Get-SCSMRelationship -name system.userhaspreference) -Source $newUser -TargetClass $notificationClass -TargetProperty $UserProperties -PassThru

    return $newUser
}


function rk.Add-WIComment {
param (
        [parameter(Mandatory=$True,Position=0)]$WIObject,
        [parameter(Mandatory=$True,Position=1)]$Comment,
        [parameter(Mandatory=$True,Position=2)]$EnteredBy,
        [parameter(Mandatory=$False,Position=3)]$AnalystComment,
        [parameter(Mandatory=$False,Position=4)]$IsPrivate
    )

    # Set thread current culture - maybe not neccessary ..
    [threading.thread]::CurrentThread.CurrentCulture = 'en-US'
    
    $CommentProperties = @{}
   
    # Generate a new GUID for the comment
    $NewGUID = ([guid]::NewGuid()).ToString()

    If ($AnalystComment -eq $true) {
            $CommentClass = $ttAnalystCommentLogClass
            $commentRel = $ttAnalystCommentLogRel
            $CommentClassName = "AnalystComments"
            $CommentProperties = @{
                Id = $NewGUID
                DisplayName = $NewGUID
                Comment = $Comment
                EnteredBy = $EnteredBy.DisplayName.tostring()        
                EnteredDate=   [DateTime]::Now.ToUniversalTime()
                IsPrivate = $IsPrivate
            };


        } else {
            $CommentClass = $ttUserCommentLogClass
            $commentRel = $ttUserCommentLogRel
            $CommentClassName = "UserComments"
            $CommentProperties = @{
                Id = $NewGUID
                DisplayName = $NewGUID
                Comment = $Comment
                EnteredBy = $EnteredBy.DisplayName.tostring()        
                EnteredDate=   [DateTime]::Now.ToUniversalTime()
            };
        }


    $CommentObject=new-screlationshipinstance -relationshipclass $commentRel -Source $WIObject -TargetClass $commentClass -TargetProperty $CommentProperties -PassThru -computername $scsmMGMTServer
    
}

#inspired and modified from Travis Wright here - https://blogs.technet.microsoft.com/servicemanager/2013/01/16/creating-membership-and-hosting-objectsrelationships-using-new-scsmobjectprojection-in-smlets/

#search the Cireson KB based on content from a New Work Item and notify the Affected User
function Search-CiresonKnowledgeBase ($message, $workItem)
{
    $searchQuery = $workItem.Title.Trim() + " " + $workItem.Description.Trim()
    $resultsToNotify = @()

    if ($ciresonPortalWindowsAuth -eq $true)
    {
        $portalLoginRequest = Invoke-WebRequest -Uri $ciresonPortalServer -Method get -UseDefaultCredentials -SessionVariable ecPortalSession
        $kbResults = Invoke-WebRequest -Uri ($ciresonPortalServer + "api/V3/KnowledgeBase/GetHTMLArticlesFullTextSearch?userId=$ciresonAccountGUID&searchValue=$searchQuery&isManager=false&userLanguageCode=$ciresonKBLanguageCode") -WebSession $ecPortalSession
    }
    else
    {
        $portalLoginRequest = Invoke-WebRequest -Uri $ciresonPortalServer -Method get -SessionVariable ecPortalSession
        $loginForm = $portalLoginRequest.Forms[0]
        $loginForm.Fields["UserName"] = $ciresonPortalUsername
        $loginForm.Fields["Password"] = $ciresonPortalPassword
    
        $portalLoginPost = Invoke-WebRequest -Uri ($ciresonPortalServer + "/Login/Login?ReturnUrl=%2f") -WebSession $ecPortalSession -Method post -Body $loginForm.Fields
        $kbResults = Invoke-WebRequest -Uri ($ciresonPortalServer + "api/V3/KnowledgeBase/GetHTMLArticlesFullTextSearch?userId=$ciresonAccountGUID&searchValue=$searchQuery&isManager=false&userLanguageCode=$ciresonKBLanguageCode") -WebSession $ecPortalSession
    }

    $kbResults = $kbResults.Content | ConvertFrom-Json
    $kbResults =  $kbResults | ?{$_.endusercontent -ne ""} | select-object articleid, title
    
    if ($kbResults)
    {
        foreach ($kbResult in $kbResults)
        {
            $resultsToNotify += "<a href=$ciresonPortalServer" + "KnowledgeBase/View/$($kbResult.articleid)#/>$($kbResult.title)</a><br />"
        }

        #build the message and send through already established EWS connection
        #determine if IR/SR to build appropriate email response
        if ($workItem.ClassName -eq "System.WorkItem.Incident")
        {
            $resolveMailTo = "<a href=`"mailto:$email" + "?subject=" + "[" + $workItem.id + "]" + ";body=This%20can%20be%20[resolved]" + "`">resolve</a>"
        }
        else
        {
            $resolveMailTo = "<a href=`"mailto:$email" + "?subject=" + "[" + $workItem.id + "]" + ";body=This%20can%20be%20[cancelled]" + "`">cancel</a>"
        }

        $body = "We found some knowledge articles that may be of assistance to you <br/><br/>
            $resultsToNotify <br/><br />
            If one of these articles resolves your request, you can use the following
            link to $resolveMailTo your request."
    
        #add the Work Item ID, so a potential reply doesn't trigger the creation of a new work item but instead updates it
        Send-EmailFromWorkflowAccount -subject ("[" + $workItem.id + "]" - $message.subject) -body $body -bodyType "HTML" -toRecipients $message.From
    }
}

#send an email from the SCSM Workflow Account
function Send-EmailFromWorkflowAccount ($subject, $body, $bodyType, $toRecipients)
{
    $emailToSendOut = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $exchangeService
    $emailToSendOut.Subject = $subject
    $emailToSendOut.Body = $body
    $emailToSendOut.ToRecipients.Add($toRecipients)
    $emailToSendOut.Body.BodyType = $bodyType
    $emailToSendOut.Send()
}

function Schedule-WorkItem ($calAppt, $wiType, $workItem)
{
    #set the Scheduled Start/End dates on the Work Item
    $scheduledHashTable =  @{"ScheduledStartDate" = $calAppt.StartTime.ToUniversalTime(); "ScheduledEndDate" = $calAppt.EndTime.ToUniversalTime()}  
    switch ($wiType)
    {
        "ir" {rk.Set-SCSMObjectHT -SMObject $workItem -propertyhashtable $scheduledHashTable}
        "sr" {rk.Set-SCSMObjectHT -SMObject $workItem -propertyhashtable $scheduledHashTable}
        "pr" {rk.Set-SCSMObjectHT -SMObject $workItem -propertyhashtable $scheduledHashTable}
        "cr" {rk.Set-SCSMObjectHT -SMObject $workItem -propertyhashtable $scheduledHashTable}
        "rr" {rk.Set-SCSMObjectHT -SMObject $workItem -propertyhashtable $scheduledHashTable}

        #activities
        "ma" {rk.Set-SCSMObjectHT -SMObject $workItem -propertyhashtable $scheduledHashTable}
        "pa" {rk.Set-SCSMObjectHT -SMObject $workItem -propertyhashtable $scheduledHashTable}
        "sa" {rk.Set-SCSMObjectHT -SMObject $workItem -propertyhashtable $scheduledHashTable}
        "da" {rk.Set-SCSMObjectHT -SMObject $workItem -propertyhashtable $scheduledHashTable}
    }

    #Trigger Update to update the Action log of the item
    Update-WorkItem -message $calAppt -wiType $wiType -workItemID $workItem.id
}

function Verify-WorkItem ($message)
{
    #If emails are being attached to New Work Items, filter on the File Attachment Description that equals the Exchange Conversation ID as defined in the Attach-EmailToWorkItem function
    if ($attachEmailToWorkItem -eq $true)
    {
        $emailAttachmentSearchObject = Get-SCSMObject -Class $fileAttachmentClass -Filter "Description -eq 'ExchangeConversationID:$($message.ConversationID);'" -ComputerName $scsmMGMTServer | select-object -first 1 
        $relatedWorkItemFromAttachmentSearch = Get-SCSMObject -Id (Get-SCSMRelationshipObject -ByTarget $emailAttachmentSearchObject -ComputerName $scsmMGMTServer).sourceobject.id -ComputerName $scsmMGMTServer
        if ($emailAttachmentSearchObject -and $relatedWorkItemFromAttachmentSearch)
        {
            switch ($relatedWorkItemFromAttachmentSearch.ClassName)
            {
                "System.WorkItem.Incident" {Update-WorkItem -message $message -wiType "ir" -workItemID $relatedWorkItemFromAttachmentSearch.id}
                "System.WorkItem.ServiceRequest" {Update-WorkItem -message $message -wiType "sr" -workItemID $relatedWorkItemFromAttachmentSearch.id}
            }
        }
        else
        {
            #no match was found, Create a New Work Item
            New-WorkItem $message $defaultNewWorkItem
        }
    }
    else
    {
        #will never engage as Verify-WorkItem currently only works when attaching emails to work items 
    }
}
#endregion






#determine merge logic
if ($mergeReplies -eq $true)
{
    $attachEmailToWorkItem = $true
}

#define Exchange assembly and connect to EWS
[void] [Reflection.Assembly]::LoadFile("$exchangeEWSAPIPath")
$exchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService

switch ($exchangeAuthenticationType)
{
    "impersonation" {$exchangeService.Credentials = New-Object Net.NetworkCredential($username_wf, $password_wf, $domain)}
    "windows" {$exchangeService.UseDefaultCredentials = $true}
}



# $exchangeService.AutodiscoverUrl($workflowEmailAddress)
# or alternate hardcoded Server URL
$exchangeService.Url ="https://"+$mxsServer+"/EWS/Exchange.asmx"

##Login to Mailbox with Impersonation
$exchangeService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$workflowEmailAddress
) 


#define search parameters and search on the defined classes
$inboxFolderName = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox
$inboxFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService,$inboxFolderName)
$itemView = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList 1000
$propertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$propertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text
$mimeContentSchema = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.ItemSchema]::MimeContent)
$dateTimeItem = [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived
$now = get-date
$searchFilter = New-Object -TypeName Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo -ArgumentList $dateTimeItem,$now

#build the Where-Object scriptblock based on defined configuration
$emailFilterString = '($_.ItemClass -eq "IPM.Note")'
$calendarFilterString = '($_.ItemClass -eq "IPM.Schedule.Meeting.Request")'
$unreadFilterString = '($_.isRead -eq $false)'
$inboxFilterString = $emailFilterString
if ($processCalendarAppointment -eq $true)
{
    $inboxFilterString = $emailFilterString + " -or " + $calendarFilterString
}

#finalize the where-object string by ensuring to look for all Unread Items
$inboxFilterString = "(" + $inboxFilterString + ")" + " -and " + $unreadFilterString
$inboxFilterString = [scriptblock]::Create("$inboxFilterString")

#filter the inbox
$inbox = $exchangeService.FindItems($inboxFolder.Id,$searchFilter,$itemView) | where-object $inboxFilterString


#parse each message
foreach ($message in $inbox)
{
    #load the entire message
    $message.Load($propertySet)
   

    #Process an Email
    if ($message.ItemClass -eq "IPM.Note")
    {
        $email = New-Object System.Object 
        $email | Add-Member -type NoteProperty -name From -value $message.From.Address
        $email | Add-Member -type NoteProperty -name To -value $message.ToRecipients
        $email | Add-Member -type NoteProperty -name CC -value $message.CcRecipients
        $email | Add-Member -type NoteProperty -name Subject -value $message.Subject
        $email | Add-Member -type NoteProperty -name Attachments -value $message.Attachments
        $email | Add-Member -type NoteProperty -name Body -value $message.Body.Text
        $email | Add-Member -type NoteProperty -name DateTimeSent -Value $message.DateTimeSent
        $email | Add-Member -type NoteProperty -name DateTimeReceived -Value $message.DateTimeReceived
        $email | Add-Member -type NoteProperty -name ID -Value $message.ID
        $email | Add-Member -type NoteProperty -name ConversationID -Value $message.ConversationID
        $email | Add-Member -type NoteProperty -name ConversationTopic -Value $message.ConversationTopic

        switch -Regex ($email.subject) 
        { 
            #### primary work item types ####
            "\[[I][R][0-9]+\]" {$result = get-workitem $matches[0] $irClass; if ($result){update-workitem $email "ir" $result.id} else {new-workitem $email $defaultNewWorkItem}}
            "\[[S][R][0-9]+\]" {$result = get-workitem $matches[0] $srClass; if ($result){update-workitem $email "sr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
            "\[[P][R][0-9]+\]" {$result = get-workitem $matches[0] $prClass; if ($result){update-workitem $email "pr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
            "\[[C][R][0-9]+\]" {$result = get-workitem $matches[0] $crClass; if ($result){update-workitem $email "cr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
 
            #### activities ####
            "\[[R][A][0-9]+\]" {$result = get-workitem $matches[0] $raClass; if ($result){update-workitem $email "ra" $result.id}}
            "\[[M][A][0-9]+\]" {$result = get-workitem $matches[0] $maClass; if ($result){update-workitem $email "ma" $result.id}}

            #### 3rd party classes, work items, etc. add here ####

            #### Email is a Reply and does not contain a [Work Item ID]
            # Check if Work Item (Title, Body, Sender, CC, etc.) exists
            # and the user was replying too fast to receive Work Item ID notification
            "([R][E][:])(?!.*\[(([I|S|P|C][R])|([M|R][A]))[0-9]+\])(.+)" {if($mergeReplies -eq $true){Verify-WorkItem $email} else{new-workitem $email $defaultNewWorkItem}}

            #### default action, create work item ####
            default {new-workitem $email $defaultNewWorkItem} 
        }

      
        #mark the message as read on Exchange, move to deleted items
        ##        
        ##
        $message.IsRead = $true
        $hideInVar01 = $message.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve)
        $hideInVar02 = $message.Move([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems)
    }

    #Process a Calendar Appointment
    elseif ($message.ItemClass -eq "IPM.Schedule.Meeting.Request")
    {
        $appointment = New-Object System.Object 
        $appointment | Add-Member -type NoteProperty -name StartTime -value $message.Start
        $appointment | Add-Member -type NoteProperty -name EndTime -value $message.End
        $appointment | Add-Member -type NoteProperty -name To -value $message.ToRecipients
        $appointment | Add-Member -type NoteProperty -name From -value $message.From.Address
        $appointment | Add-Member -type NoteProperty -name Attachments -value $message.Attachments
        $appointment | Add-Member -type NoteProperty -name Subject -value $message.Subject
        $appointment | Add-Member -type NoteProperty -name DateTimeReceived -Value $message.DateTimeReceived
        $appointment | Add-Member -type NoteProperty -name DateTimeSent -Value $message.DateTimeSent
        $appointment | Add-Member -type NoteProperty -name Body -value $message.Body.Text
        $appointment | Add-Member -type NoteProperty -name ID -Value $message.ID
        $appointment | Add-Member -type NoteProperty -name ConversationID -Value $message.ConversationID
        $appointment | Add-Member -type NoteProperty -name ConversationTopic -Value $message.ConversationTopic

        switch -Regex ($appointment.subject) 
        { 
            #### primary work item types ####
            "\[[I][R][0-9]+\]" {$result = get-workitem $matches[0] $irClass; if ($result){schedule-workitem $appointment "ir" $result; $message.Accept($true)}}
            "\[[S][R][0-9]+\]" {$result = get-workitem $matches[0] $srClass; if ($result){schedule-workitem $appointment "sr" $result; $message.Accept($true)}}
            "\[[P][R][0-9]+\]" {$result = get-workitem $matches[0] $prClass; if ($result){schedule-workitem $appointment "pr" $result; $message.Accept($true)}}
            "\[[C][R][0-9]+\]" {$result = get-workitem $matches[0] $crClass; if ($result){schedule-workitem $appointment "cr" $result; $message.Accept($true)}}
            "\[[R][R][0-9]+\]" {$result = get-workitem $matches[0] $rrClass; if ($result){schedule-workitem $appointment "rr" $result; $message.Accept($true)}}

            #### activities ####
            "\[[M][A][0-9]+\]" {$result = get-workitem $matches[0] $maClass; if ($result){schedule-workitem $appointment "ma" $result; $message.Accept($true)}}
            "\[[P][A][0-9]+\]" {$result = get-workitem $matches[0] $paClass; if ($result){schedule-workitem $appointment "pa" $result; $message.Accept($true)}}
            "\[[S][A][0-9]+\]" {$result = get-workitem $matches[0] $saClass; if ($result){schedule-workitem $appointment "sa" $result; $message.Accept($true)}}
            "\[[D][A][0-9]+\]" {$result = get-workitem $matches[0] $daClass; if ($result){schedule-workitem $appointment "da" $result; $message.Accept($true)}}

            #### 3rd party classes, work items, etc. add here ####

            #### default action, create/schedule a new default work item ####
            default {$returnedNewWorkItemToSchedule = new-workitem $appointment $defaultNewWorkItem $true; schedule-workitem -calAppt $appointment -wiType $defaultNewWorkItem -workItem $returnedNewWorkItemToSchedule; $message.Accept($true)} 
        }
    }
}


