Using namespace System.Net
using namespace Microsoft.Xrm.Sdk
using namespace Microsoft.Xrm.Sdk.Query
using namespace Microsoft.Crm.Sdk.Messages
using namespace System.Collections.Generic


Import-Module .\test_mod.psm1

$token = Get-MGAuthToken

#Connect-MgGraph -AccessToken ($token.AccessToken | ConvertTo-SecureString -AsPlainText -Force)

$dvservice = Get-DataverseService

$dvmailboxes = Get-DataverseIncomingMailbox -service $dvservice

foreach($mb in $dvmailboxes) {

    if($null -ne $mb.userid) {
        
        ######## Process Incoming Emails ###########

        $messages = Get-ExchangeMessages -upn $mb.emailaddress
        foreach($our_message in $messages) {
            $userpref = Get-UserTrackingPreferences -UserId $mb.userid -service $dvservice
            if("all" -eq $userpref) {
                $emailid = Add-IncomingEmailInDataverse -service $dvservice -newemail $our_message
            }
            elseif("correlate" -eq $userpref) {
                $correlatedDvEmail = Get-CorrelatedEmail -service $dvservice -inreplyto ($our_message.inreplyto.Replace("<", "\u003C").Replace(">", "\u003E")) -newEmail $our_message
                if($null -ne $correlatedDvEmail -and $null -ne $correlatedDvEmail.id) {
                    $emailid = Add-IncomingEmailInDataverse -service $dvservice -newemail $our_message -correlatedemail $correlatedDvEmail
                }
            }
        }

        ######## Process Outgoing Emails ###########
        $emails = Get-DataverseOutgoingEmail -service $dvservice
        foreach($email in $emails) {
            Send-Email -token $token -message $email
        }
    }
}
