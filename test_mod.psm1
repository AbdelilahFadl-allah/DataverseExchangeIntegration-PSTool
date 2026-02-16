Using namespace System.Net
using namespace Microsoft.Xrm.Sdk
using namespace Microsoft.Xrm.Sdk.Query
using namespace Microsoft.Crm.Sdk.Messages
using namespace Microsoft.Xrm.Sdk.Messages
using namespace System.Collections.Generic

Import-Module .\LoadPackages.psm1
Load-Packages

Function Get-DataverseService {
    param()
    begin {
        $connStr = @"
        <dataverse_connection_string>
"@
        $service = Get-CrmConnection -ConnectionString $connStr
        Write-Output $service
    }
}

Function Get-DataverseIncomingMailbox {
    param(
        [Parameter(Mandatory = $true)]
        [psobject]
        $service
    )
    begin {
        $fetch = @"
            <fetch>
                <entity name="mailbox">
                    <attribute name="emailaddress" />
                    <attribute name="lastmessageid" />
                    <attribute name="mailboxid" />
                    <filter>
                        <condition attribute="incomingemaildeliverymethod" operator="eq" value="2" />
                        <condition attribute="statecode" operator="eq" value="0" />
                        <condition attribute="emailaddress" operator="not-null" />
                    </filter>
                    <link-entity name="queue" from="queueid" to="regardingobjectid" link-type="outer" alias="queue">
                        <attribute name="queueid" />
                    </link-entity>
                    <link-entity name="systemuser" from="systemuserid" to="regardingobjectid" link-type="outer" alias="user">
                        <attribute name="systemuserid" />
                        <filter>
                            <condition attribute="islicensed" operator="eq" value="true" />
                        </filter>
                    </link-entity>
                </entity>
                </fetch>
"@
        $result = $service.GetEntityDataByFetchSearchEC($fetch)
        $mailboxes = @()
        foreach ($email in $result.Entities) {
            $outo = @{
                emailaddress = [string]$email.Attributes["emailaddress"]
                mailboxid    = [guid]$email.Id
                userid       = if ($email.Attributes.Contains("user.systemuserid")) { [guid]([AliasedValue]$email.Attributes["user.systemuserid"]).Value }else { $null }
                queueid      = if ($email.Attributes.Contains("queue.queueid")) { [guid]([AliasedValue]$email.Attributes["queue.queueid"]).Value }else { $null }
            }
            
            $mb = New-Object psobject -Property $outo
            $mailboxes += $mb
        }

        Write-Output $mailboxes
    }
}

Function Get-EmailTrackingConfiguration {
    param(
        [Parameter(Mandatory = $true)]
        [psobject]
        $service
    )
    begin {
        $fetch = @"
            <fetch>
                <entity name="organization">
                    <attribute name="emailcorrelationenabled" />
                    <attribute name="trackingprefix" />
                    <attribute name="emailconnectionchannel" />
                    <filter>
                        <condition attribute="emailconnectionchannel" operator="eq" value="0" />
                    </filter>
                </entity>
            </fetch>
"@
        $result = $service.GetEntityDataByFetchSearchEC($fetch)
        
        if (($null -ne $result) -and ($null -ne $result.Entities) -and ($result.Entities.Count -eq 1)) {
            $organization = $result.Entities[0]
            #manage trackingtoken prefixes
            $prefixes = if ($null -ne $organization.Attributes["trackingprefix"]) { $organization.Attributes["trackingprefix"].Split(";") } else { $null }
            $filtered_prefixes = if ($null -ne $prefixes) { $prefixes | Where-Object { $true -ne [string]::IsNullOrEmpty($_) } } else { $null }
            $props = @{}
            $props["emailcorrelationenabled"] = $organization.Attributes["emailcorrelationenabled"]
            $props["trackingprefix"] = $filtered_prefixes
            $props["emailconnectionchannel"] = ([OptionSetValue]$organization.Attributes["emailconnectionchannel"]).Value
            $output_org = New-Object psobject -Property $props
            Write-Output $output_org
        }
        else {
            Write-Output $null
        }
    }
    process {}
    end {}
}

Function Get-MGAuthToken {
    param ()

    begin {
        <#
        $connectionDetails = @{ 
            'TenantId'     = '<tenant_id>' 
            'ClientId'     = '<app_id>' 
            'ClientSecret' = '<app_secret>' | ConvertTo-SecureString -AsPlainText -Force 
        }
        #>
        $tenantID = '<tenant_id>'
        $tokenBody = @{
            Grant_Type    = "client_credentials"
            Scope         = "https://graph.microsoft.com/.default"
            Client_Id     = '<app_id>'
            Client_Secret = '<app_secret>'
         }
         $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token" -Method POST -Body $tokenBody
        
        Write-Output $tokenResponse
    }
}

#Retrieves incoming messages
Function Get-ExchangeMessages {
    Param(
        [Parameter(Mandatory = $true)]
        [string]
        $upn
    )
    begin {
        try {
            #Get Mail Folder
            $mailFolderUrl = "https://graph.microsoft.com/v1.0/users/$($upn)/mailFolders/Inbox"
            $token = Get-MGAuthToken
            $headers = @{Authorization = ("Bearer " + $token.access_token) }
            $inboxFolder = Invoke-RestMethod -Method Get -Uri $mailFolderUrl -Headers $headers -ErrorAction Continue
            if ($null -eq $inboxFolder) {
                Write-Output @()
            }
            Connect-MgGraph -AccessToken ($token.access_token | ConvertTo-SecureString -AsPlainText -Force)
            $messages = Get-MgUserMailFolderMessage -MailFolderId $inboxFolder.Id `
                -UserId $upn `
                -Filter "not(categories/any(c:c eq 'trackedindataverse'))" `
                -Top 100 `
                -Property "internetMessageHeaders", "body", "torecipients", "from", "ccrecipients", "subject", "internetmessageid", "importance", "hasattachments", "attachments" `
                -ExpandProperty Attachments
            $out_messages = @()
            foreach ($msg in $messages) {
                $props = @{}
                $props["messageid"] = $msg.internetmessageid
                $props["inreplyto"] = $msg.InternetMessageHeaders | Where-Object { $_.Name -ieq "In-Reply-To" } | Select-Object -Property Value -ExpandProperty Value
                $props["subject"] = $msg.Subject
                $props["body"] = $msg.Body.Content
                $props["importance"] = $msg.Importance
                $props["sender"] = $msg.From.EmailAddress.Address
                $props["torecipients"] = $msg.ToRecipients | Where-Object { $null -ne $_.EmailAddress } | Select-Object -ExpandProperty EmailAddress | Select-Object -ExpandProperty Address
                $props["ccrecipients"] = $msg.CcRecipients | Where-Object { $null -ne $_.EmailAddress } | Select-Object -ExpandProperty EmailAddress | Select-Object -ExpandProperty Address
                $props["bccrecipients"] = $msg.BccRecipients | Where-Object { $null -ne $_.EmailAddress } | Select-Object -ExpandProperty EmailAddress | Select-Object -ExpandProperty Address
                $props["attachments"] = if ($true -eq $msg.HasAttachments) { $msg.Attachments | Select-Object -Property Name, Size, ContentType, @{Name = "Filecontent"; Expression = { $_.AdditionalProperties.contentBytes } } } { @() }

                $message = New-Object psobject -Property $props
                $out_messages += $message 
            }

            Write-Output $out_messages
        }
        catch {
            Write-Output $null
        }
    }
    process {}
    end {}
}

Function Get-UserTrackingPreferences {
    Param(
        [Parameter(Mandatory = $true)]
        [guid]
        $userid,
        [Parameter(Mandatory = $true)]
        [psobject]
        $service
    )
    begin {
        $fetch = @"
            <fetch>
                <entity name="usersettings">
                    <attribute name="incomingemailfilteringmethod" />
                    <filter>
                    <condition attribute="systemuserid" operator="eq" value="$($userid)" />
                    </filter>
                </entity>
            </fetch>
"@
        $result = $service.GetEntityDataByFetchSearchEC($fetch)
        $out_setting = [string]::Empty
        if ($null -ne $result -and $result.Entities.Count -eq 1) {
            $setting = $result.Entities | Select-Object -First 1
            $setting = ([OptionSetValue]$setting.Attributes["incomingemailfilteringmethod"]).Value
            $out_setting = [string]::Empty
            switch ($setting) {
                0 { $out_setting = "all" }
                1 { $out_setting = "correlate" }
                2 { $out_setting = "correlate" }
                3 { $out_setting = "correlate" }
                Default { $out_setting = "none" }
            }
        }
        Write-Output $out_setting
    }
}

#Find email activity in Dataverse by internetmessageid
Function Get-EmailByMessageId {
    param(
        [Parameter(Mandatory = $true)]
        [guid]
        $userid,
        [Parameter(Mandatory = $true)]
        [psobject]
        $service
    )
    begin {

    }
}

Function Add-IncomingEmailInDataverse {
    param(
        [Parameter(Mandatory = $false)]
        [psobject]
        $correlatedemail,
        [Parameter(Mandatory = $true)]
        [psobject]
        $newemail,
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [psobject]
        $service
    )
    begin {
        
        $email = New-Object Entity("email")
        
        $from_party = New-Object Entity("activityparty")
        $from_party.Attributes.Add("addressused", $newemail.sender)
        $from_party.Attributes.Add("participationtypemask", [OptionSetValue](New-Object OptionSetValue(1)))
        $resolvedFrom = Get-ResolvedAddress -service $service -email $newemail.sender | Sort-Object -Property order | Select-Object -First 1
        
        if ($null -ne $resolvedFrom) {
            $from_party.Attributes.Add("partyid", $resolvedFrom.ref)
        }

        $from_collection = New-Object List[Entity]
        $from_collection.Add($from_party)

        $to_collection = New-Object List[Entity]
        $cc_collection = New-Object List[Entity]
        $bcc_collection = New-Object List[Entity]

        #add to recipients
        foreach ($to in $newemail.torecipients) {
            $resolutions = Get-ResolvedAddress -email $to -service $service | Sort-Object -Property order -Unique
            foreach ($res in $resolutions) {
                $to_party = New-Object Entity("activityparty")
                $to_party.Attributes.Add("participationtypemask", [OptionSetValue](New-Object OptionSetValue(2)))
                $to_party.Attributes.Add("addressused", $to)
                $to_party.Attributes.Add("partyid", $res.ref)
                $to_collection.Add($to_party)
            }
        }
        #add cc recipients
        foreach ($cc in $newemail.ccrecipients) {
            $resolutions = Get-ResolvedAddress -email $cc -service $service
            foreach ($res in $resolutions) {
                $cc_party = New-Object Entity("activityparty")
                $cc_party.Attributes.Add("participationtypemask", [OptionSetValue](New-Object OptionSetValue(3)))
                $cc_party.Attributes.Add("addressused", $cc)
                $cc_party.Attributes.Add("partyid", [EntityReference]$res.ref)
                $cc_collection.Add($cc_party)
            }
        }

        #add bcc recipients
        foreach ($bcc in $newemail.bccrecipients) {
            $resolutions = Get-ResolvedAddress -email $bcc -service $service
            foreach ($res in $resolutions) {
                $bcc_party = New-Object Entity("activityparty")
                $bcc_party.Attributes.Add("participationtypemask", [OptionSetValue](New-Object OptionSetValue(4)))
                $bcc_party.Attributes.Add("addressused", $bcc)
                $bcc_party.Attributes.Add("partyid", $res.ref)
                $bcc_collection.Add($bcc_party)
            }
        }

        $email["from"] = [EntityCollection]::new($from_collection)
        $email["to"] = [EntityCollection]::new($to_collection)
        if ($cc_collection.Count -gt 0) {
            $email["cc"] = [EntityCollection]::new($cc_collection)
        }
        if ($bcc_collection.Count -gt 0) {
            $email["bcc"] = [EntityCollection]::new($bcc_collection)
        }
        $email["subject"] = $newemail.subject
        $email["description"] = $newemail.body
        if ($true -ne [string]::IsNullOrEmpty($newemail.inreplyto)) {
            #inreplyto is readonly, dont try to set it
            #$email["inreplyto"] = $newemail.inreplyto
            if($null -eq $correlatedemail){
                $correlatedemail = Get-CorrelatedEmail -service $service -inreplyto $newemail.inreplyto.Replace("<", "\u003C").Replace(">", "\u003E") -newEmail $newemail
            }
            if ($null -ne $correlatedemail) {
                $correlatedactivityid = New-Object EntityReference -Property @{
                    LogicalName = "email"
                    Id          = $correlatedemail.id
                }
                $email.Attributes.Add("correlatedactivityid", [EntityReference]$correlatedactivityid)
                $parentactivityid = New-Object EntityReference -Property @{
                    LogicalName = "email"
                    Id          = $correlatedemail.id
                }
                $email.Attributes.Add("parentactivityid", [EntityReference]$parentactivityid)
                if ($null -ne $correlatedemail.regardingobjectid) {
                    $email["regardingobjectid"] = [EntityReference]$correlatedemail.regardingobjectid
                }
            }
        }
        #$email["importance"] = $newemail.importance
        $email["messageid"] = $newemail.messageid
        $emailid = $dvservice.Create($email) 

        #Manage Attachments if any
        if ($null -ne $newemail.attachments -and 0 -ne $newemail.attachments.Count) {
            foreach ($att in $newemail.attachments) {
                
                $attachment = New-Object Entity("activitymimeattachment");
                $attachment["objecttypecode"] = "email"
                $attachment["subject"] = $att.Name
                $attachment["filename"] = $att.Name
                $attachment["mimetype"] = $att.ContentType
                $bytes = [System.Text.Encoding]::ASCII.GetBytes($att.Filecontent)
                $attachment["body"] = [Convert]::ToBase64String($bytes)
                $emailRef = New-Object EntityReference -Property @{
                    LogicalName = "email"
                    Id          = $emailid
                }
                $attachment["objectid"] = [EntityReference]$emailRef

                $service.Create($attachment)
            }
        }

        Write-Output $emailid
    }
}

Function Get-ResolvedAddress {
    Param(
        [Parameter(Mandatory = $true)]
        [string]
        $email,
        [Parameter(Mandatory = $true)]
        [psobject]
        $service
    )
    begin {
        $fetch_emailsearch = @"
        <fetch>
            <entity name="emailsearch">
                <attribute name="emailaddress" />
                <attribute name="parentobjectid" />
                <filter>
                    <condition attribute="emailaddress" operator="eq" value="$($email)" />
                </filter>
                <link-entity name="systemuser" from="systemuserid" to="parentobjectid" link-type="outer" alias="user">
                    <attribute name="systemuserid" />
                    <attribute name="fullname" />
                    <filter>
                        <condition attribute="isdisabled" operator="eq" value="0" />
                    </filter>
                </link-entity>
                <link-entity name="queue" from="queueid" to="parentobjectid" link-type="outer" alias="queue">
                    <attribute name="queueid" />
                    <attribute name="name" />
                    <filter>
                        <condition attribute="statecode" operator="eq" value="0" />
                    </filter>
                </link-entity>
                <link-entity name="contact" from="contactid" to="parentobjectid" link-type="outer" alias="contact">
                    <attribute name="contactid" />
                    <attribute name="fullname" />
                    <filter>
                        <condition attribute="statecode" operator="eq" value="0" />
                    </filter>
                </link-entity>
                <link-entity name="account" from="accountid" to="parentobjectid" link-type="outer" alias="account">
                    <attribute name="accountid" />
                    <attribute name="name" />
                    <filter>
                        <condition attribute="statecode" operator="eq" value="0" />
                    </filter>
                </link-entity>
            </entity>
        </fetch>
"@

        $result = $service.GetEntityDataByFetchSearchEC($fetch_emailsearch)
        if ($result.Entities.Count -gt 0 -and $result.Entities.Count -lt 100) {
            $resolutions = @()
            foreach ($addressresolution in $result.Entities) {
                $ref = $null
                $props = @{}
                #is it user
                if ($addressresolution.Attributes.Contains("user.systemuserid")) {
                    $id = [Guid]([AliasedValue]$addressresolution.Attributes["user.systemuserid"]).Value
                    $name = [string]([AliasedValue]$addressresolution.Attributes["user.fullname"]).Value
                    $ref = New-Object EntityReference -Property @{
                        Id          = $id
                        LogicalName = "systemuser"
                        Name        = $name
                    }
                    
                    $props["order"] = 1
                    $props["ref"] = $ref
                    
                    $r = New-Object psobject -Property $props
                    $resolutions += $r
                }
                #is it queue
                elseif ($addressresolution.Attributes.Contains("queue.queueid")) {
                    $id = [Guid]([AliasedValue]$addressresolution.Attributes["queue.queueid"]).Value
                    $name = [string]([AliasedValue]$addressresolution.Attributes["queue.name"]).Value
                    $ref = New-Object EntityReference -Property @{
                        Id          = $id
                        LogicalName = "queue"
                        Name        = $name
                    }

                    $props["order"] = 2
                    $props["ref"] = $ref
                    
                    $r = New-Object psobject -Property $props
                    $resolutions += $r
                }
                #is it contact
                elseif ($addressresolution.Attributes.Contains("contact.contactid")) {
                    $id = [Guid]([AliasedValue]$addressresolution.Attributes["contact.contactid"]).Value
                    $name = [string]([AliasedValue]$addressresolution.Attributes["contact.fullname"]).Value
                    $ref = New-Object EntityReference -Property @{
                        Id          = $id
                        LogicalName = "contact"
                        Name        = $name
                    }
                    $props["order"] = 3
                    $props["ref"] = $ref
                    
                    $r = New-Object psobject -Property $props
                    $resolutions += $r
                }
                #is it account
                elseif ($addressresolution.Attributes.Contains("account.accountid")) {
                    $id = [Guid]([AliasedValue]$addressresolution.Attributes["account.accountid"]).Value
                    $name = [string]([AliasedValue]$addressresolution.Attributes["account.name"]).Value
                    $ref = New-Object EntityReference -Property @{
                        Id          = $id
                        LogicalName = "account"
                        Name        = $name
                    }

                    $props["order"] = 4
                    $props["ref"] = $ref
                    
                    $r = New-Object psobject -Property $props
                    $resolutions += $r
                }
            }
            
            Write-Output ($resolutions | Sort-Object -Property order) 
        }
        else {
            Write-Output @()
        }
    }
}

#Checks in dataverse if the 
Function Get-CorrelatedEmail {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $inreplyto,
        [psobject]
        $newEmail,
        [psobject]
        $service
    )

    begin {

        <#
            $props["emailcorrelationenabled"] = $organization.Attributes["emailcorrelationenabled"]
            $props["trackingprefix"] = $filtered_prefixes
            $props["emailconnectionchannel"] = ([OptionSetValue]$organization.Attributes["emailconnectionchannel"]).Value
            
        #>

        $orgsettings = Get-EmailTrackingConfiguration -service $service
        if ($null -eq $orgsettings -or $false -eq $orgsettings.emailcorrelationenabled) {
            Write-Output $null
        }
        else {
            $trackingNumber = $newEmail.subject.Split($orgsettings.trackingprefix) | Select-Object -Last 1
            $trackingtoken = ""
            if ($null -ne $trackingNumber) {

                $trackingtoken = "$($orgsettings.trackingprefix)$($trackingNumber)"
                $trackingtokenfilter = "<condition attribute=`"subject`" operator=`"ends-with`" value=`"$($trackingtoken)`" />"
            }
            #1 - check the inreplyto field
            $fetch = @"
            <fetch top="1">
                <entity name="email">
                    <attribute name="inreplyto" />
                    <attribute name="subject" />
                    <attribute name="activityid" />
                    <attribute name="regardingobjectid" />
                    <filter type="or">
                        <condition attribute="inreplyto" operator="eq" value="$($inreplyto)" />
                        $trackingtokenfilter
                    </filter>
                </entity>
            </fetch>
"@

            $emails = $service.GetEntityDataByFetchSearchEC($fetch)
            
            if ($null -ne $emails -and $emails.Entities.Count -eq 1) {
                $cemail = $emails.Entities | Select-Object -First 1
                $props = @{
                    "id"                = [guid]$cemail.Attributes["activityid"]
                    "regardingobjectid" = [EntityReference]$cemail.Attributes["regardingobjectid"]
                }
                $outemail = New-Object psobject -Property $props
                Write-Output $outemail
            }
            else {
                Write-Output $null
            }
        }
    }
}

Function Get-DataverseOutgoingEmail {
    param(
        [Parameter(Mandatory = $true)]
        [psobject] $service
    )
    begin {
        $fetch = @"
            <fetch top="100">
                <entity name="email">
                    <attribute name="inreplyto" />
                    <attribute name="subject" />
                    <attribute name="activityid" />
                    <attribute name="description" />
                    <attribute name="regardingobjectid" />
                    <attribute name="messageid" />
                    <attribute name="sender" />
                    <attribute name="torecipients" />
                    <filter type="and">
                        <condition attribute="statuscode" operator="eq" value="6" />
                        <condition attribute="directioncode" operator="eq" value="1" />
                    </filter>
                    <link-entity name="activityparty" from="activityid" to="activityid" alias="cc" link-type="outer">
                        <attribute name="addressused" />
                    </link-entity>
                    <link-entity name="activityparty" from="activityid" to="activityid" alias="bcc" link-type="outer">
                        <attribute name="addressused" />
                    </link-entity>
                    <link-entity name="activitymimeattachment" from="objectid" to="activityid" link-type="outer" alias="mimeattachments">
                        <attribute name="attachmentid" />
                        <link-entity name="attachment" from="attachmentid" to="attachmentid" link-type="outer" alias="attachments">
                            <attribute name="filename" />
                            <attribute name="body" />
                            <attribute name="mimetype" />
                        </link-entity>
                    </link-entity>
                </entity>
            </fetch>
"@
        $emails = $service.GetEntityDataByFetchSearchEC($fetch);
        $arrEmails = @()

        if ($null -ne $emails -and $null -ne $emails.Entities -and 0 -lt $emails.Entities.Count) {
            $messageids = $emails.Entities | Sort-Object -Property Id -Unique | Select-Object -ExpandProperty Id
            $outmsgs = @()
            foreach ($id in $messageids) {
                $e = $emails.Entities | Where-Object { $_.Id -eq $id }
                $fe = $e | Select-Object -First 1
                $ccs = @()
                foreach ($cce in $e) {
                    if ($null -ne $cce["cc.addressused"]) {
                        $recipient = [string]([AliasedValue]$cce["cc.addressused"]).Value
                        $ccs += @{ emailAddress = @{ address = $recipient } }
                    }
                }
                $bccs = @()
                foreach ($bcce in $e) {
                    if ($null -ne $bcce["bcc.addressused"]) {
                        $recipient = [string]([AliasedValue]$bcce["bcc.addressused"]).Value
                        $bccs += @{ emailAddress = @{ address = $recipient } }
                    }
                }
                $atts = @()
                foreach ($att in $e) {
                    if ($null -ne $att["attachments.mimetype"]) {
                        $atts += @{
                            "@odata.type" = "#microsoft.graph.fileAttachment"
                            name          = [string]([AliasedValue]$att["attachments.filename"]).Value
                            contentType   = [string]([AliasedValue]$att["attachments.mimetype"]).Value
                            contentBytes  = [string]([AliasedValue]$att["attachments.body"]).Value
                        }
                    }
                }
                $tos = $fe["torecipients"].Split(";")
                $torecipients = @()
                foreach ($to in $tos) {
                    $torecipients += @{ emailAddress = @{ address = $to } }
                }

                $props = @{
                    "messageid"   = $fe["messageid"]
                    "body"        = $fe["description"]
                    "subject"        = $fe["subject"]
                    "to"          = $torecipients
                    "from"        = $fe["sender"]
                    "cc"          = $ccs
                    "bcc"         = $bccs
                    "attachments" = $atts
                }
                $outmsgs += (New-Object psobject -Property $props)
            }
            Write-Output $outmsgs
        }
        else {
            Write-Output $null
        }
    }
}

Function Send-Email {
    param(
        [Parameter(Mandatory = $true)]
        [psobject] $token,
        [psobject] $message
    )
    begin {
        $headers = @{
            "Authorization" = "Bearer $($token.access_token)"
            "Content-type" = "application/json"
        }
        $MailFrom = $message.from
        $URLMail = "https://graph.microsoft.com/v1.0/users/$MailFrom/messages"
        $BodyJsonsend = @{
                subject = $message.subject
                body = @{
                    contentType = "HTML"
                    content = $message.body
                }
                toRecipients = $message.to
                ccRecipients = $message.cc
            #attachments = ($message.attachments | Select-Object -First 2)
        } | ConvertTo-Json -Depth 10
        #"ccRecipients": "$($message.cc | ConvertTo-Json -Depth 4)",
        #"bccRecipients": "$($message.bcc | ConvertTo-Json -Depth 4)"
        Write-Host ($BodyJsonsend | ConvertTo-Json -Depth 10) 
        try{
            $createdmessage = Invoke-RestMethod -Method POST -Uri $URLMail -Headers $headers -Body $BodyJsonsend
            if($null -ne $createdmessage -and $null -ne $createdmessage.id) {
                foreach($att in $message.attachments) {
                    $attachmentsBody = $att | ConvertTo-Json -Depth 10
                    $URLAttachment = "https://graph.microsoft.com/v1.0/users/$MailFrom/messages/$($createdmessage.id)/attachments"
                    Invoke-RestMethod -Method POST -Uri $URLAttachment -Headers $headers -Body $attachmentsBody
                }
                $urlSend = "https://graph.microsoft.com/v1.0/users/$MailFrom/messages/$($createdmessage.id)/Send"
                Invoke-RestMethod -Method POST -Uri $urlSend -Headers $headers
            }
        }catch{
            
            write-Host $_.Exception.Response.StatusDescription
        }
    }
}


Export-ModuleMember *