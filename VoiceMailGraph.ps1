function Send-VoiceMail {
    [CmdletBinding()]
    param (       
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $MailboxName,
        [Parameter(Position = 2, Mandatory = $false)]
        [String]
        $ClientId,
        [Parameter(Position = 3, Mandatory = $true)] [string]$Mp3FileName,
        [Parameter(Position = 4, Mandatory = $true)] [string]$ToAddress, 
        [Parameter(Position = 5, Mandatory = $false)] [string]$Transcription 
    )
    Begin {        
        $shell = New-Object -COMObject Shell.Application
        $folder = Split-Path $Mp3FileName
        $file = Split-Path $Mp3FileName -Leaf
        $shellfolder = $shell.Namespace($folder)
        $shellfile = $shellfolder.ParseName($file)
        $dt = [DateTime]::ParseExact($shellfolder.GetDetailsOf($shellfile, 27), "HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture);
        if ([String]::IsNullOrEmpty($ClientId)) {
            $ClientId = "5471030d-f311-4c5d-91ef-74ca885463a7"
        }		
        Import-Module .\Microsoft.IdentityModel.Clients.ActiveDirectory.dll -Force
        $PromptBehavior = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters -ArgumentList Auto       
        $Context = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext("https://login.microsoftonline.com/common")
        $token = ($Context.AcquireTokenAsync("https://graph.microsoft.com", $ClientId , "urn:ietf:wg:oauth:2.0:oob", $PromptBehavior)).Result
        $Header = @{
            'Content-Type'  = 'application\json'
            'Authorization' = $token.CreateAuthorizationHeader()
        }

        $UserResult = Invoke-RestMethod -Headers $Header -Uri ("https://graph.microsoft.com/v1.0/users('" + $MailboxName + "')?`$Select=displayName,businessPhones,mobilePhone,mail,jobTitle,companyName") -Method Get -ContentType "Application/json"
        $VoiceMailSuject = "Voice Mail (" + $dt.TimeOfDay.TotalSeconds + " seconds)"
        $duration = $dt.TimeOfDay.TotalSeconds
        $voiceMailFrom = $UserResult.displayName
        if ($UserResult.businessPhones.count -gt 0) {
            $callerId = $UserResult.businessPhones[0]
        }        
        $jobTitle = $UserResult.jobTitle.ToString()
        $Company = $UserResult.companyName
        $BusinessPhone = $callerId
        $emailAddress = $UserResult.mail
        $MobilePhone = $UserResult.mobilePhone.ToString()
       
        $BodyHtml = "<html><head><META HTTP-EQUIV=`"Content-Type`" CONTENT=`"text/html; charset=us-ascii`">"
        $BodyHtml += "<style type=`"text/css`"> a:link { color: #3399ff; } a:visited { color: #3366cc; } a:active { color: #ff9900; } </style>"
        $BodyHtml += "</head><body><style type=`"text/css`"> a:link { color: #3399ff; } a:visited { color: #3366cc; } a:active { color: #ff9900; } </style>"
        $BodyHtml += "<div style=`"font-family: Tahoma; background-color: #ffffff; color: #000000; font-size:10pt;`"><div id=`"UM-call-info`" lang=`"en`">"
        $BodyHtml += "<div style=`"font-family: Arial; font-size: 10pt; color:#000066; font-weight: bold;`">You received a voice mail from " + $voiceMailFrom + " at " + $MailboxName + "</div>"
        $BodyHtml += "<br><table border=`"0`" width=`"100%`">"
        $BodyHtml += "<tr><td width=`"12px`"></td><td width=`"28%`" nowrap=`"`" style=`"font-family: Tahoma; color: #686a6b; font-size:10pt;border-width: 0in;`">"
        $BodyHtml += "Company:</td><td width=`"72%`" style=`"font-family: Tahoma; background-color: #ffffff; color: #000000; font-size:10pt;`">"
        $BodyHtml += $Company + "</td></tr>"
        $BodyHtml += "<tr><td width=`"12px`"></td><td width=`"28%`" nowrap=`"`" style=`"font-family: Tahoma; color: #686a6b; font-size:10pt;border-width: 0in;`">"
        $BodyHtml += "Title:</td><td width=`"72%`" style=`"font-family: Tahoma; background-color: #ffffff; color: #000000; font-size:10pt;`">"
        $BodyHtml += $jobTitle + "</td></tr><tr><td width=`"12px`"></td><td width=`"28%`" nowrap=`"`" style=`"font-family: Tahoma; color: #686a6b; font-size:10pt;border-width: 0in;`">"
        $BodyHtml += "Work:</td><td width=`"72%`" style=`"font-family: Tahoma; background-color: #ffffff; color: #000000; font-size:10pt;`">"
        $BodyHtml += "<a style=`"color: #3399ff; `" dir=`"ltr`" href=`"tel:" + $BusinessPhone + "`">" + $BusinessPhone + "</a></td></tr>"
        $BodyHtml += "<tr><td width=`"12px`"></td><td width=`"28%`" nowrap=`"`" style=`"font-family: Tahoma; color: #686a6b; font-size:10pt;border-width: 0in;`">"
        $BodyHtml += "Mobile:</td><td width=`"72%`" style=`"font-family: Tahoma; background-color: #ffffff; color: #000000; font-size:10pt;`">"
        $BodyHtml += "<a style=`"color: #3399ff; `" dir=`"ltr`" href=`"tel:&#43;" + $MobilePhone + "`">&#43;" + $MobilePhone + "</a></td></tr>"
        $BodyHtml += "</table></div></div></body></html>"
        $ToRecp = "" | Select-Object Name, Address
        $ToRecp.Name = $ToAddress 
        $ToRecp.Address = $ToAddress
        $SenderAddress = "" | Select-Object Name, Address
        $SenderAddress.Name = $MailboxName 
        $SenderAddress.Address = $MailboxName
        $ItemClassProp = "" | Select Id, DataType, PropertyType, Value
        $ItemClassProp.id = "0x001A"
        $ItemClassProp.DataType = "String"
        $ItemClassProp.PropertyType = "Tagged"
        $ItemClassProp.Value = "IPM.Note.Microsoft.Voicemail.UM.CA"
        $VoiceMailLength = "" | Select Id, DataType, PropertyType, Type, Guid, Value
        $VoiceMailLength.id = "0x6801"
        $VoiceMailLength.DataType = "Integer"
        $VoiceMailLength.Guid = "{00020328-0000-0000-c000-000000000046}"
        $VoiceMailLength.PropertyType = "Named"
        $VoiceMailLength.Type = "Id"
        $VoiceMailLength.Value = $dt.TimeOfDay.TotalSeconds
        $VoiceMessageConfidenceLevel = "" | Select Id, DataType, Type, PropertyType, Guid, Value
        $VoiceMessageConfidenceLevel.Id = "X-VoiceMessageConfidenceLevel"
        $VoiceMessageConfidenceLevel.DataType = "String"
        $VoiceMessageConfidenceLevel.Guid = "{00020386-0000-0000-C000-000000000046}"
        $VoiceMessageConfidenceLevel.PropertyType = "Named"
        $VoiceMessageConfidenceLevel.Value = "high"
        $VoiceMessageConfidenceLevel.Type = "String"
        $VoiceMessageTranscription = "" | Select Id, DataType, Type, PropertyType, Guid, Value
        $VoiceMessageTranscription.Id = "X-VoiceMessageTranscription"
        $VoiceMessageTranscription.DataType = "String"
        $VoiceMessageTranscription.Guid = "{00020386-0000-0000-C000-000000000046}"
        $VoiceMessageTranscription.PropertyType = "Named"
        $VoiceMessageTranscription.Value = $Transcription
        $VoiceMessageTranscription.Type = "String"
        $PidTagVoiceMessageAttachmentOrder = "" | Select Id, DataType, PropertyType, Value
        $PidTagVoiceMessageAttachmentOrder.id = "0x6805"
        $PidTagVoiceMessageAttachmentOrder.DataType = "String"
        $PidTagVoiceMessageAttachmentOrder.PropertyType = "Tagged"
        $PidTagVoiceMessageAttachmentOrder.Value = "audio.mp3"
        $exProp = @()
        $exProp += $ItemClassProp
        $exProp += $VoiceMailLength
        $exProp += $VoiceMessageConfidenceLevel
        $exProp += $VoiceMessageTranscription
        $exProp += $PidTagVoiceMessageAttachmentOrder
        $Attachment = "" | Select name, contentBytes
        $Attachment.name = "audio.mp3"
        $Attachment.contentBytes = [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($Mp3FileName))
        $NewMessage = Get-MessageJSONFormat -Subject $VoiceMailSuject -Body $BodyHtml.Replace("`"", "\`"") -SenderEmailAddress $SenderAddress -Attachments @($Attachment) -ToRecipients @($ToRecp) -SaveToSentItems "true" -SendMail -ExPropList $exProp
        $Result = Invoke-RestMethod -Headers $Header -Uri ("https://graph.microsoft.com/v1.0/users('" + $MailboxName + "')/sendmail") -Method Post -Body $NewMessage -ContentType "Application/json"
        if ($Result.value -ne $null) {
            foreach ($Message in $Result.value ) {
                write-output $Message
            }
        }
       
       	
		
    }
}

function Get-MessageJSONFormat {
    [CmdletBinding()]
    param (
        [Parameter(Position = 1, Mandatory = $false)]
        [String]
        $Subject,
		
        [Parameter(Position = 2, Mandatory = $false)]
        [String]
        $Body,
		
        [Parameter(Position = 3, Mandatory = $false)]
        [psobject]
        $SenderEmailAddress,
		
        [Parameter(Position = 5, Mandatory = $false)]
        [psobject]
        $Attachments,
		
        [Parameter(Position = 5, Mandatory = $false)]
        [psobject]
        $ReferanceAttachments,
		
        [Parameter(Position = 6, Mandatory = $false)]
        [psobject]
        $ToRecipients,
		
        [Parameter(Position = 7, Mandatory = $false)]
        [psobject]
        $CcRecipients,
		
        [Parameter(Position = 7, Mandatory = $false)]
        [psobject]
        $bccRecipients,
		
        [Parameter(Position = 8, Mandatory = $false)]
        [psobject]
        $SentDate,
		
        [Parameter(Position = 9, Mandatory = $false)]
        [psobject]
        $StandardPropList,
		
        [Parameter(Position = 10, Mandatory = $false)]
        [psobject]
        $ExPropList,
		
        [Parameter(Position = 11, Mandatory = $false)]
        [switch]
        $ShowRequest,
		
        [Parameter(Position = 12, Mandatory = $false)]
        [String]
        $SaveToSentItems,
		
        [Parameter(Position = 13, Mandatory = $false)]
        [switch]
        $SendMail,
		
        [Parameter(Position = 14, Mandatory = $false)]
        [psobject]
        $ReplyTo,
		
        [Parameter(Position = 17, Mandatory = $false)]
        [bool]
        $RequestReadRecipient,
		
        [Parameter(Position = 18, Mandatory = $false)]
        [bool]
        $RequestDeliveryRecipient
    )
    Process {
        $NewMessage = "{" + "`r`n"
        if ($SendMail.IsPresent) {
            $NewMessage += "  `"Message`" : {" + "`r`n"
        }
        if (![String]::IsNullOrEmpty($Subject)) {
            $NewMessage += "`"Subject`": `"" + $Subject + "`"" + "`r`n"
        }
        if ($SenderEmailAddress -ne $null) {
            if ($NewMessage.Length -gt 5) { $NewMessage += "," }
            $NewMessage += "`"Sender`":{" + "`r`n"
            $NewMessage += " `"EmailAddress`":{" + "`r`n"
            $NewMessage += "  `"Name`":`"" + $SenderEmailAddress.Name + "`"," + "`r`n"
            $NewMessage += "  `"Address`":`"" + $SenderEmailAddress.Address + "`"" + "`r`n"
            $NewMessage += "}}" + "`r`n"
        }
        if (![String]::IsNullOrEmpty($Body)) {
            if ($NewMessage.Length -gt 5) { $NewMessage += "," }
            $NewMessage += "`"Body`": {" + "`r`n"
            $NewMessage += "`"ContentType`": `"HTML`"," + "`r`n"
            $NewMessage += "`"Content`": `"" + $Body + "`"" + "`r`n"
            $NewMessage += "}" + "`r`n"
        }
		
        $toRcpcnt = 0;
        if ($ToRecipients -ne $null) {
            if ($NewMessage.Length -gt 5) { $NewMessage += "," }
            $NewMessage += "`"ToRecipients`": [ " + "`r`n"
            foreach ($EmailAddress in $ToRecipients) {
                if ($toRcpcnt -gt 0) {
                    $NewMessage += "      ,{ " + "`r`n"
                }
                else {
                    $NewMessage += "      { " + "`r`n"
                }
                $NewMessage += " `"EmailAddress`":{" + "`r`n"
                $NewMessage += "  `"Name`":`"" + $EmailAddress.Name + "`"," + "`r`n"
                $NewMessage += "  `"Address`":`"" + $EmailAddress.Address + "`"" + "`r`n"
                $NewMessage += "}}" + "`r`n"
                $toRcpcnt++
            }
            $NewMessage += "  ]" + "`r`n"
        }
        $ccRcpcnt = 0
        if ($CcRecipients -ne $null) {
            if ($NewMessage.Length -gt 5) { $NewMessage += "," }
            $NewMessage += "`"CcRecipients`": [ " + "`r`n"
            foreach ($EmailAddress in $CcRecipients) {
                if ($ccRcpcnt -gt 0) {
                    $NewMessage += "      ,{ " + "`r`n"
                }
                else {
                    $NewMessage += "      { " + "`r`n"
                }
                $NewMessage += " `"EmailAddress`":{" + "`r`n"
                $NewMessage += "  `"Name`":`"" + $EmailAddress.Name + "`"," + "`r`n"
                $NewMessage += "  `"Address`":`"" + $EmailAddress.Address + "`"" + "`r`n"
                $NewMessage += "}}" + "`r`n"
                $ccRcpcnt++
            }
            $NewMessage += "  ]" + "`r`n"
        }
        $bccRcpcnt = 0
        if ($bccRecipients -ne $null) {
            if ($NewMessage.Length -gt 5) { $NewMessage += "," }
            $NewMessage += "`"BccRecipients`": [ " + "`r`n"
            foreach ($EmailAddress in $bccRecipients) {
                if ($bccRcpcnt -gt 0) {
                    $NewMessage += "      ,{ " + "`r`n"
                }
                else {
                    $NewMessage += "      { " + "`r`n"
                }
                $NewMessage += " `"EmailAddress`":{" + "`r`n"
                $NewMessage += "  `"Name`":`"" + $EmailAddress.Name + "`"," + "`r`n"
                $NewMessage += "  `"Address`":`"" + $EmailAddress.Address + "`"" + "`r`n"
                $NewMessage += "}}" + "`r`n"
                $bccRcpcnt++
            }
            $NewMessage += "  ]" + "`r`n"
        }
        $ReplyTocnt = 0
        if ($ReplyTo -ne $null) {
            if ($NewMessage.Length -gt 5) { $NewMessage += "," }
            $NewMessage += "`"ReplyTo`": [ " + "`r`n"
            foreach ($EmailAddress in $ReplyTo) {
                if ($ReplyTocnt -gt 0) {
                    $NewMessage += "      ,{ " + "`r`n"
                }
                else {
                    $NewMessage += "      { " + "`r`n"
                }
                $NewMessage += " `"EmailAddress`":{" + "`r`n"
                $NewMessage += "  `"Name`":`"" + $EmailAddress.Name + "`"," + "`r`n"
                $NewMessage += "  `"Address`":`"" + $EmailAddress.Address + "`"" + "`r`n"
                $NewMessage += "}}" + "`r`n"
                $ReplyTocnt++
            }
            $NewMessage += "  ]" + "`r`n"
        }
        if ($RequestDeliveryRecipient) {
            $NewMessage += ",`"IsDeliveryReceiptRequested`": true`r`n"
        }
        if ($RequestReadRecipient) {
            $NewMessage += ",`"IsReadReceiptRequested`": true `r`n"
        }
        if ($StandardPropList -ne $null) {
            foreach ($StandardProp in $StandardPropList) {
                if ($NewMessage.Length -gt 5) { $NewMessage += "," }
                switch ($StandardProp.PropertyType) {
                    "Single" {
                        if ($StandardProp.QuoteValue) {
                            $NewMessage += "`"" + $StandardProp.Name + "`": `"" + $StandardProp.Value + "`"" + "`r`n"
                        }
                        else {
                            $NewMessage += "`"" + $StandardProp.Name + "`": " + $StandardProp.Value + "`r`n"
                        }
						
						
                    }
                    "Object" {
                        if ($StandardProp.isArray) {
                            $NewMessage += "`"" + $StandardProp.PropertyName + "`": [ {" + "`r`n"
                        }
                        else {
                            $NewMessage += "`"" + $StandardProp.PropertyName + "`": {" + "`r`n"
                        }
                        $acCount = 0
                        foreach ($PropKeyValue in $StandardProp.PropertyList) {
                            if ($acCount -gt 0) {
                                $NewMessage += ","
                            }
                            $NewMessage += "`"" + $PropKeyValue.Name + "`": `"" + $PropKeyValue.Value + "`"" + "`r`n"
                            $acCount++
                        }
                        if ($StandardProp.isArray) {
                            $NewMessage += "}]" + "`r`n"
                        }
                        else {
                            $NewMessage += "}" + "`r`n"
                        }
						
                    }
                    "ObjectCollection" {
                        if ($StandardProp.isArray) {
                            $NewMessage += "`"" + $StandardProp.PropertyName + "`": [" + "`r`n"
                        }
                        else {
                            $NewMessage += "`"" + $StandardProp.PropertyName + "`": {" + "`r`n"
                        }
                        foreach ($EnclosedStandardProp in $StandardProp.PropertyList) {
                            $NewMessage += "`"" + $EnclosedStandardProp.PropertyName + "`": {" + "`r`n"
                            foreach ($PropKeyValue in $EnclosedStandardProp.PropertyList) {
                                $NewMessage += "`"" + $PropKeyValue.Name + "`": `"" + $PropKeyValue.Value + "`"," + "`r`n"
                            }
                            $NewMessage += "}" + "`r`n"
                        }
                        if ($StandardProp.isArray) {
                            $NewMessage += "]" + "`r`n"
                        }
                        else {
                            $NewMessage += "}" + "`r`n"
                        }
                    }
					
                }
            }
        }
        $atcnt = 0
        $processAttachments = $false
        if ($Attachments -ne $null) { $processAttachments = $true }
        if ($ReferanceAttachments -ne $null) { $processAttachments = $true }
        if ($processAttachments) {
            if ($NewMessage.Length -gt 5) { $NewMessage += "," }
            $NewMessage += "  `"Attachments`": [ " + "`r`n"
            if ($Attachments -ne $null) {
                foreach ($Attachment in $Attachments) {
                    if ($atcnt -gt 0) {
                        $NewMessage += "   ,{" + "`r`n"
                    }
                    else {
                        $NewMessage += "    {" + "`r`n"
                    }
                    if ($Attachment.name) {
                        $NewMessage += "     `"@odata.type`": `"#Microsoft.OutlookServices.FileAttachment`"," + "`r`n"
                        $NewMessage += "     `"Name`": `"" + $Attachment.name + "`"," + "`r`n"
                        $NewMessage += "     `"ContentBytes`": `" " + $Attachment.contentBytes + "`"" + "`r`n"
                    }
                    else {
                        $Item = Get-Item $Attachment

                        $NewMessage += "     `"@odata.type`": `"#Microsoft.OutlookServices.FileAttachment`"," + "`r`n"
                        $NewMessage += "     `"Name`": `"" + $Item.Name + "`"," + "`r`n"
                        $NewMessage += "     `"ContentBytes`": `" " + [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($Attachment)) + "`"" + "`r`n"

                    }
                    $NewMessage += "    } " + "`r`n"
                    $atcnt++
					
                }
            }
            $atcnt = 0
            if ($ReferanceAttachments -ne $null) {
                foreach ($Attachment in $ReferanceAttachments) {
                    if ($atcnt -gt 0) {
                        $NewMessage += "   ,{" + "`r`n"
                    }
                    else {
                        $NewMessage += "    {" + "`r`n"
                    }
                    $NewMessage += "     `"@odata.type`": `"#Microsoft.OutlookServices.ReferenceAttachment`"," + "`r`n"
                    $NewMessage += "     `"Name`": `"" + $Attachment.Name + "`"," + "`r`n"
                    $NewMessage += "     `"SourceUrl`": `"" + $Attachment.SourceUrl + "`"," + "`r`n"
                    $NewMessage += "     `"ProviderType`": `"" + $Attachment.ProviderType + "`"," + "`r`n"
                    $NewMessage += "     `"Permission`": `"" + $Attachment.Permission + "`"," + "`r`n"
                    $NewMessage += "     `"IsFolder`": `"" + $Attachment.IsFolder + "`"" + "`r`n"
                    $NewMessage += "    } " + "`r`n"
                    $atcnt++
                }
            }
            $NewMessage += "  ]" + "`r`n"
        }
		
        if ($ExPropList -ne $null) {
            if ($NewMessage.Length -gt 5) { $NewMessage += "," }
            $NewMessage += "`"SingleValueExtendedProperties`": [" + "`r`n"
            $propCount = 0
            foreach ($Property in $ExPropList) {
                if ($propCount -eq 0) {
                    $NewMessage += "{" + "`r`n"
                }
                else {
                    $NewMessage += ",{" + "`r`n"
                }
                if ($Property.PropertyType -eq "Tagged") {
                    $NewMessage += "`"id`":`"" + $Property.DataType + " " + $Property.Id + "`", " + "`r`n"
                }
                else {
                    if ($Property.Type -eq "String") {
                        $NewMessage += "`"id`":`"" + $Property.DataType + " " + $Property.Guid + " Name " + $Property.Id + "`", " + "`r`n"
                    }
                    else {
                        $NewMessage += "`"id`":`"" + $Property.DataType + " " + $Property.Guid + " Id " + $Property.Id + "`", " + "`r`n"
                    }
                }
                if ($Property.Value -eq "null") {
                    $NewMessage += "`"Value`":null" + "`r`n"
                }
                else {
                    $NewMessage += "`"Value`":`"" + $Property.Value + "`"" + "`r`n"
                }				
                $NewMessage += " } " + "`r`n"
                $propCount++
            }
            $NewMessage += "]" + "`r`n"
        }
        if (![String]::IsNullOrEmpty($SaveToSentItems)) {
            $NewMessage += "}   ,`"SaveToSentItems`": `"" + $SaveToSentItems.ToLower() + "`"" + "`r`n"
        }
        $NewMessage += "}"
        if ($ShowRequest.IsPresent) {
            Write-Host $NewMessage
        }
        return, $NewMessage
    }
}

function Get-VoiceMail {
    [CmdletBinding()]
    param (       
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $MailboxName,
        [Parameter(Position = 2, Mandatory = $false)]
        [String]
        $ClientId
    )
    Process {        
        if ([String]::IsNullOrEmpty($ClientId)) {
            $ClientId = "5471030d-f311-4c5d-91ef-74ca885463a7"
        }		
        Import-Module .\Microsoft.IdentityModel.Clients.ActiveDirectory.dll -Force
        $PromptBehavior = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters -ArgumentList Auto       
        $Context = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext("https://login.microsoftonline.com/common")
        $token = ($Context.AcquireTokenAsync("https://graph.microsoft.com", $ClientId , "urn:ietf:wg:oauth:2.0:oob", $PromptBehavior)).Result
        $Header = @{
            'Content-Type'  = 'application\json'
            'Authorization' = $token.CreateAuthorizationHeader()
        }
        $Result = Invoke-RestMethod -Headers $Header -Uri ("https://graph.microsoft.com/v1.0/users('" + $MailboxName + "')/mailFolders/voicemail/messages?`$expand=SingleValueExtendedProperties(`$filter=Id%20eq%20'Integer%20{00020328-0000-0000-C000-000000000046}%20Id%200x6801'%20or%20Id%20eq%20'String%20{00020386-0000-0000-C000-000000000046}%20Name%20X-VoiceMessageConfidenceLevel'%20or%20Id%20eq%20'String%20{00020386-0000-0000-C000-000000000046}%20Name%20X-VoiceMessageTranscription')&`$top=100&`$select=Subject,From,Body,IsRead,Id,ReceivedDateTime")
        if ($Result.value -ne $null) {
            foreach ($Message in $Result.value ) {
                Expand-ExtendedProperties -Item $Message
                write-output $Message
            }
        }
       
       	
		
    }
}

function Expand-ExtendedProperties {
    [CmdletBinding()] 
    param (
        [Parameter(Position = 1, Mandatory = $false)]
        [psobject]
        $Item
    )
	
    process {
        if ($Item.singleValueExtendedProperties -ne $null) {
            foreach ($Prop in $Item.singleValueExtendedProperties) {
                Switch ($Prop.Id) {
   	                "Integer {00020328-0000-0000-C000-000000000046} Id 0x6801" {
                        Add-Member -InputObject $Item -NotePropertyName "PidTagVoiceMessageDuration" -NotePropertyValue $Prop.Value
                    }
                    "String {00020386-0000-0000-C000-000000000046} Name X-VoiceMessageTranscription" {
                        Add-Member -InputObject $Item -NotePropertyName "X-VoiceMessageTranscription" -NotePropertyValue $Prop.Value
                    }
                    "String {00020386-0000-0000-C000-000000000046} Name X-VoiceMessageConfidenceLevel" {
                        Add-Member -InputObject $Item -NotePropertyName "X-VoiceMessageConfidenceLevel" -NotePropertyValue $Prop.Value
                    }
 
                }
            }
        }
    }
}

