<#
.SYNOPSIS
Converts an Exchange EntryId (hex or base64) into a Microsoft Graph REST id using the translateExchangeIds endpoint.

.DESCRIPTION
Invoke-TranslateId accepts either a Base64-encoded folder/entry id (as produced by some Outlook/Exchange properties)
or a hex string representation of an EntryId and converts it into a REST id suitable for Microsoft Graph calls.
If Base64FolderIdToTranslate is provided the function will decode it and compute the hex EntryId. If HexEntryId is provided
directly it will be converted to a byte array and then to the base64 string required by the Graph translateExchangeIds API.
The function issues a POST to the Microsoft Graph translateExchangeIds endpoint for the specified mailbox and returns the
target REST id from the response.

.PARAMETER MailboxName
The SMTP address or user id of the mailbox to run the translation against. Used to scope the translateExchangeIds call
to the specified user's mailbox.

.PARAMETER Base64FolderIdToTranslate
A base64-encoded EntryId to translate. If supplied, this value is decoded and converted into the hex EntryId used internally.

.PARAMETER HexEntryId
A hex string representation of an EntryId. If supplied, it will be converted to the byte array and then to base64 for the translate call.
If Base64FolderIdToTranslate is provided that takes precedence and HexEntryId will be calculated from it.

.OUTPUTS
System.String
Returns the translated REST id (string) returned by the Graph translateExchangeIds API (targetId).

.EXAMPLE
# Convert a hex EntryId to a REST id
Invoke-TranslateId -MailboxName 'user@contoso.com' -HexEntryId 'AABBCCDDEEFF...'

.EXAMPLE
# Convert a base64 EntryId to a REST id
Invoke-TranslateId -MailboxName 'user@contoso.com' -Base64FolderIdToTranslate 'AQMk...'

.NOTES
- Requires that the calling context can call the Microsoft Graph translateExchangeIds endpoint (proper authentication and scopes).
- Uses Invoke-MgGraphRequest to call Graph; ensure Microsoft Graph PowerShell module (or equivalent request helper) is available.
- The function operates purely on id translation and does not access mailbox content directly except via the Graph translate API.
#>

function Invoke-TranslateId {
    [CmdletBinding()] 
    param( 
        [Parameter(Position = 1, Mandatory = $false)]
        [String]
        $MailboxName = "",
        [Parameter(Position = 2, Mandatory = $false)]
        [String]
        $Base64FolderIdToTranslate = "",
        [Parameter(Position = 3, Mandatory = $false)]
        [String]
        $HexEntryId = ""

    )  
    Process {
        Add-Type -AssemblyName System.Web
        if (![String]::IsNullOrEmpty($Base64FolderIdToTranslate)) {
            $HexEntryId = [System.BitConverter]::ToString([Convert]::FromBase64String($Base64FolderIdToTranslate)).Replace("-", "").Substring(2)  
            $HexEntryId = $HexEntryId.SubString(0, ($HexEntryId.Length - 2))
        }
        $IdBytes = [byte[]]::new($HexEntryId.Length / 2)
        For ($i = 0; $i -lt $HexEntryId.Length; $i += 2) {
            $IdBytes[$i / 2] = [convert]::ToByte($HexEntryId.Substring($i, 2), 16)
        }

        $IdToConvert = ConvertTo-UrlToken -BytesInput $IdBytes

        $ConvertRequest = @"
{
    "inputIds" : [
      "$IdToConvert"
    ],
    "sourceIdType": "entryId",
    "targetIdType": "restId"
  }
"@

        $ConvertResult = Invoke-MgGraphRequest -Method POST -Uri https://graph.microsoft.com/v1.0/users/$MailboxName/translateExchangeIds -Body $ConvertRequest
        return $ConvertResult.Value.targetId
    }
}

function ConvertTo-UrlToken {
    [CmdletBinding(DefaultParameterSetName='String')]
    param(
        [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true, ParameterSetName='String')]
        [string]$StringInput,

        [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true, ParameterSetName='Bytes')]
        [byte[]]$BytesInput,

        [string]$Encoding = 'UTF8'
    )
    
    begin {
        # Check if we are running in an older environment that might require loading the assembly
        # Not strictly needed in modern PS but good practice for compatibility.
        # Add-Type -AssemblyName System.Web # Uncomment if you need other System.Web functions
    }

    process {
        # 1. Convert String to Bytes if necessary
        if ($PSCmdlet.ParameterSetName -eq 'String') {
            $bytes = [System.Text.Encoding]::GetEncoding($Encoding).GetBytes($StringInput)
        } else {
            $bytes = $BytesInput
        }
        
        if ($bytes.Length -eq 0) {
            return ""
        }

        # 2. Standard Base64 Encoding
        $base64Str = [System.Convert]::ToBase64String($bytes)
        
        # 3. Find padding characters ('=')
        $endPos = $base64Str.Length
        $numPadChars = 0

        # Loop backwards to count padding characters
        for ($i = $base64Str.Length - 1; $i -ge 0; $i--) {
            if ($base64Str[$i] -eq '=') {
                $endPos--
                $numPadChars++
            } else {
                break
            }
        }

        # 4. Perform replacements and remove padding
        $token = $base64Str.Substring(0, $endPos)
        $token = $token.Replace('+', '-') # Replace '+' with '-'
        $token = $token.Replace('/', '_') # Replace '/' with '_'

        # 5. Append the padding count (0, 1, or 2) as the last character
        $token + $numPadChars.ToString()
    }
}
<#
.SYNOPSIS
Retrieves mailbox folder permissions by calling the new Exchange admin API and returns a lookup hashtable keyed by user.

.DESCRIPTION
Get-DelegateFolderPermissions uses an Outlook/Exchange admin API endpoint to run a Get-MailboxFolderPermission style operation for the
specified mailbox and folder. It constructs the admin API URL using the tenant discovered from the mailbox's domain and will
use a provided AccessToken or call Get-EntraToken to obtain one. The function sets headers including X-AnchorMailbox and
posts the CmdletInput body to retrieve permission entries. Results are returned in a hashtable where each key is the user
identifier and the value is the permission object returned by the admin API.

.PARAMETER MailboxName
The SMTP address of the mailbox to query (e.g. user@contoso.com). Used to determine tenant and to form the Identity parameter.

.PARAMETER FolderName
The mailbox folder name to retrieve permissions for (e.g. Calendar, Inbox, Tasks, Contacts, Notes, Journal).

.PARAMETER AccessToken
(Optional) A pre-obtained OAuth access token to use for authentication. If not provided, the function attempts to obtain one via Get-EntraToken.

.PARAMETER UseServicePrincipal
(Optional) Switch parameter. When set, the X-AnchorMailbox will be formed for the SystemMailbox anchor (used for service-principal or application scenarios).

.OUTPUTS
System.Collections.Hashtable
A hashtable mapping permission user identifiers (keys) to the permission objects returned by the admin API (values).

.EXAMPLE
# Get calendar permissions for a mailbox using an already-obtained token
Get-DelegateFolderPermissions -MailboxName 'user@contoso.com' -FolderName Calendar -AccessToken $token

.EXAMPLE
# Get permissions using service principal anchor
Get-DelegateFolderPermissions -MailboxName 'user@contoso.com' -FolderName Inbox -UseServicePrincipal

.NOTES
- This function calls a legacy Outlook/Exchange admin API endpoint; appropriate admin privileges and API access are required.
- If AccessToken is not provided the function will call Get-EntraToken; ensure that helper exists and returns a valid token.
- The returned permission objects reflect the structure returned by the admin API and may include properties like AccessRights and SharingPermissionFlags.
#>
function Get-DelegateFolderPermissions {
    [CmdletBinding()] 
    param( 
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $MailboxName,
        [Parameter(Position = 2, Mandatory = $true)]
        [String]
        $FolderName,
        [Parameter(Position = 3, Mandatory = $false)]
        [String]
        $AccessToken,
        [Parameter(Position = 4, Mandatory = $false)]
        [switch]
        $UseServicePrincipal = $false    
    )  
    Process {
        $domain = $MailboxName.Split('@')[1]
        $TenantId = (Invoke-WebRequest -Uri "https://login.microsoftonline.com/$domain/.well-known/openid-configuration" -Verbose:$false | ConvertFrom-Json).authorization_endpoint.Split('/')[3]
        $adminUri = "https://outlook.office365.com/adminapi/v2.0/$TenantId/MailboxFolderPermission"
        if ([String]::IsNullOrEmpty($AccessToken)) {
            $Token = Get-EntraToken
            $AccessToken = $Token.AccessToken
        }
        if (!$UseServicePrincipal) {
            $MailboxAnchor = "UPN:$MailboxName"
        }
        else {
            $MailboxAnchor = "UPN:SystemMailbox{bb558c35-97f1-4cb9-8ff7-d53741dc928c}@" + $MailboxName.Split('@')[1]   
        }
        $headers = @{
            'Content-Type'     = 'application/json'
            'Authorization'    = "Bearer $AccessToken"
            'X-ResponseFormat' = "json"
            'X-AnchorMailbox'  = $MailboxAnchor
        }
        $ContentType = 'application/json'
        $Body = @"
{
  "CmdletInput": {
    "CmdletName": "Get-MailboxFolderPermission",
    "Parameters": {
      "Identity": "${MailboxName}:\\${FolderName}",
      "ResultSize": "Unlimited"
    }
  }
}
"@
        $permissionHash = @{}
        $Result = Invoke-RestMethod -URI $adminUri -Headers $headers -Method "POST" -Body $Body -ContentType $ContentType
        foreach ($Permission in $Result.Value) {
            $permissionHash.add($Permission.User, $Permission)
        }
        return $permissionHash

    }    
}

<#
.SYNOPSIS
Retrieves the mailbox's local free/busy/delegate configuration message (the 'local free/busy object') via Microsoft Graph.

.DESCRIPTION
Get-LocalFreeBusyObject finds and returns the message that stores the local free/busy/delegate configuration for a mailbox.
It reads the root mail folder's extended properties to locate the PR_FREEBUSY_ENTRYIDS data, extracts the FreeBusy object EntryId,
translates that EntryId to a Graph REST id using Invoke-TranslateId, and then retrieves the message by id while expanding the
relevant multiValueExtendedProperties and singleValueExtendedProperties for delegate configuration.

.PARAMETER MailboxName
The SMTP address or user id of the mailbox to inspect. Used to call Get-MgUserMailFolder and Get-MgUserMessage.

.PARAMETER UseServicePrincipal
(Optional) Switch parameter. When set, indicates the function is being called in a service principal context.

.OUTPUTS
Microsoft.Graph.Message (or the equivalent object returned by Get-MgUserMessage)
Returns the message object that represents the local free/busy entry for the mailbox, including expanded extended properties
that contain delegate information.

.EXAMPLE
# Retrieve the local free/busy object for a mailbox
Get-LocalFreeBusyObject -MailboxName 'user@contoso.com'

.EXAMPLE
# Retrieve with service principal context
Get-LocalFreeBusyObject -MailboxName 'user@contoso.com'

.NOTES
- Requires Microsoft Graph calls (Get-MgUserMailFolder and Get-MgUserMessage) and adequate permissions to read mailbox folders/messages.
- Relies on the presence of specific extended properties and certain script-level property id variables (e.g. $Script:PidTagScheduleInfoDelegateEntryIds).
- If the expected extended properties are not present the function may return null or throw.
#>
function Get-LocalFreeBusyObject {
    [CmdletBinding()] 
    param( 
        [Parameter(Position = 1, Mandatory = $false)]
        [String]
        $MailboxName = ""
    )  
    Process {
        $PR_FREEBUSY_ENTRYIDS = Get-MgUserMailFolder -UserId $MailboxName -MailFolderId root -ExpandProperty "multiValueExtendedProperties(`$filter=id eq 'BinaryArray 0x36E4')"
        $LocalFreeBusyObjectIdBytes = [System.Convert]::FromBase64String($PR_FREEBUSY_ENTRYIDS.MultiValueExtendedProperties[0].Value[1])
        $LocalFreeBusyObjectIdHex = [System.BitConverter]::ToString($LocalFreeBusyObjectIdBytes).Replace("-", "")
        $LocalFreeBusyObjectRestId = Invoke-TranslateId -HexEntryId $LocalFreeBusyObjectIdHex -MailboxName $MailboxName
   
        $MultiValueFilterString = "id eq '$Script:PidTagScheduleInfoDelegateEntryIds' or " + `
            "id eq '$Script:PidTagScheduleInfoDelegateNamesW' or " + `
            "id eq '$Script:PidTagDelegateFlags'" 
        $SingleValueFilterString = "id eq '$Script:PidTagScheduleInfoDelegatorWantsInfo' or " + `
            "id eq '$Script:PidTagScheduleInfoDelegatorWantsCopy' or " + `
            "id eq '$Script:PidTagScheduleInfoDontMailDelegates'"            


        return Get-MgUserMessage -UserId $MailboxName -MessageId $LocalFreeBusyObjectRestId `
            -ExpandProperty "multiValueExtendedProperties(`$filter=$MultiValueFilterString), singleValueExtendedProperties(`$filter=$SingleValueFilterString)"
    }
}

<#
.SYNOPSIS
Parses and displays delegate configuration information from a mailbox's local free/busy object.

.DESCRIPTION
Get-DelegateInfoConfig retrieves the mailbox's local free/busy object (via Get-LocalFreeBusyObject) and parses its extended properties
to display delegate names, flags (e.g. whether the delegate can view private items), and entry ids. It also inspects single-value
properties to determine the delegator's settings for receiving copies of meeting messages and whether the delegator wants meeting info.
Output is written to the host using Write-Host for human-readable display.

.PARAMETER MailboxName
The SMTP address or user id of the mailbox to inspect.

.OUTPUTS
None (writes human-readable information to the host).
The function displays per-delegate information and one-line summary of the delegate configuration.

.EXAMPLE
# Show delegate configuration for a user
Get-DelegateInfoConfig -MailboxName 'user@contoso.com'

.EXAMPLE
# Show delegate configuration using service principal
Get-DelegateInfoConfig -MailboxName 'user@contoso.com' -UseServicePrincipal

.NOTES
- This function is intended for interactive/human consumption (uses Write-Host). It does not return structured delegate objects.
- Relies on Get-LocalFreeBusyObject and script-level property id variables (e.g. $Script:PidTagScheduleInfoDelegateNamesW, $Script:PidTagDelegateFlags).
- Handles missing or incomplete extended properties, but the displayed output depends on the presence and format of the properties returned by Graph.
#>
function Get-DelegateInfoConfig {
    [CmdletBinding()] 
    param( 
        [Parameter(Position = 1, Mandatory = $false)]
        [String]
        $MailboxName = ""
    )  
    Process {
        $Msg = Get-LocalFreeBusyObject -MailboxName $MailboxName 
    
        if ($Msg.MultiValueExtendedProperties) {
            Write-Host "Delegate Information Found:" -ForegroundColor Green
        
            # Create a hashtable for easy lookup
            $Props = @{}
            foreach ($p in $Msg.MultiValueExtendedProperties) {
                $Props[$p.Id] = $p.Value
            }

            # Display parallel arrays (assuming index 0 is Delegate 1, index 1 is Delegate 2, etc.)
            # Note: We use the Count of the Name array to determine how many delegates there are.
            $DelegateCount = if ($Props[$Script:PidTagScheduleInfoDelegateNamesW]) { $Props[$Script:PidTagScheduleInfoDelegateNamesW].Count } else { 0 }

            for ($i = 0; $i -lt $DelegateCount; $i++) {
                Write-Host "------------------------------------------------"
                Write-Host "Delegate #$($i + 1)" -ForegroundColor Cyan
            
                Write-Host "  Name:           $($Props[$Script:PidTagScheduleInfoDelegateNamesW][$i])"
                Write-Host "  Can See Private Items:          $($Props[$Script:PidTagDelegateFlags][$i])"
                Write-Host "  EntryID (B64):  $($Props[$Script:PidTagScheduleInfoDelegateEntryIds][$i])"
            }
        }
    
        # Fixed: Proper condition checking for SingleValueExtendedProperties
        if ($Msg.SingleValueExtendedProperties -and $Msg.SingleValueExtendedProperties.Count -ge 2) {
            $prop0 = $Msg.SingleValueExtendedProperties[0]
            $prop1 = $Msg.SingleValueExtendedProperties[1]
      
            if ($prop0 -and $prop1 -and $prop0.Value -ne $null -and $prop1.Value -ne $null) {
                $wantsInfo = [boolean]::Parse($prop0.Value)
                $wantsCopy = [boolean]::Parse($prop1.Value)
        
                if ($wantsCopy -and $wantsInfo) {
                    $DelegateConfig = "My Delegates only but send a copy"
                }
                elseif ($wantsCopy) {
                    $DelegateConfig = "My Delegates and Me"
                }
                else {
                    $DelegateConfig = "My Delegates only"
                }
            }
            else {
                $DelegateConfig = "Delegate configuration properties not found."
            }
        }
        else {
            $DelegateConfig = "Delegate configuration properties not found."
        }
    
        Write-Host $DelegateConfig -ForegroundColor Green
        Write-Host  
    }
}

<#
.SYNOPSIS
Builds and returns structured delegate objects for an Office 365 mailbox by combining folder permissions with the mailbox's delegate list.

.DESCRIPTION
Get-O365Delegate aggregates mailbox delegate information and mailbox folder permissions into a structured collection of objects.
It queries folder permissions for Calendar, Inbox, Tasks, Contacts, Notes, and Journal using Get-FolderPermissions, retrieves the
mailbox's local free/busy/delegate object via Get-LocalFreeBusyObject, and then produces an array of PSCustomObject entries where
each entry represents a delegate with folder permission levels, whether they receive copies of meeting messages, and whether they
can view private items.

.PARAMETER MailboxName
The SMTP address or user id of the mailbox to inspect.

.PARAMETER AccessToken
(Optional) A pre-obtained OAuth access token to use for authentication with the admin API. If not provided, the function will call Get-EntraToken.

.PARAMETER UseServicePrincipal
(Optional) Switch parameter. When set, folder permission queries will use the service principal/system mailbox anchor behavior
for X-AnchorMailbox (useful when calling the admin API from an application/service principal context).

.OUTPUTS
System.Object[]
An array of PSCustomObject items. Each object contains:
- UserId: delegate identifier/name
- Permissions: object with CalendarFolderPermissionLevel, TasksFolderPermissionLevel, InboxFolderPermissionLevel, ContactsFolderPermissionLevel, NotesFolderPermissionLevel, JournalFolderPermissionLevel
- ReceiveCopiesOfMeetingMessages: boolean
- ViewPrivateItems: boolean

.EXAMPLE
# Get delegates and their folder permissions for a mailbox
Get-O365Delegate -MailboxName 'user@contoso.com'

.EXAMPLE
# Use service principal anchor when querying folder permissions
Get-O365Delegate -MailboxName 'user@contoso.com' -UseServicePrincipal

.EXAMPLE
# Use with a specific access token
Get-O365Delegate -MailboxName 'user@contoso.com' -AccessToken $token -UseServicePrincipal

.NOTES
- This function depends on Get-DelegateFolderPermissions and Get-LocalFreeBusyObject; ensure those helpers are available and functioning.
- Requires appropriate permissions to read mailbox folder permissions and message extended properties.
- The final objects are suitable for programmatic consumption (exporting, filtering, further automation).
- When using a service principal, ensure proper application permissions are granted (e.g., Mail.ReadWrite, Calendars.ReadWrite).
#>
function Get-O365Delegate {
    [CmdletBinding()] 
    param( 
        [Parameter(Position = 1, Mandatory = $false)]
        [String]
        $MailboxName = "",
        [Parameter(Position = 2, Mandatory = $false)]
        [String]
        $AccessToken = "",
        [Parameter(Position = 3, Mandatory = $false)]
        [switch]
        $UseServicePrincipal     
    )  
    Process {
        # Prepare parameters for Get-FolderPermissions calls
        $folderParams = @{
            MailboxName         = $MailboxName
            UseServicePrincipal = $UseServicePrincipal.IsPresent
        }
    
        if (![String]::IsNullOrEmpty($AccessToken)) {
            $folderParams.AccessToken = $AccessToken
        }
    
        # Retrieve folder permissions for all relevant folders
        $CalendarPermissions = Get-DelegateFolderPermissions @folderParams -FolderName Calendar
        $InboxPermissions = Get-DelegateFolderPermissions @folderParams -FolderName Inbox
        $TasksPermissions = Get-DelegateFolderPermissions @folderParams -FolderName Tasks
        $ContactsPermissions = Get-DelegateFolderPermissions @folderParams -FolderName Contacts
        $NotesPermissions = Get-DelegateFolderPermissions @folderParams -FolderName Notes
        $JournalPermissions = Get-DelegateFolderPermissions @folderParams -FolderName Journal
    
        $DelegateObjects = [PSCustomObject]@{
            Delegates              = @()
            DeliverMeetingRequests = ""
        }  

        $Msg = Get-LocalFreeBusyObject -MailboxName $MailboxName 
    
        if ($Msg.MultiValueExtendedProperties) {
            $Props = @{}
            foreach ($p in $Msg.MultiValueExtendedProperties) {
                $Props[$p.Id] = $p.Value
            }
            $DelegateCount = if ($Props[$Script:PidTagScheduleInfoDelegateNamesW]) { $Props[$Script:PidTagScheduleInfoDelegateNamesW].Count } else { 0 }
      
            for ($i = 0; $i -lt $DelegateCount; $i++) {
                $PermissionsObj = [PSCustomObject]@{
                    CalendarFolderPermissionLevel = "None"
                    TasksFolderPermissionLevel    = "None"
                    InboxFolderPermissionLevel    = "None"
                    ContactsFolderPermissionLevel = "None"
                    NotesFolderPermissionLevel    = "None"
                    JournalFolderPermissionLevel  = "None"
                }

                if ($($Props[$Script:PidTagDelegateFlags][$i]) -eq 1) {
                    $ViewPrivateItems = $true
                }
                else {
                    $ViewPrivateItems = $false
                }

                # Create the Final DelegateUser Object
                $DelegateUser = [PSCustomObject]@{
                    UserId                         = $Props[$Script:PidTagScheduleInfoDelegateNamesW][$i]
                    Permissions                    = $PermissionsObj
                    ReceiveCopiesOfMeetingMessages = $false
                    ViewPrivateItems               = $ViewPrivateItems
                }  
            
                if ($CalendarPermissions.ContainsKey($Props[$Script:PidTagScheduleInfoDelegateNamesW][$i]) -and $CalendarPermissions[$Props[$Script:PidTagScheduleInfoDelegateNamesW][$i]] -ne $null -and $CalendarPermissions[$Props[$Script:PidTagScheduleInfoDelegateNamesW][$i]].AccessRights -ne $null) {
                    $DelegateUser.Permissions.CalendarFolderPermissionLevel = [String]$CalendarPermissions[$Props[$Script:PidTagScheduleInfoDelegateNamesW][$i]].AccessRights
                    $sharingFlags = $CalendarPermissions[$Props[$Script:PidTagScheduleInfoDelegateNamesW][$i]].SharingPermissionFlags
                    if ($sharingFlags -and ($sharingFlags -is [System.Collections.IEnumerable])) {
                        $DelegateUser.ReceiveCopiesOfMeetingMessages = $sharingFlags.Contains("Delegate")
                    }
                    else {
                        $DelegateUser.ReceiveCopiesOfMeetingMessages = $false
                    }
                }
                if ($InboxPermissions.ContainsKey($Props[$Script:PidTagScheduleInfoDelegateNamesW][$i]) -and $InboxPermissions[$Props[$Script:PidTagScheduleInfoDelegateNamesW][$i]] -ne $null -and $InboxPermissions[$Props[$Script:PidTagScheduleInfoDelegateNamesW][$i]].AccessRights -ne $null) {
                    $DelegateUser.Permissions.InboxFolderPermissionLevel = [String]$InboxPermissions[$Props[$Script:PidTagScheduleInfoDelegateNamesW][$i]].AccessRights
                }
                if ($TasksPermissions.ContainsKey($Props[$Script:PidTagScheduleInfoDelegateNamesW][$i]) -and $TasksPermissions[$Props[$Script:PidTagScheduleInfoDelegateNamesW][$i]] -ne $null -and $TasksPermissions[$Props[$Script:PidTagScheduleInfoDelegateNamesW][$i]].AccessRights -ne $null) {
                    $DelegateUser.Permissions.TasksFolderPermissionLevel = [String]$TasksPermissions[$Props[$Script:PidTagScheduleInfoDelegateNamesW][$i]].AccessRights
                }
                if ($ContactsPermissions.ContainsKey($Props[$Script:PidTagScheduleInfoDelegateNamesW][$i]) -and $ContactsPermissions[$Props[$Script:PidTagScheduleInfoDelegateNamesW][$i]] -ne $null -and $ContactsPermissions[$Props[$Script:PidTagScheduleInfoDelegateNamesW][$i]].AccessRights -ne $null) {
                    $DelegateUser.Permissions.ContactsFolderPermissionLevel = [String]$ContactsPermissions[$Props[$Script:PidTagScheduleInfoDelegateNamesW][$i]].AccessRights
                }
                if ($NotesPermissions.ContainsKey($Props[$Script:PidTagScheduleInfoDelegateNamesW][$i]) -and $NotesPermissions[$Props[$Script:PidTagScheduleInfoDelegateNamesW][$i]] -ne $null -and $NotesPermissions[$Props[$Script:PidTagScheduleInfoDelegateNamesW][$i]].AccessRights -ne $null) {
                    $DelegateUser.Permissions.NotesFolderPermissionLevel = [String]$NotesPermissions[$Props[$Script:PidTagScheduleInfoDelegateNamesW][$i]].AccessRights
                }
                if ($JournalPermissions.ContainsKey($Props[$Script:PidTagScheduleInfoDelegateNamesW][$i]) -and $JournalPermissions[$Props[$Script:PidTagScheduleInfoDelegateNamesW][$i]] -ne $null -and $JournalPermissions[$Props[$Script:PidTagScheduleInfoDelegateNamesW][$i]].AccessRights -ne $null) {
                    $DelegateUser.Permissions.JournalFolderPermissionLevel = [String]$JournalPermissions[$Props[$Script:PidTagScheduleInfoDelegateNamesW][$i]].AccessRights
                }        
                # Add the delegate to the array
                $DelegateObjects.Delegates += $DelegateUser
            }
        }
        if ($Msg.SingleValueExtendedProperties -and $Msg.SingleValueExtendedProperties.Count -ge 2) {
            $prop0 = $Msg.SingleValueExtendedProperties[0]
            $prop1 = $Msg.SingleValueExtendedProperties[1]
      
            if ($prop0 -and $prop1 -and $prop0.Value -ne $null -and $prop1.Value -ne $null) {
                $wantsInfo = [boolean]::Parse($prop0.Value)
                $wantsCopy = [boolean]::Parse($prop1.Value)
        
                if ($wantsCopy -and $wantsInfo) {
                    $DelegateConfig = "My Delegates only but send a copy"
                }
                elseif ($wantsCopy) {
                    $DelegateConfig = "My Delegates and Me"
                }
                else {
                    $DelegateConfig = "My Delegates only"
                }
            }
            else {
                $DelegateConfig = "Delegate configuration properties not found."
            }
        }
        else {
            $DelegateConfig = "Delegate configuration properties not found."
        }
        $DelegateObjects.DeliverMeetingRequests = $DelegateConfig
    
        return $DelegateObjects
    }
}

$Script:PidTagScheduleInfoDelegateEntryIds = "BinaryArray 0x6845"
$Script:PidTagScheduleInfoDelegatorWantsInfo = "Boolean 0x684B"
$Script:PidTagScheduleInfoDelegatorWantsCopy = "Boolean 0x6842"
$Script:PidTagScheduleInfoDontMailDelegates = "Boolean 0x6843"
$Script:PidTagScheduleInfoDelegateNamesW = "StringArray 0x684A"
$Script:PidTagDelegateFlags = "IntegerArray 0x686B"