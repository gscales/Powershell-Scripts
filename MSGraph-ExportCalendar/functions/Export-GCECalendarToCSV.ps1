function Export-GCECalendarToCSV {
    [CmdletBinding()]
    <#
	.SYNOPSIS
		Exports Calendar events from  an Office365 Mailbox Calendar to a CSV file using the Microsoft Graph API
	
	.DESCRIPTION
		Exports Calendar events from  an Office365 Mailbox Calendar to a CSV file using the Microsoft Graph API
		

	
	.PARAMETER MailboxName
		A description of the MailboxName parameter.
	
	.PARAMETER ClientId
		ClientId for the Azure Applicaiton registration 
	
	.PARAMETER StartTime
	   Start Time to Start searching for Appointments to Export
	
	.PARAMETER EndTime
	   End Time to End searching for Appointments to Export
	
	.PARAMETER TimeZone
		Used to Outlook Prefer header (uses local by default)
	
	.PARAMETER FileName
		File to Export the Calendar Appointments to
	
	.EXAMPLE
        Export the last years Calendar appointments to a CSV
        Export-GCECalendarToCSV -MailboxName gscales@datarumble.com -StartTime (Get-Date).AddYears(-1) -EndTime (Get-Date) -FileName c:\export\lastyear.csv
	
	
#>
    param (   
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $MailboxName,
        [Parameter(Position = 2, Mandatory = $false)]
        [String]
        $ClientId,
        [Parameter(Position = 3, Mandatory = $true)]
        [datetime]
        $StartTime,
        [Parameter(Position = 4, Mandatory = $true)]
        [datetime]
        $EndTime,
        [Parameter(Position = 5, Mandatory = $false)]
        [String]
        $TimeZone,
        [Parameter(Position = 6, Mandatory = $true)]
        [String]
        $FileName,
        [Parameter(Position = 7, Mandatory = $false)]
        [String]
        $redirectURL
    )
    Begin {  
        $Events = Export-GCECalendar -MailboxName $MailboxName -ClientId $ClientId -StartTime $StartTime -EndTime $EndTime -TimeZone $TimeZone
        $Events | Export-Csv -NoTypeInformation -Path $FileName
        Write-Verbose("Exported to " + $FileName)
    }
}
function Export-GCECalendar {
    [CmdletBinding()]
    param (       
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $MailboxName,
        [Parameter(Position = 2, Mandatory = $false)]
        [String]
        $ClientId,
        [Parameter(Position = 3, Mandatory = $true)]
        [datetime]
        $StartTime,
        [Parameter(Position = 4, Mandatory = $true)]
        [datetime]
        $EndTime,
        [Parameter(Position = 5, Mandatory = $false)]
        [String]
        $TimeZone,
        [Parameter(Position = 6, Mandatory = $false)]
        [String]
        $redirectURL

        

    )
    Begin {  
        $rptCollection = @()      
        if ([String]::IsNullOrEmpty($ClientId)) {
            $ClientId = "c14a7c0c-9a2f-4985-b4bf-7228780a254c"
        }	
        if ([String]::IsNullOrEmpty($redirectURL)) {
            $redirectURL = "https://login.microsoftonline.com/common/oauth2/nativeclient"
        }		
        $adal = Join-Path $script:ModuleRoot "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
        $adalforms = Join-Path $script:ModuleRoot "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
        if ([System.IO.File]::Exists($adal)) { 
            Import-Module $adal -Force
        }
        if ([System.IO.File]::Exists($adalforms)) { 
            Import-Module $adalforms -Force
        }  
        $PromptBehavior = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters -ArgumentList Auto       
        $Context = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext("https://login.microsoftonline.com/common")
        $token = ($Context.AcquireTokenAsync("https://graph.microsoft.com", $ClientId , $redirectURL, $PromptBehavior)).Result
        $HttpClient = Get-HTTPClient -MailboxName $MailboxName -Token $token.AccessToken
        if ([String]::IsNullOrEmpty($TimeZone)) {
            $TimeZone = [TimeZoneInfo]::Local.Id;
        }
        $AppointmentState = @{0 = "None" ; 1 = "Meeting" ; 2 = "Received" ; 4 = "Canceled" ; }
        $HttpClient.DefaultRequestHeaders.Add("Prefer", ("outlook.timezone=`"" + $TimeZone + "`""))
        $EndPoint = "https://graph.microsoft.com/v1.0/users"
        $RequestURL = $EndPoint + "('$MailboxName')/calendar/calendarview?startdatetime=" + $StartTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ") + "&enddatetime=" + $EndTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ") + "&`$Top=500"
        $RequestURL += "&`$expand=SingleValueExtendedProperties(`$filter=(Id%20eq%20'Integer%20%7B00062002-0000-0000-C000-000000000046%7D%20Id%200x8213') or (Id%20eq%20'Integer%20%7B00062002-0000-0000-C000-000000000046%7D%20Id%200x8217'))"
        do {
            Write-Verbose $RequestURL
            $ClientResult = $HttpClient.GetAsync($RequestURL)
            if ($ClientResult.Result.StatusCode -ne [System.Net.HttpStatusCode]::OK) {
                Write-Error $ClientResult.Result   
                $RequestURL = "";             
            }
            else {
                Write-Verbose $ClientResult.Result
                $JSONOutput = ExpandPayload($ClientResult.Result.Content.ReadAsStringAsync().Result)
                foreach ($CalendarEvent in $JSONOutput.Value) {
                    Expand-ExtendedProperties -Item $CalendarEvent
                    $rptObj = "" | Select StartTime, EndTime, Duration, Type, Subject, Location, Organizer, Attendees, Resources, AppointmentState, Notes, HasAttachments, IsReminderSet
                    $rptObj.HasAttachments = $false;
                    $rptObj.IsReminderSet = $false
                    $rptObj.StartTime = ([DateTime]$CalendarEvent.Start.dateTime).ToString("yyyy-MM-dd HH:mm")  
                    $rptObj.EndTime = ([DateTime]$CalendarEvent.End.dateTime).ToString("yyyy-MM-dd HH:mm")  
                    $rptObj.Duration = $CalendarEvent.AppointmentDuration
                    $rptObj.Subject = $CalendarEvent.Subject   
                    $rptObj.Type = $CalendarEvent.type
                    $rptObj.Location = $CalendarEvent.Location.displayName
                    $rptObj.Organizer = $CalendarEvent.organizer.emailAddress.address
                    $aptStat = "";
                    $AppointmentState.Keys | where { $_ -band $CalendarEvent.AppointmentState } | foreach { $aptStat += $AppointmentState.Get_Item($_) + " " }
                    $rptObj.AppointmentState = $aptStat	
                    if ($CalendarEvent.hasAttachments) { $rptObj.HasAttachments = $CalendarEvent.hasAttachments }
                    if ($CalendarEvent.IsReminderSet) { $rptObj.IsReminderSet = $CalendarEvent.IsReminderSet }               
                    foreach ($attendee in $CalendarEvent.attendees) {
                        if ($attendee.type -eq "resource") {
                            $rptObj.Resources += $attendee.emailaddress.address + " " + $attendee.type + " " + $attendee.status.response + ";"
                        }
                        else {
                            $atn = $attendee.emailaddress.address + " " + $attendee.type + " " + $attendee.status.response + ";"
                            $rptObj.Attendees += $atn
                        }
                    }
                    $rptObj.Notes = $CalendarEvent.Body.content
                    $rptCollection += $rptObj
                } 
            }  
            Write-Verbose ("Appointment Count : " + $rptCollection.count)  
            $RequestURL = $JSONOutput.'@odata.nextLink'
        }while (![String]::IsNullOrEmpty($RequestURL)) 
        return $rptCollection
		
    }
}

function Get-HTTPClient {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string]
        $MailboxName,
        [Parameter(Position = 2, Mandatory = $true)]
        [string]
        $Token
    )
    process {
        Add-Type -AssemblyName System.Net.Http
        $handler = New-Object  System.Net.Http.HttpClientHandler
        $handler.CookieContainer = New-Object System.Net.CookieContainer
        $handler.AllowAutoRedirect = $true;
        $HttpClient = New-Object System.Net.Http.HttpClient($handler);
        $HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", $Token);
        $Header = New-Object System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json")
        $HttpClient.DefaultRequestHeaders.Accept.Add($Header);
        $HttpClient.Timeout = New-Object System.TimeSpan(0, 0, 90);
        $HttpClient.DefaultRequestHeaders.TransferEncodingChunked = $false
        if (!$HttpClient.DefaultRequestHeaders.Contains("X-AnchorMailbox")) {
            $HttpClient.DefaultRequestHeaders.Add("X-AnchorMailbox", $MailboxName);
        }
        $Header = New-Object System.Net.Http.Headers.ProductInfoHeaderValue("GraphRest", "1.1")
        $HttpClient.DefaultRequestHeaders.UserAgent.Add($Header);
        return $HttpClient
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
                    "Integer {00062002-0000-0000-c000-000000000046} Id 0x8213" {
                        Add-Member -InputObject $Item -NotePropertyName "AppointmentDuration" -NotePropertyValue $Prop.Value -Force
                    }
                    "Integer {00062002-0000-0000-c000-000000000046} Id 0x8217" {
                        Add-Member -InputObject $Item -NotePropertyName "AppointmentState" -NotePropertyValue $Prop.Value -Force
                    }                    
                }
            }
        }
    }
}
function ExpandPayload {
    [CmdletBinding()]
    Param (
        $response
    )
    if ($PSVersionTable.PSEdition -eq "Core") {
        ConvertFrom-JsonNewtonsoft $response
    }
    else {
        ## Start Code Attribution
        ## ExpandPayload function is the work of the following Authors and should remain with the function if copied into other scripts
        ## https://www.powershellgallery.com/profiles/chriswahl/
        ## End Code Attribution
        [void][System.Reflection.Assembly]::LoadWithPartialName('System.Web.Extensions')
        return ParseItem -jsonItem ((New-Object -TypeName System.Web.Script.Serialization.JavaScriptSerializer -Property @{
                    MaxJsonLength = [Int32]::MaxValue
                }).DeserializeObject($response))
    }
}

function ConvertFrom-JsonNewtonsoft {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true, ValueFromPipeline = $true)]$string) 
    ## Start Code Attribution
    ## ExpandPayload function is the work of the following Authors and should remain with the function if copied into other scripts
    ## https://www.powershellgallery.com/profiles/chriswahl/
    ## End Code Attribution
    $HandleDeserializationError = 
    {
        param ([object] $sender, [Newtonsoft.Json.Serialization.ErrorEventArgs] $errorArgs)
        $currentError = $errorArgs.ErrorContext.Error.Message
        write-warning $currentError
        $errorArgs.ErrorContext.Handled = $true
        
    }

    $settings = new-object "Newtonsoft.Json.JSonSerializerSettings"
    if ($ErrorActionPreference -eq "Ignore") {
        $settings.Error = $HandleDeserializationError
    }
    $obj = [Newtonsoft.Json.JsonConvert]::DeserializeObject($string, [Newtonsoft.Json.Linq.JObject], $settings)    

    return ConvertFrom-JObject $obj
}

function ConvertFrom-JObject($obj) {
    ## Start Code Attribution
    ## ExpandPayload function is the work of the following Authors and should remain with the function if copied into other scripts
    ## https://www.powershellgallery.com/profiles/chriswahl/
    ## End Code Attribution
    if ($obj -is [Newtonsoft.Json.Linq.JArray]) {
        $a = foreach ($entry in $obj.GetEnumerator()) {
            @(convertfrom-jobject $entry)
        }
        return $a
    }
    elseif ($obj -is [Newtonsoft.Json.Linq.JObject]) {
        $h = [ordered]@{ }
        foreach ($kvp in $obj.GetEnumerator()) {
            $val = convertfrom-jobject $kvp.value
            if ($kvp.value -is [Newtonsoft.Json.Linq.JArray]) { $val = @($val) }
            $h += @{ "$($kvp.key)" = $val }
        }
        return [pscustomobject]$h
    }
    elseif ($obj -is [Newtonsoft.Json.Linq.JValue]) {
        return $obj.Value
    }
    else {
        return $obj
    }
}

function ParseItem {
    [CmdletBinding()]
    Param (
        $JsonItem
    )
	
    if ($jsonItem.PSObject.TypeNames -match 'Array') {
        return ParseJsonArray -jsonArray ($jsonItem)
    }
    elseif ($jsonItem.PSObject.TypeNames -match 'Dictionary') {
        return ParseJsonObject -jsonObj ([HashTable]$jsonItem)
    }
    else {
        return $jsonItem
    }
}
function ParseJsonObject {
    [CmdletBinding()]
    Param (
        $jsonObj
    )
    ## Start Code Attribution
    ## ParseJsonObject function is the work of the following Authors and should remain with the function if copied into other scripts
    ## https://www.powershellgallery.com/profiles/chriswahl/
    ## End Code Attribution
    $result = New-Object -TypeName PSCustomObject
    foreach ($key in $jsonObj.Keys) {
        $item = $jsonObj[$key]
        if ($item) {
            $parsedItem = ParseItem -jsonItem $item
        }
        else {
            $parsedItem = $null
        }
        $result | Add-Member -MemberType NoteProperty -Name $key -Value $parsedItem
    }
    return $result
}
function ParseJsonArray {
    [CmdletBinding()]
    Param (
        $jsonArray
    )
    ## Start Code Attribution
    ## ParseJsonArray function is the work of the following Authors and should remain with the function if copied into other scripts
    ## https://www.powershellgallery.com/profiles/chriswahl/
    ## End Code Attribution
    $result = @()
    $jsonArray | ForEach-Object -Process {
        $result += , (ParseItem -jsonItem $_)
    }
    return $result
}

