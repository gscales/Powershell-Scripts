function Export-GSdkCalendarToCSV {
    [CmdletBinding()]
    <#
	.SYNOPSIS
		Exports Calendar events from  an Microsoft365 Mailbox Calendar to a CSV file using the Microsoft Graph API and the Microsoft Graph Powershell SDK
	
	.DESCRIPTION
		Exports Calendar events from  an Microsoft365 Mailbox Calendar to a CSV file using the Microsoft Graph API Microsoft Graph Powershell SDK. This script
        requires you make a connection to the Microsoft Graph first using any of the methods outlined in
         https://learn.microsoft.com/en-us/powershell/microsoftgraph/get-started?view=graph-powershell-1.0

		

	
	.PARAMETER MailboxName
		A description of the MailboxName parameter.
	
		
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
        Export-GSdkCalendarToCSV -MailboxName gscales@datarumble.com -StartTime (Get-Date).AddYears(-1) -EndTime (Get-Date) -FileName c:\export\lastyear.csv
       
	
	
#>
    param (   
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $MailboxName,
        [Parameter(Position = 2, Mandatory = $true)]
        [datetime]
        $StartTime,
        [Parameter(Position = 3, Mandatory = $true)]
        [datetime]
        $EndTime,
        [Parameter(Position = 4, Mandatory = $false)]
        [String]
        $TimeZone,
        [Parameter(Position = 5, Mandatory = $true)]
        [String]
        $FileName
    )
    Process {  
        $Events = Invoke-GSdkExportCalendar -MailboxName $MailboxName -StartTime $StartTime -EndTime $EndTime -TimeZone $TimeZone 
        $Events | Export-Csv -NoTypeInformation -Path $FileName
        Write-Verbose("Exported to " + $FileName)
    }
}
function Invoke-GSdkExportCalendar {
    [CmdletBinding()]
    param (       
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $MailboxName,
        [Parameter(Position = 2, Mandatory = $true)]
        [datetime]
        $StartTime,
        [Parameter(Position = 3, Mandatory = $true)]
        [datetime]
        $EndTime,
        [Parameter(Position = 4, Mandatory = $false)]
        [String]
        $TimeZone       

    )
    Begin {  
        $rptCollection = @()
        if ([String]::IsNullOrEmpty($TimeZone)) {
            $TimeZone = [TimeZoneInfo]::Local.Id;
        }
        $AppointmentState = @{0 = "None" ; 1 = "Meeting" ; 2 = "Received" ; 4 = "Canceled" ; }

        $headers = @{
            "Prefer" = "outlook.timezone=`"" + $TimeZone + "`", outlook.body-content-type='text'"
        }
        $queryStartTime = $StartTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        $queryEndTime = $EndTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

        $events = Get-MgUserCalendarView -UserId $MailboxName -Headers $headers -StartDateTime $queryStartTime -EndDateTime $queryEndTime -PageSize 500 -All -ExpandProperty "SingleValueExtendedProperties(`$filter=(Id eq 'Integer {00062002-0000-0000-C000-000000000046} Id 0x8213') or (Id eq 'Integer {00062002-0000-0000-C000-000000000046} Id 0x8217'))"
        foreach ($CalendarEvent in $events) {
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
   
        return $rptCollection
		
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


