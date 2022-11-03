function Export-MGPCalendar {
    [CmdletBinding()]
    param (       
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $MailboxName,
        [Parameter(Position = 3, Mandatory = $true)]
        [datetime]
        $StartTime,
        [Parameter(Position = 4, Mandatory = $true)]
        [datetime]
        $EndTime,
        [Parameter(Position = 5, Mandatory = $false)]
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
            'Prefer' = ("outlook.timezone=`"" + $TimeZone + "`"")
        } 
        $EndPoint = "https://graph.microsoft.com/v1.0/users"
        $RequestURL = $EndPoint + "('$MailboxName')/calendar/calendarview?startdatetime=" + $StartTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ") + "&enddatetime=" + $EndTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ") + "&`$Top=50"
        $RequestURL += "&`$expand=SingleValueExtendedProperties(`$filter=(Id%20eq%20'Integer%20%7B00062002-0000-0000-C000-000000000046%7D%20Id%200x8213') or (Id%20eq%20'Integer%20%7B00062002-0000-0000-C000-000000000046%7D%20Id%200x8217'))"
        do {
            Write-Verbose $RequestURL
            $ClientResult = Invoke-MgGraphRequest -Uri $RequestURL -Headers $headers
            Write-Verbose $ClientResult
            foreach ($CalendarEvent in $ClientResult.value) {
                Expand-ExtendedProperties -Item $CalendarEvent
                $rptObj = "" | Select StartTime, EndTime, Duration, Type, Subject, Location, Organizer, Attendees, Resources, AppointmentState, HasAttachments, IsReminderSet
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
                #$rptObj.Notes = $CalendarEvent.Body.content
                $rptCollection += $rptObj
            } 
              
            Write-Verbose ("Appointment Count : " + $rptCollection.count)  
            $RequestURL = $ClientResult.'@odata.nextLink'
        }while (![String]::IsNullOrEmpty($RequestURL)) 
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