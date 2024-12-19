function Get-MailBoxFolderFromPath {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string]
        $FolderPath,
		
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $MailboxName
    )

    process {
        if ($FolderPath -eq '\') {
            return Get-MgUserMailFolder -UserId $MailboxName -MailFolderId msgFolderRoot 
        }
        $fldArray = $FolderPath.Split("\")
        #Loop through the Split Array and do a Search for each level of folder 
        $folderId = "MsgFolderRoot"
        for ($lint = 1; $lint -lt $fldArray.Length; $lint++) {
            #Perform search based on the displayname of each folder level
            $FolderName = $fldArray[$lint];
            $tfTargetFolder = Get-MgUserMailFolderChildFolder -UserId $MailboxName -Filter "DisplayName eq '$FolderName'" -MailFolderId $folderId -All 
            if ($tfTargetFolder.displayname -eq $FolderName) {
                $folderId = $tfTargetFolder.Id.ToString()
            }
            else {
                throw ("Folder Not found")
            }
        }
        return $tfTargetFolder 
    }
}


function Invoke-EnumerateMailBoxFolders {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $false)]
        [string]
        $FolderPath,

        [Parameter(Position = 1, Mandatory = $false)]
        [string]
        $WellKnownFolder,
		
        [Parameter(Position = 2, Mandatory = $true)]
        [String]
        $MailboxName,

        [Parameter(Position = 3, Mandatory = $false)]
        [switch]
        $returnSearchRoot
    )

    process {
        $Script:Mailboxfolders = @()
        if ($FolderPath) {
            $searchRootFolder = Get-MailBoxFolderFromPath -MailboxName $MailboxName -FolderPath $FolderPath
            Add-Member -InputObject $searchRootFolder -NotePropertyName "FolderPath" -NotePropertyValue $FolderPath
        }
        if ($WellKnownFolder) {
            $searchRootFolder = Get-MgUserMailFolder -UserId $MailboxName -MailFolderId $WellKnownFolder             
        }
        if ($returnSearchRoot) {               
            $Script:Mailboxfolders += $searchRootFolder

        }      
        if ($searchRootFolder) {
            Invoke-EnumerateChildMailFolders -MailboxName $MailboxName -folderId $searchRootFolder.id
        }      
        return $Script:Mailboxfolders 
    }
}

function Invoke-EnumerateChildMailFolders {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string]
        $folderId,
		
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $MailboxName
    )

    process {
        $childFolders = Get-MgUserMailFolderChildFolder -UserId $MailboxName -MailFolderId $folderId -All -ExpandProperty "singleValueExtendedProperties(`$filter=id eq 'String 0x66b5' or id eq 'Binary 0xfff')"
        Write-Verbose ("Returned " + $childFolders.Count)
        foreach ($childfolder in $childFolders) {
            Expand-ExtendedProperties -Item $childfolder
            $Script:Mailboxfolders += $childfolder
            if ($childfolder.ChildFolderCount -gt 0) {
                Invoke-EnumerateChildMailFolders -MailboxName $MailboxName -folderId $childfolder.id
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
                    "String 0x3FFA" {
                        $muser = Invoke-ConvertToStringFromExchange($Prop.Value)
                        Add-Member -InputObject $Item -NotePropertyName "lastModifiedUser" -NotePropertyValue $muser -Force
                    }
                    "Integer 0x405A" {
                        Add-Member -InputObject $Item -NotePropertyName "lastModifiedFlags" -NotePropertyValue $Prop.Value.Replace(" ", "\") -Force
                    }
                    "String 0x66b5"{
                        $fpath = Invoke-ConvertToStringFromExchange($Prop.Value)
                        Add-Member -InputObject $Item -NotePropertyName "FolderPath" -NotePropertyValue $fpath -Force
                    }
                    "Binary 0x348A"{                            
                        Add-Member  -InputObject $Item -NotePropertyName "LastActiveParentEntryId" -NotePropertyValue ([System.BitConverter]::ToString([Convert]::FromBase64String($Prop.Value)).Replace("-",""))
                    }
                    "Binary 0xfff" {
                        Add-Member  -InputObject $Item -NotePropertyName "PR_ENTRYID" -NotePropertyValue ([System.BitConverter]::ToString([Convert]::FromBase64String($Prop.Value)).Replace("-", ""))
                    }
                    "SystemTime 0x39" {
                        Add-Member -InputObject $Item -NotePropertyName "PidTagClientSubmitTime" -NotePropertyValue ([DateTime]::Parse($Prop.Value)).ToUniversalTime()
                    }
                    "Integer 0x1081"{
                        Add-Member -InputObject $Item -NotePropertyName "PidTagLastVerbExecuted" -NotePropertyValue $Prop.Value -Force
                        $verbHash = Get-LASTVERBEXECUTEDHash;
                        if($verbHash.ContainsKey($Prop.Value)){
                            Add-Member -InputObject $Item -NotePropertyName "LastVerb" -NotePropertyValue $verbHash[$Prop.Value]
                        } 
                    }  
                    "SystemTime 0x1082" {
                        Add-Member -InputObject $Item -NotePropertyName "PidTagLastVerbExecutedTime" -NotePropertyValue ([DateTime]::Parse($Prop.Value)).ToUniversalTime()
                    }   
                }
            }
        }
    }
}

function Invoke-ConvertToStringFromExchange($ipInputString) { 
    $binarry = [Text.Encoding]::UTF8.GetBytes($ipInputString)  
    $hexArr = $binarry | ForEach-Object { $_.ToString("X2") }  
    $hexString = $hexArr -join ''  
    $hexString = $hexString.Replace("FEFF", "5C00")   
    $Val1Text = ""  
    for ($clInt = 0; $clInt -lt $hexString.length; $clInt++) {  
        $Val1Text = $Val1Text + [Convert]::ToString([Convert]::ToChar([Convert]::ToInt32($hexString.Substring($clInt, 2), 16)))  
        $clInt++  
    }  
    return $Val1Text  
} 

function Get-LASTVERBEXECUTEDHash(){
    $repHash = @{}
    $repHash.Add("0","open")
    $repHash.Add("102","ReplyToSender")
    $repHash.Add("103","ReplyToAll")
    $repHash.Add("104","Forward")
    $repHash.Add("105","Print")
    $repHash.Add("106","Save as")
    $repHash.Add("108","ReplyToFolder")
    $repHash.Add("500","Save")
    $repHash.Add("510","Properties")
    $repHash.Add("511","Followup")
    $repHash.Add("512","Accept")
    $repHash.Add("513","Tentative")
    $repHash.Add("514","Reject")
    $repHash.Add("515","Decline")
    $repHash.Add("516","Invite")
    $repHash.Add("517","Update")
    $repHash.Add("518","Cancel")
    $repHash.Add("519","SilentInvite")
    $repHash.Add("520","SilentCancel")
    $repHash.Add("521","RecallMessage")
    $repHash.Add("522","ForwardResponse")
    $repHash.Add("523","ForwardCancel")
    $repHash.Add("524","FollowupClear")
    $repHash.Add("525","ForwardAppointment")
    $repHash.Add("526","OpenResend")
    $repHash.Add("527","StatusReport")
    $repHash.Add("528","JournalOpen")
    $repHash.Add("529","JournalOpenLink")
    $repHash.Add("530","ComposeReplace")
    $repHash.Add("531","Edit")
    $repHash.Add("532","DeleteProcess")
    $repHash.Add("533","TentativeAppointmentTime")
    $repHash.Add("534","EditTemplate")
    $repHash.Add("535","FindInCalendar")
    $repHash.Add("536","ForwardAsFile")
    $repHash.Add("537","ChangeAttendees")
    $repHash.Add("538","RecalculateTitle")
    $repHash.Add("539","PropertyChange")
    $repHash.Add("540","ForwardAsVcal")
    $repHash.Add("541","ForwardAsIcal")
    $repHash.Add("542","ForwardAsBusinessCard")
    $repHash.Add("543","DeclineAppointmentTime")
    $repHash.Add("544","Process")
    $repHash.Add("545","OpenWithWord")
    $repHash.Add("546","OpenInstanceOfSeries")
    $repHash.Add("547","FilloutThisForm")
    $repHash.Add("548","FollowupDefault")
    $repHash.Add("549","ReplyWithMail")
    $repHash.Add("566","ToDoToday")
    $repHash.Add("567","ToDoTomorrow")
    $repHash.Add("568","ToDoThisWeek")
    $repHash.Add("569","ToDoNextWeek")
    $repHash.Add("570","ToDoThisMonth")
    $repHash.Add("571","ToDoNextMonth")
    $repHash.Add("572","ToDoNoDate")
    $repHash.Add("573","FollowupComplete")
    $repHash.Add("574","CopyToPostFolder")
    $repHash.Add("579","SeriesInvitationUpdateToPartialAttendeeList")
    $repHash.Add("580","SeriesCancellationUpdateToPartialAttendeeList")
    return $repHash
}

function Find-ModifiedMessages {	
    [CmdletBinding()] 
    param (
        $MailboxName,
        [Parameter(Position = 0, Mandatory = $true)]
        [datetime]
        $StartTime,
        [Parameter(Position = 1, Mandatory = $false)]
        [datetime]
        $Endtime,
        [Parameter(Position = 2, Mandatory = $false)]
        [String]
        $WellKnownFolder = "msgFolderRoot"
    )

    process {

        $Filter = "(lastModifiedDateTime ge $($StartTime.ToUniversalTime().ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ssZ"))) and (lastModifiedDateTime lt $($Endtime.ToUniversalTime().ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ssZ")))"

        $rptCollection = @()

        $Folders = Invoke-EnumerateMailBoxFolders -MailboxName $MailboxName -WellKnownFolder $WellKnownFolder
        $LapFidfldIndex = @{};
        $FolderIndex = New-Object hashtable
        foreach ($folder in $Folders) {
            $FolderIndex.Add($folder.Id, $folder.FolderPath)
            $laFid = $folder.PR_ENTRYID.substring(44,44)
            $LapFidfldIndex.Add($laFid,$folder);
        }
        foreach($folder in $Folders){
            Write-Verbose ("Processing " + $folder.FolderPath + " Item Count " + $folder.TotalItemCount)
            if($folder.TotalItemCount -gt 0){
                Get-MgUserMailFolderMessage -MailFolderId $folder.id -PageSize 999 -All -UserId $MailboxName -ExpandProperty "singleValueExtendedProperties(`$filter=id eq 'String 0x3FFA' or id eq 'Integer 0x405A' or id eq 'Binary 0x348A' or id eq 'SystemTime 0x39' or id eq 'Integer 0x1081' or id eq 'SystemTime 0x1082')" -Filter $Filter -Select id, isRead , parentFolderId, lastModifiedDateTime, ReceivedDateTime, createdDateTime, Subject | ForEach-Object {
                    $rptObj = "" | Select MailboxName, FolderPath, Subject, isRead , ReceivedDateTime, CreatedDateTime, lastModifiedDateTime, PidTagClientSubmitTime, lastModifiedUser, lastModifiedFlags, LastActiveParentEntryId, LastActiveParentFolderPath, PidTagLastVerbExecuted, PidTagLastVerbExecutedTime, LastVerb
                    $rptObj.MailboxName = $MailboxName
                    if ($FolderIndex.ContainsKey($_.parentFolderId)) {
                        $rptObj.FolderPath = $FolderIndex[$_.parentFolderId]
                    }  
                    Expand-ExtendedProperties -Item $_  
                    $rptObj.Subject = $_.Subject
                    $rptObj.lastModifiedDateTime = $_.lastModifiedDateTime
                    $rptObj.lastModifiedUser = $_.lastModifiedUser
                    $rptObj.ReceivedDateTime = $_.ReceivedDateTime
                    $rptObj.CreatedDateTime = $_.createdDateTime
                    $rptObj.lastModifiedFlags = $_.lastModifiedFlags
                    $rptObj.PidTagClientSubmitTime  = $_.PidTagClientSubmitTime
                    $rptObj.PidTagLastVerbExecuted = $_.PidTagLastVerbExecuted
                    $rptObj.LastVerb = $_.LastVerb
                    $rptObj.isRead = $_.isRead 
                    $rptObj.PidTagLastVerbExecutedTime = $_.PidTagLastVerbExecutedTime
                    if($_.LastActiveParentEntryId){
                            if($LapFidfldIndex.ContainsKey($_.LastActiveParentEntryId)){
                                $rptObj.LastActiveParentFolderPath = $LapFidfldIndex[$_.LastActiveParentEntryId].FolderPath 
                            }
                    }
                    $rptObj.lastActiveParentEntryId = $_.LastActiveParentEntryId
                    $rptCollection += $rptObj    
                }
            }else{
                write-Verbose "Skipping Empty Folder"
            }

        }
        return $rptCollection 
    }
}

