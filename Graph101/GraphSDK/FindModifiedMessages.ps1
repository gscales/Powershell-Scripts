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
                        Add-Member -InputObject $Item -NotePropertyName "lastModifiedUser" -NotePropertyValue $Prop.Value.Replace(" ", "\") -Force
                    }
                    "Integer 0x405A" {
                        Add-Member -InputObject $Item -NotePropertyName "lastModifiedFlags" -NotePropertyValue $Prop.Value.Replace(" ", "\") -Force
                    }
                    "String 0x66b5"{
                        Add-Member -InputObject $Item -NotePropertyName "FolderPath" -NotePropertyValue $Prop.Value.Replace(" ","\") -Force
                    }
                    "Binary 0x348A"{                            
                        Add-Member  -InputObject $Item -NotePropertyName "LastActiveParentEntryId" -NotePropertyValue ([System.BitConverter]::ToString([Convert]::FromBase64String($Prop.Value)).Replace("-",""))
                    }
                    "Binary 0xfff" {
                        Add-Member  -InputObject $Item -NotePropertyName "PR_ENTRYID" -NotePropertyValue ([System.BitConverter]::ToString([Convert]::FromBase64String($Prop.Value)).Replace("-", ""))
                    }
                }
            }
        }
    }
}

function Find-ModifiedMessages {	
    [CmdletBinding()] 
    param (
        [Parameter(Position = 1, Mandatory = $false)]
        [String]
        $ApplicationClientId,
        [Parameter(Position = 2, Mandatory = $false)]
        [String]
        $ApplicationClientSecret,
        [Parameter(Position = 3, Mandatory = $false)]
        [String]
        $TenantId,
        [Parameter(Position = 4, Mandatory = $false)]
        [String]
        $MailboxName,
        [Parameter(Position = 5, Mandatory = $false)]
        [datetime]
        $StartTime,
        [Parameter(Position = 6, Mandatory = $false)]
        [datetime]
        $Endtime
    )

    process {
        # Convert the Client Secret to a Secure String
        $SecureClientSecret = ConvertTo-SecureString -String $ApplicationClientSecret -AsPlainText -Force

        # Create a PSCredential Object Using the Client ID and Secure Client Secret
        $ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ApplicationClientId, $SecureClientSecret
        # Connect to Microsoft Graph Using the Tenant ID and Client Secret Credential
        Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential

        $Filter = "(lastModifiedDateTime ge $($StartTime.ToUniversalTime().ToString("yyyy-MM-ddThh:mm:ssZ"))) and (lastModifiedDateTime le $($Endtime.ToUniversalTime().ToString("yyyy-MM-ddThh:mm:ssZ")))"

        $rptCollection = @()

        $Folders = Invoke-EnumerateMailBoxFolders -MailboxName $MailboxName -WellKnownFolder msgFolderRoot
        $LapFidfldIndex = @{};
        $FolderIndex = New-Object hashtable
        foreach ($folder in $Folders) {
            $FolderIndex.Add($folder.Id, $folder.FolderPath)
            $laFid = $folder.PR_ENTRYID.substring(44,44)
            $LapFidfldIndex.Add($laFid,$folder);
        }
        Write-Verbose $FolderIndex.Values
        Get-MgUserMessage -PageSize 999 -All -UserId $MailboxName -ExpandProperty "singleValueExtendedProperties(`$filter=id eq 'String 0x3FFA' or id eq 'Integer 0x405A' or id eq 'Binary 0x348A')" -Filter $Filter -Select id, parentFolderId, lastModifiedDateTime, ReceivedDateTime, createdDateTime, Subject | ForEach-Object {
            $rptObj = "" | Select MailboxName, FolderPath, Subject, ReceivedDateTime, CreatedDateTime, lastModifiedDateTime, lastModifiedUser, lastModifiedFlags, LastActiveParentEntryId, LastActiveParentFolderPath
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
            if($_.LastActiveParentEntryId){
                    if($LapFidfldIndex.ContainsKey($_.LastActiveParentEntryId)){
                        $rptObj.LastActiveParentFolderPath = $LapFidfldIndex[$_.LastActiveParentEntryId].FolderPath 
                    }
            }
            $rptObj.lastActiveParentEntryId = $_.LastActiveParentEntryId
            $rptCollection += $rptObj    
        }
        return $rptCollection 
    }
}

