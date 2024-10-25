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
        if($FolderPath -eq '\'){
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
        if($FolderPath){
            $searchRootFolder = Get-MailBoxFolderFromPath -MailboxName $MailboxName -FolderPath $FolderPath
            Add-Member -InputObject $searchRootFolder -NotePropertyName "FolderPath" -NotePropertyValue $FolderPath
        }
        if($WellKnownFolder){
            $searchRootFolder = Get-MgUserMailFolder -UserId $MailboxName -MailFolderId $WellKnownFolder             
        }
        if($returnSearchRoot){               
            $Script:Mailboxfolders += $searchRootFolder

        }      
        if($searchRootFolder){
            Invoke-EnumerateChildMailFolders -MailboxName $MailboxName -folderId $searchRootFolder.id
        }      
        return $Script:Mailboxfolders 
    }
}

function Invoke-EnumerateChildMailFolders{
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
        $childFolders = Get-MgUserMailFolderChildFolder -UserId $MailboxName -MailFolderId $folderId -All -ExpandProperty "singleValueExtendedProperties(`$filter=id eq 'String 0x66b5')"
        Write-Verbose ("Returned " + $childFolders.Count)
        foreach($childfolder in $childFolders){
            Expand-ExtendedProperties -Item $childfolder
            $Script:Mailboxfolders += $childfolder
            if($childfolder.ChildFolderCount -gt 0){
                Invoke-EnumerateChildMailFolders -MailboxName $MailboxName -folderId $childfolder.id
            }
        }
    }
}

function Expand-ExtendedProperties
{
	[CmdletBinding()] 
    param (
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$Item
	)
	
 	process
	{
		if ($Item.singleValueExtendedProperties -ne $null)
		{
			foreach ($Prop in $Item.singleValueExtendedProperties)
			{
				Switch ($Prop.Id)
				{
                    "String 0x66b5"{
                          Add-Member -InputObject $Item -NotePropertyName "FolderPath" -NotePropertyValue $Prop.Value.Replace(" ","\") -Force
                    }
                }
            }
        }
    }
}


# Define the Application (Client) ID and Secret
$ApplicationClientId = '12928ba2-2b75-4d04-9916-xxxxx' # Application (Client) ID
$ApplicationClientSecret = 'xxxxxx' # Application Secret Value
$TenantId = '1c3a18bf-da31-4f6c-a404-xxxx' # Tenant ID

# Convert the Client Secret to a Secure String
$SecureClientSecret = ConvertTo-SecureString -String $ApplicationClientSecret -AsPlainText -Force

# Create a PSCredential Object Using the Client ID and Secure Client Secret
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ApplicationClientId, $SecureClientSecret
# Connect to Microsoft Graph Using the Tenant ID and Client Secret Credential
Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential

$MailboxName = "gscales@datarumble.com"

$StartTime = (Get-Date).AddDays(-14).ToUniversalTime().ToString("yyyy-MM-ddThh:mm:ssZ")
$Endtime = (Get-Date).AddDays(-7).ToUniversalTime().ToString("yyyy-MM-ddThh:mm:ssZ")
$Filter = "(lastModifiedDateTime ge $StartTime) and (lastModifiedDateTime le $Endtime)"

$rptCollection = @()

$Folders = Invoke-EnumerateMailBoxFolders -MailboxName $MailboxName -WellKnownFolder msgFolderRoot
$FolderIndex = New-Object hashtable
foreach($folder in $Folders){
    $FolderIndex.Add($folder.Id,$folder.FolderPath)
}

Get-MgUserMessage -PageSize 999 -All -UserId $MailboxName -ExpandProperty "singleValueExtendedProperties(`$filter=id eq 'String 0x3FFA')" -Filter $Filter -Select id,parentFolderId,lastModifiedDateTime,ReceivedDateTime,createdDateTime,Subject | ForEach-Object{
    $rptObj = "" | Select MailboxName,FolderPath,Subject,ReceivedDateTime,CreatedDateTime,lastModifiedDateTime,lastModifiedUser
    $rptObj.MailboxName = $MailboxName
    if($FolderIndex.ContainsKey($_.parentFolderId)){
        $rptObj.FolderPath = $FolderIndex[$_.parentFolderId]
    }    
    $rptObj.Subject = $_.Subject
    $rptObj.lastModifiedDateTime = $_.lastModifiedDateTime
    if($_.singleValueExtendedProperties -ne $null){
        $rptObj.lastModifiedUser = $_.singleValueExtendedProperties[0].Value
    }    
    $rptObj.ReceivedDateTime = $_.ReceivedDateTime
    $rptObj.CreatedDateTime = $_.createdDateTime
    $rptCollection += $rptObj    
}
$rptCollection 