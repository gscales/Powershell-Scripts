function Invoke-ListMailboxFolderItems{
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)] [String]$MailboxId,
        [Parameter(Position = 1, Mandatory = $true)] [String]$MailFolderId,
        [Parameter(Position = 2, Mandatory = $false)] [String]$ItemCount=10
    )
    Process{
        $ItemsProcessed = 0
        $TopVal = 999
        if($ItemCount -lt 999){
            $TopVal = $ItemCount
        }
        $RequestURL = "https://graph.microsoft.com/beta/admin/exchange/mailboxes/$MailboxId/folders/$MailFolderId/items?`$top=$TopVal&`$expand=singleValueExtendedProperties(`$filter=(id eq 'String 0x0037') or (id eq 'SystemTime 0x0E06'))"
        do {
            $Results = Invoke-MgGraphRequest -Method Get -Uri $RequestURL        
            $RequestURL  = $null
            if($Results){
                foreach($itemResult in $Results.Value){
                    write-verbose("Processing " + $itemResult.id)
                    Expand-ExtendedProperties -Item $itemResult
                    Write-Output ([PSCustomObject]$itemResult)
                     $ItemsProcessed++
                    if($ItemCount -gt 0 -band $ItemsProcessed -gt $ItemCount){
                         break
                    }                           
                    $RequestURL = $Results.'@odata.nextlink'
                    if($ItemCount -gt 0 -band $ItemsProcessed -gt $ItemCount){
                       $RequestURL = $null
                       break
                    }
                    $Results = $null     
                }  
            }
        } until (!($RequestURL)) 
    }    
}

function Invoke-ImportMailboxItem{
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)] [string]$ImportURL,
        [Parameter(Position = 1, Mandatory = $true)] [string]$FileName,
        [Parameter(Position = 2, Mandatory = $true)] [string]$FolderId
    )
    Process{
        $ImportPost = @{}
        $ImportPost.Add("FolderId",$FolderId)
        $ImportPost.Add("Mode","create")
        $ImportPost.Add("Data",[Convert]::ToBase64String([IO.File]::ReadAllBytes($FileName)))
        $CreateImportSession = Invoke-GraphRequest -Uri $ImportURL -Method Post -Body (ConvertTo-json $ImportPost -Depth 10)
        return $CreateImportSession
    }
}

function Invoke-CreateImportSession{
    [CmdletBinding()]
    param (
        [Parameter(Position = 1, Mandatory = $true)] [PsObject]$MailboxId
    )
    Process{
        $RequestURL = "https://graph.microsoft.com/beta/admin/exchange/mailboxes/$MailboxId/createImportSession"
        $CreateImportSession = Invoke-GraphRequest -Uri $RequestURL -Method Post
        return $CreateImportSession
    }
}

function Invoke-ExportItems{
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)] [PsObject]$Items,
        [Parameter(Position = 1, Mandatory = $true)] [PsObject]$MailboxId,
        [Parameter(Position = 2, Mandatory = $true)] [String]$ExportPath
    )
    Process{
        $ExportUrl = "https://graph.microsoft.com/beta/admin/exchange/mailboxes/$MailboxId/exportItems"
        $ExportBatch = @()
        foreach($Item in $Items){
            $ExportBatch += $Item.id
            if($ExportBatch.Count -ge 20){
                $ExportHash = @{}
                $ExportHash.Add("itemIds", $ExportBatch)
                $ExportItemsResult = Invoke-GraphRequest -Uri $ExportUrl -Method Post -Body (ConvertTo-json $ExportHash -Depth 10)
                foreach($ExportedItem in $ExportItemsResult){
                    $fileName = $ExportPath + [Guid]::NewGuid() + ".fts"
                    [IO.File]::WriteAllBytes($fileName, ([Convert]::FromBase64String($ExportedItem.data))) 
                }
                $ExportBatch = @()
            }
        }
        if($ExportBatch.Count -ge 0){
            $ExportHash = @{}
            $ExportHash.Add("itemIds", $ExportBatch)
            $ExportItemsResult = Invoke-GraphRequest -Uri $ExportUrl -Method Post -Body (ConvertTo-json $ExportHash -Depth 10)
            foreach($ExportedItem in $ExportItemsResult.Value){
                $fileName = $ExportPath + [Guid]::NewGuid() + ".fts"
                [IO.File]::WriteAllBytes($fileName, ([Convert]::FromBase64String($ExportedItem.data))) 
            }
            $ExportBatch = @()
        }
    }
}

function Invoke-GetMailboxSettings{
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)] [String]$Upn
    )
    Process{
       $RequestURL = "https://graph.microsoft.com/beta/users/$Upn/settings/exchange"
       $Results = Invoke-MgGraphRequest -Method Get -Uri $RequestURL
       return $Results
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
                write-verbose("Processing " + $Prop.Id)
				Switch ($Prop.Id)
				{
                    "String 0x37"{
                        $Item.Add("subject",$Prop.Value)
                    }
                    "SystemTime 0xE06"{
                        $Item.Add("receivedDateTime",$Prop.Value)
                    }
                    "String 0x66b5"{
                        $fpath = Invoke-ConvertToStringFromExchange($Prop.Value)
                        $Item.Add("FolderPath",$fpath)
                    }
                    "SystemTime 0x3007"{
                        $Item.Add("createdDateTime",$Prop.Value)
                    }
                    "Integer 0x3601"{
                        $Item.Add("FolderType",$Prop.Value)
                    }
                }
            }
        }
    }
}

function Invoke-GetMailboxFolders{
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)] [String]$MailboxId,
        [Parameter(Position = 1, Mandatory = $true)] [String]$RootFolder="MsgFolderRoot",
        [Parameter(Position = 2, Mandatory = $false)] [String]$FolderCount=1000
    )
    Process{
        $Script:Mailboxfolders = @()
        $ItemsProcessed = 0
        $TopVal = 999
        if($FolderCount -lt 999){
            $TopVal = $FolderCount
        }
        $RequestURL = "https://graph.microsoft.com/beta/admin/exchange/mailboxes/$MailboxId/folders?`$Top=$TopVal&`$count=true"
        $Results = Invoke-MgGraphRequest -Method Get -Uri $RequestURL        
        $RequestURL  = $null
        do {
            if($Results){
                foreach($itemResult in $Results.Value){
                    write-verbose("Processing " + $itemResult.id)
                    Expand-ExtendedProperties -Item $itemResult
                    Write-Output ([PSCustomObject]$itemResult)
                    $ItemsProcessed++
                    if($FolderCount -gt 0 -band $ItemsProcessed -gt $FolderCount){
                         break
                    }                           
                    $RequestURL = $Results.'@odata.nextlink'
                    if($FolderCount -gt 0 -band $ItemsProcessed -gt $FolderCount){
                       $RequestURL = $null
                       break
                    }
                    $Results = $null     
                }  
            }
        } until (!($RequestURL)) 
    }    
}


function Invoke-EnumerateChildMailFolders{
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $false)]
        [string]
        $folderId = "msgfolderroot",
		
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $MailboxId
    )

    process {
        $RequestURL = "https://graph.microsoft.com/beta/admin/exchange/mailboxes/$MailboxId/folders/$folderId/childFolders?`$Top=999&`$count=true&`$expand=singleValueExtendedProperties(`$filter=(id eq 'String 0x66b5') or (id eq 'SystemTime 0x3007') or (id eq 'Integer 0x3601'))&includeHiddenFolders=true"
        do {
            $Results = Invoke-MgGraphRequest -Method Get -Uri $RequestURL        
            $RequestURL  = $null
            if($Results){
                foreach($folder in $Results.Value){
                    write-verbose("Processing " + $folder.id)
                    Expand-ExtendedProperties -Item $folder
                    Write-Output ([PSCustomObject]$folder)
                    if($folder.ChildFolderCount -gt 0){
                        Invoke-EnumerateChildMailFolders -MailboxId $MailboxId -folderId $folder.id
                    }                       
                    $RequestURL = $Results.'@odata.nextlink'
                    $Results = $null     
                }  
            }
        } until (!($RequestURL)) 
    }
}

function Invoke-CreateNewFolder{
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string]
        $ParentFolderId,
		
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $MailboxId,

        [Parameter(Position = 2, Mandatory = $true)]
        [String]
        $FolderName,

        [Parameter(Position = 3, Mandatory = $true)]
        [String]
        $FolderType
    )
    Process{
        $RequestURL = "https://graph.microsoft.com/beta/admin/exchange/mailboxes/$MailboxId/folders/$ParentFolderId/childFolders"
        $NewFolderRequest = @{
            'displayName' = $FolderName
            'type' = $FolderType
        }
        $Results = Invoke-MgGraphRequest -Method POST -Uri $RequestURL -Body (ConvertTo-Json $NewFolderRequest -depth 10)
        return $Results
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