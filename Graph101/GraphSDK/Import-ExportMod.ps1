function Invoke-ListMailboxFolderItems{
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)] [String]$MailboxId,
        [Parameter(Position = 1, Mandatory = $true)] [String]$MailFolderId,
        [Parameter(Position = 2, Mandatory = $false)] [String]$ItemCount=10,
        [Parameter(Position = 3, Mandatory = $false)] [String]$Filter
    )
    Process{
        $ItemsProcessed = 0
        $TopVal = 999
        $RequestURL = "https://graph.microsoft.com/beta/admin/exchange/mailboxes/$MailboxId/folders/$MailFolderId/items?`$top=$TopVal&`$expand=singleValueExtendedProperties(`$filter=(id eq 'String 0x0037') or (id eq 'SystemTime 0x0E06'))"
        if($filter){
            $RequestURL += "&`$filter=$Filter"
        }
        do {
            $Results = Invoke-MgGraphRequest -Method Get -Uri $RequestURL  -Headers @{"Prefer"="IdType=`"ImmutableId`""}     
            $RequestURL  = $null
            if($Results){
                $RequestURL = $Results.'@odata.nextlink'
                if($RequestURL){
                    Write-Verbose $RequestURL
                }else{
                    Write-Verbose "No more pages"
                }                
                foreach($itemResult in $Results.Value){
                    write-verbose("Processing " + $itemResult.id)
                    Expand-ExtendedProperties -Item $itemResult
                    Write-Output ([PSCustomObject]$itemResult)
                    $ItemsProcessed++                                            
                    if($ItemCount -gt 0 -band $ItemsProcessed -ge $ItemCount){
                       $RequestURL = $null
                       break
                    }
                    $Results = $null     
                }  
            }
        } until ([String]::IsNullOrEmpty(($RequestURL))) 
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
        $Items = @()        
        $ImportPost = @{}
        $ImportPost.Add("FolderId",$FolderId)
        $ImportPost.Add("Mode","create")
        $ImportPost.Add("Data",[Convert]::ToBase64String([IO.File]::ReadAllBytes($FileName)))
        $Items += $ItemPost
        $result = Invoke-GraphRequest -Uri $ImportURL -Method Post -Body (ConvertTo-json $Items -Depth 10)
        Write-Verbose $result.StatusCode
        
    }
}

function Invoke-ImportItemFromDirectory{
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)] [string]$ImportURL,
        [Parameter(Position = 1, Mandatory = $true)] [string]$FolderName,
        [Parameter(Position = 2, Mandatory = $true)] [string]$FolderId
    )
    Process{
        $StartTime = Get-Date
        $report = "" | select NumberOfItems, TotalSize, ElapsedSeconds, Speed, Rate, ErrorCount, Errors
        $report.NumberOfItems = 0 
        $report.TotalSize = 0 
        $report.ElapsedSeconds = 0 
        $report.Speed = 0 
        $report.ErrorCount = 0 
        $report.Errors = @()
        $webSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        $BytesImported = 0;
        $stopwatch =  [system.diagnostics.stopwatch]::StartNew()
        Get-ChildItem $FolderName  -Filter *.fts | 
        Foreach-Object {
            Write-Verbose $_.FullName
            Invoke-UploadItem -FileName $_.FullName -FolderId $FolderId -ImportURL $ImportURL -webSession $webSession -Report $report -RetryCount 0
            $BytesImported += $_.Length
            $report.NumberOfItems++
            $report.TotalSize += $_.Length
            Write-Verbose ($BytesImported)
            Write-Verbose ($stopwatch.Elapsed.TotalSeconds)       
        }  
        $EndTime = Get-Date  
        $span = New-TimeSpan -Start $StartTime -end $EndTime    
        $report.ElapsedSeconds = $span.TotalSeconds
        $UploadSize = [Math]::Round(($report.TotalSize / 1Mb),2)
        $UploadTimeSec = $span.TotalSeconds
        $UploadSpeed = [Math]::Round((($UploadSize / $UploadTimeSec) * 8),2)
        $timeInHours = $UploadTimeSec / 3600
        $uploadRateMBPerHour = [Math]::Round(($UploadSize / $timeInHours),2)
        $report.Rate = "Upload Rate: $uploadRateMBPerHour MB per hour"
        $report.Speed = $UploadSpeed 
        return $report
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
                foreach($ExportedItem in $ExportItemsResult.value){
                    Write-Verbose ("Exported " + $ExportedItem.itemId)
                    $fileName = $ExportPath + [Guid]::NewGuid() + ".fts"
                    [IO.File]::WriteAllBytes($fileName, ([Convert]::FromBase64String($ExportedItem.data))) 
                }
                $ExportBatch = @()
            }
        }
        if($ExportBatch.Count -gt 0){
            $ExportHash = @{}
            $ExportHash.Add("itemIds", $ExportBatch)
            $ExportItemsResult = Invoke-GraphRequest -Uri $ExportUrl -Method Post -Body (ConvertTo-json $ExportHash -Depth 10)
            foreach($ExportedItem in $ExportItemsResult.value){
                $fileName = $ExportPath + [Guid]::NewGuid() + ".fts"
                [IO.File]::WriteAllBytes($fileName, ([Convert]::FromBase64String($ExportedItem.data))) 
            }
            $ExportBatch = @()
        }
    }
}

function Invoke-GetMailboxFolderFromPath {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string]
        $FolderPath,
		
        [Parameter(Position = 1, Mandatory = $true)] [PsObject]$MailboxId
    )

    process {
        
        $fldArray = $FolderPath.Split("\")
        #Loop through the Split Array and do a Search for each level of folder 
        $folderId = "MsgFolderRoot"
        for ($lint = 1; $lint -lt $fldArray.Length; $lint++) {
            #Perform search based on the displayname of each folder level
            $FolderName = $fldArray[$lint];
            $tfTargetFolder = Invoke-GetMailboxChildFolder -MailboxId $MailboxId -ParentFolderId $folderId -ChildFolderName $FolderName 
            if ($tfTargetFolder.displayname -match $FolderName) {
                $folderId = $tfTargetFolder.Id.ToString()
            }
            else {
                throw ("Folder Not found")
            }
        }
        return $tfTargetFolder 
    }
}


function Invoke-BatchExportItems{
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)] [PsObject]$Items,
        [Parameter(Position = 1, Mandatory = $true)] [PsObject]$MailboxId,
        [Parameter(Position = 2, Mandatory = $true)] [String]$ExportPath
    )
    Process{
        $stopwatch =  [system.diagnostics.stopwatch]::StartNew()
        $Report = "" | Select ItemsExported,ItemErrors,TotalItemSize
        $Report.ItemsExported = 0
        $Report.ItemErrors - 0
        $Report.TotalItemSize = 0
        $ExportUrl = "/admin/exchange/mailboxes/$MailboxId/exportItems"
        $ExportBatchs = @()
        $ExportBatch = @()
        foreach($Item in $Items){
            $ExportBatch += $Item.id
            $Report.TotalItemSize += $Item.size
            if($ExportBatch.Count -ge 20){
                $ExportHash = @{}
                $ExportHash.Add("itemIds", $ExportBatch)
                $ExportBatchs += $ExportHash
                $ExportBatch = @()
            }
        }
        if($ExportBatch.Count -gt 0){
            $ExportHash = @{}
            $ExportHash.Add("itemIds", $ExportBatch)
            $ExportBatchs += $ExportHash
            $ExportBatch = @()
        }
        if($ExportBatchs.Count -gt 0){
            $BatchRequestContent = @{}
            $BatchRequestContent.add("requests",@())
            $batchCount = 1
            foreach($ExportHashBatch in $ExportBatchs){
                $BatchEntry = @{}
                $BatchEntry.Add("id",[Int32]$batchCount)
                $BatchEntry.Add("method","POST")
                $BatchEntry.Add("url",$ExportUrl) 
                $BatchEntry.Add("body",$ExportHashBatch)           
                $BatchHeaders = @{
                    'Content-Type' =  "application/json"
                } 
                $BatchEntry.Add("headers",$BatchHeaders)
                $BatchRequestContent["requests"] += $BatchEntry
                if($batchCount -ge 20){
                    Invoke-BatchPost -BatchRequestContent $BatchRequestContent -ExportPath $ExportPath -Report $Report
                    $BatchRequestContent = @{}
                    $BatchRequestContent.add("requests",@())
                    $batchCount = 0
                }
                $batchCount++
            }
        }
        if($batchCount -ge 0){
            Invoke-BatchPost -BatchRequestContent $BatchRequestContent -ExportPath $ExportPath -Report $Report
            $BatchRequestContent = @{}
            $BatchRequestContent.add("requests",@())
            $batchCount = 0
        }
        Write-Verbose ($stopwatch.Elapsed.TotalMinutes)
        return $Report
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

function Invoke-GetMailboxFolder{
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)] [String]$MailboxId,
        [Parameter(Position = 1, Mandatory = $true)] [String]$FolderId="MsgFolderRoot"
        
    )
    Process{
        $RequestURL = "https://graph.microsoft.com/beta/admin/exchange/mailboxes/$MailboxId/folders/$FolderId"
        $Results = Invoke-MgGraphRequest -Method Get -Uri $RequestURL        
        return $Results
    }    
}

function Invoke-GetMailboxChildFolder{
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)] [String]$MailboxId,
        [Parameter(Position = 1, Mandatory = $true)] [String]$ParentFolderId,
        [Parameter(Position = 2, Mandatory = $true)] [String]$ChildFolderName
        
    )
    Process{
        $expandval = "`$expand=singleValueExtendedProperties(`$filter=(id eq 'String 0x66b5') or (id eq 'SystemTime 0x3007') or (id eq 'Integer 0x3601'))"
        $RequestURL = "https://graph.microsoft.com/beta/admin/exchange/mailboxes/$MailboxId/folders/$ParentFolderId/childFolders`?`$Filter = DisplayName eq '$ChildFolderName'&$expandval"
        $Results = Invoke-MgGraphRequest -Method Get -Uri $RequestURL  
        if($Results){
            foreach($itemResult in $Results.Value){
                write-verbose("Processing " + $itemResult.id)
                Expand-ExtendedProperties -Item $itemResult
                Write-Output ([PSCustomObject]$itemResult)      
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

function Invoke-BatchPost{
    param (
        [Parameter(Position = 1, Mandatory = $true)]
        [PSObject]
        $BatchRequestContent,
        [Parameter(Position = 2, Mandatory = $true)] 
        [String]$ExportPath,
        [Parameter(Position = 3, Mandatory = $true)] 
        [psobject]$Report

    )
    Process{
        $RequestURL = "https://graph.microsoft.com/beta/`$batch"
        $BatchResponse = Invoke-MgGraphRequest -Method POST -Uri $RequestURL -Body (ConvertTo-json  $BatchRequestContent -depth 10 -Compress)        
        if($BatchResponse.responses){
            foreach($Response in $BatchResponse.responses){                        
                if([Int32]$Response.status -eq 200){
                     Write-Verbose "Good Request"
                     $ExportedItemsResponse = $Response.body["value"]
                     foreach($ExportedItem in $ExportedItemsResponse){                        
                        if($ExportedItem.error){
                            Write-Verbose ("Error in Export" + $ExportedItem.error)
                            $Report.ItemErrors++
                        }else{
                            $fileName = $ExportPath + [Guid]::NewGuid() + ".fts"
                            [IO.File]::WriteAllBytes($fileName, ([Convert]::FromBase64String($ExportedItem.data)))
                            Write-Verbose("Exported " + $fileName)                            
                            $Report.ItemsExported++
                        } 
                    }
                }else{
                     Write-Verbose $Response.status
                }
           }
        } 
    }
}


function Invoke-UploadItem{
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string]
        $FolderId,
		
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $ImportURL,

        [Parameter(Position = 2, Mandatory = $true)]
        [String]
        $FileName,

        [Parameter(Position = 3, Mandatory = $true)]
        [Microsoft.PowerShell.Commands.WebRequestSession]
        $WebSession,

        [Parameter(Position = 4, Mandatory = $true)]
        [PSObject]
        $Report,

        [Parameter(Position = 5, Mandatory = $true)]
        [int]
        $RetryCount
    )
    Process{
        try{
            $Request = @{}
            $Request.Add("FolderId",$FolderId)
            $Request.Add("Mode","create")
            $Request.Add("Data", [Convert]::ToBase64String([IO.File]::ReadAllBytes($FileName)))
            $result = Invoke-WebRequest -method post -Uri $ImportURL -Body (ConvertTo-Json $Request -Depth 10) -ContentType "application/json" -WebSession $webSession  
            if($result.StatusCode -ne 200){
                $report.ErrorCount++                
            } 
        }catch{
            Write-Verbose("####Error")           
            $result = $_.Exception.Response
			Write-Verbose("Status Code : " + [int]$result.StatusCode) 
            $report.ErrorCount++ 
            if([int]$result.StatusCode -eq 429){
                if($result.Headers["Retry-After"]){
                    Write-Verbose ("Sleep for " + $result.Headers["Retry-After"])
                    Start-Sleep -Seconds $result.Headers["Retry-After"]
                    $RetryCount++
                    if($RetryCount -le 3){
                        Invoke-UploadItem -FolderId $FolderId -ImportURL $ImportURL -FileName $FileName -WebSession $WebSession -Report $Report -RetryCount $RetryCount
                    }else{
                        Write-Verbose "Retry Count Exceeded"
                    }
                    
                }
            }
            if([int]$result.StatusCode -eq 401){
                Write-Verbose("####Auth errors")  
            }         
        }

        
    }
        

}