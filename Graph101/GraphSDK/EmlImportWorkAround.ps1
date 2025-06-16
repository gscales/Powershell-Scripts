<#
.SYNOPSIS
 Import an EML into a Exchange Mailbox
 Requires a Powershell Graph context with the following scopes
    # Connect to Microsoft Graph

    $scopes = @("Mail.ReadWrite", "MailboxItem.Read", "MailboxItem.ImportExport")
    Connect-MgGraph -Scopes $scopes


#>
function Invoke-ImportEmlFile {
    param (        
        [string]$UserId = "",
        [string]$EmlPath = "",
        [string]$TargetFolder = ""
    )

    # Convert EML file to base64
    $emlBytes = [System.IO.File]::ReadAllBytes($EmlPath)
    $base64Eml = [Convert]::ToBase64String($emlBytes)
    $contentBytes = [System.Text.Encoding]::UTF8.GetBytes($base64Eml)

    # Import Message from Mime Will create a draft.
    $createdMessage = Invoke-MgGraphRequest `
        -Method POST `
        -Uri "https://graph.microsoft.com/v1.0/users/$UserId/messages" `
        -Headers @{ "Content-Type" = "text/plain" } `
        -Body $contentBytes

    if ($createdMessage -and $createdMessage.id) {
        Write-Host "Draft message created. ID: $($createdMessage.id)"
    } else {
        Write-Host "Message creation failed."
        return
    }

    # Export, mark as sent, and remove draft message
    $MailboxSettings = Invoke-GetMailboxSettings -Upn $UserId
    $ExportItem = Invoke-ExportItems -ItemId $createdMessage.Id -MailboxId $MailboxSettings.primaryMailboxId
    $ModifiedDataStream = Invoke-MarkMessageAsSent -Data $ExportItem
    Remove-MgUserMessage -UserId $UserId -MessageId $createdMessage.Id -Confirm:$false

    # Import the item to the correct folder
    $ImportSession = Invoke-CreateImportSession -MailboxId $MailboxSettings.primaryMailboxId
    $FolderId = Get-MgUserMailFolder -MailFolderId $TargetFolder -UserId $UserId
    Invoke-UploadItem -FolderId $FolderId.Id -ImportURL $ImportSession.importUrl -Data $ModifiedDataStream

    
}


function Invoke-CreateImportSession {
    param ([Parameter(Mandatory = $true)][PsObject]$MailboxId)
    Process {
        $RequestURL = "https://graph.microsoft.com/beta/admin/exchange/mailboxes/$MailboxId/createImportSession"
        return Invoke-GraphRequest -Uri $RequestURL -Method Post
    }
}

function Invoke-GetMailboxSettings {
    param ([Parameter(Mandatory = $true)][string]$Upn)
    Process {
        $RequestURL = "https://graph.microsoft.com/beta/users/$Upn/settings/exchange"
        return Invoke-MgGraphRequest -Method Get -Uri $RequestURL
    }
}

function Invoke-MarkMessageAsSent {
    param ([byte[]]$Data)

    function Find-PidTagMessageFlags {
        param([byte[]]$Bytes)
        $pattern = 0x03, 0x00, 0x07, 0x0E
        for ($i = 0; $i -le $Bytes.Length - $pattern.Length; $i++) {
            if ($pattern -eq $Bytes[$i..($i + $pattern.Length - 1)]) { return $i }
        }
        return -1
    }

    $UnsentFlag = 0x0008
    $SubmitFlag = 0x0004
    $tagIndex = Find-PidTagMessageFlags -Bytes $Data

    if ($tagIndex -ge 0) {
        $valueOffset = $tagIndex + 4
        $oldFlags = [BitConverter]::ToInt32($Data, $valueOffset)
        $newFlags = ($oldFlags -band (-bnot $UnsentFlag)) -bor $SubmitFlag
        [Array]::Copy([BitConverter]::GetBytes($newFlags), 0, $Data, $valueOffset, 4)
    }

    return $Data
}

function Invoke-ExportItems {
    param (
        [Parameter(Mandatory = $true)][PsObject]$ItemId,
        [Parameter(Mandatory = $true)][PsObject]$MailboxId
    )
    Process {
        $ExportUrl = "https://graph.microsoft.com/beta/admin/exchange/mailboxes/$MailboxId/exportItems"
        $ExportBatch = @($ItemId)
        $ExportItemsResult = Invoke-GraphRequest -Uri $ExportUrl -Method Post -Body (ConvertTo-Json @{ itemIds = $ExportBatch } -Depth 10)
        return [Convert]::FromBase64String($ExportItemsResult.value[0].data)
    }
}

function Invoke-UploadItem {
    param (
        [Parameter(Mandatory = $true)][string]$FolderId,
        [Parameter(Mandatory = $true)][string]$ImportURL,
        [Parameter(Mandatory = $true)][byte[]]$Data
    )
    Process {
        $Request = @{
            FolderId = $FolderId
            Mode     = "create"
            Data     = [Convert]::ToBase64String($Data)
        }
        $result = Invoke-WebRequest -Method Post -Uri $ImportURL -Body (ConvertTo-Json $Request -Depth 10) -ContentType "application/json"
        if($result.StatusCode -ne 200){
            Write-Error "Error uploading Item"              
        }else{
            Write-Host "Message imported successfully."
        } 
    }
}