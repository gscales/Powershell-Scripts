Connect-MgGraph -Scopes "Mail.Read" -TenantId "Domain.com"
$MailboxName = "User@Domain.com"
$Subject = "Daily Export"
$ProcessedFolderPath = "\Inbox\Processed"
$downloadDirectory = "c:\temp"
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
$ProcessedFolder = Get-MailBoxFolderFromPath -MailboxName $MailboxName -FolderPath $ProcessedFolderPath
$LastMessage = Get-MgUserMailFolderMessage -UserId $MailboxName -MailFolderId inbox -Filter "ReceivedDateTime ge 2022-01-01T00:00:00Z and Subject eq '$Subject' and hasAttachments eq true and isRead eq false" -Top 1 -orderby 'ReceivedDateTime  DESC'
if($LastMessage){
    $Attachment = Get-MgUserMailFolderMessageAttachment -UserId $MailboxName -MailFolderId inbox -MessageId $LastMessage.Id
    $Base64B = ($Attachment).AdditionalProperties.contentBytes
    $Bytes = [Convert]::FromBase64String($Base64B)
    $fiFile = new-object System.IO.FileStream(($downloadDirectory + "\" + $Attachment.Name.ToString()), [System.IO.FileMode]::Create)
    $fiFile.Write($Bytes, 0, $Bytes.Length)
    $fiFile.Close()
    write-host "Downloaded Attachment : " + (($downloadDirectory + "\" + $Attachment.Name.ToString()))
    Update-MgUserMailFolderMessage -MailFolderId inbox -MessageId $LastMessage.Id -IsRead -UserId $MailboxName
    Move-MgUserMailFolderMessage -MailFolderId inbox -MessageId $LastMessage.Id -UserId $MailboxName -DestinationId $ProcessedFolder.Id
}else{
    Write-Host "No Messages Found"
}




