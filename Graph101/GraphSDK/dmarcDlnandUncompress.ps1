function Invoke-DownloadAndProcessDmarc{
        [CmdletBinding()]
    param (
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $MailboxName,
        [Parameter(Position = 2, Mandatory = $true)]
        [String]
        $HoursToLookBack,
        [Parameter(Position = 2, Mandatory = $false)]
        [String]
        $FolderPath
    )
    Process{
        $FolderId = "Inbox"
        if($FolderPath){
            $FolderId = (Get-MailBoxFolderFromPath -MailboxName $MailboxName -FolderPath $FolderPath).Id
        }
        $received = (Get-Date).AddHours(-$HoursToLookBack).ToUniversalTime().ToString("yyyy-MM-ddThh:mm:ss")
        $ipv4regexDef = '^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$'
        $Messages = Get-MgUserMailFolderMessage -Userid $MailboxName -All -MailFolderId $FolderId -Search "(received>=$received) AND (subject:Report-Id) AND ((attachmentnames:.gz) OR ( attachmentnames:.zip))" -ExpandProperty "Attachments" -Select Subject,ReceivedDateTime,Attachments,Sender
        foreach ($Message in $Messages) {
            Write-Verbose ("Processing Message " + $Message.Subject) 
            if ($Message.Subject.Contains("Report-ID")) {
                foreach ($attachment in $Message.Attachments) {
                    $dmarcResult = $null
                    if ($attachment.contentType -eq "application/zip") {
                        $atContent = [System.Convert]::FromBase64String($attachment.AdditionalProperties.contentBytes)
                        $compressedStream = new-object System.IO.MemoryStream(, $atContent) 
                        $zipArchive = new-object System.IO.Compression.ZipArchive($compressedStream, [System.IO.Compression.CompressionMode]::Decompress)
                        $reader = [System.IO.StreamReader]::new($zipArchive.entries[0].Open())
                        [XML]$dracResult = $reader.ReadToEnd()
                    }
                    if ($attachment.contentType -eq "application/gzip") {
                        $atContent = [System.Convert]::FromBase64String($attachment.AdditionalProperties.contentBytes)
                        $compressedStream = new-object System.IO.MemoryStream(, $atContent) 
                        $gzipStream = new-object System.IO.Compression.GZipStream($compressedStream, [System.IO.Compression.CompressionMode]::Decompress)
                        $reader = [System.IO.StreamReader]::new($gzipStream)
                        [XML]$dracResult = $reader.ReadToEnd()                
                    }  
                    if($dracResult){
                        foreach($row in $dracResult.feedback.record){
                            $rptObj = "" | Select ReportName,Sender,org_name,report_id,StartData,EndDate,source_ip,GeoCountry,GeoCity,GeoRecord,count,disposition,dkim,spf,auth_result_dkim,auth_result_spf,record
                            $rptObj.ReportName = $Message.Subject
                            $rptObj.Sender = $Message.Sender.EmailAddress.Address
                            $rptObj.org_name = $dracResult.feedback.report_metadata.org_name
                            $rptObj.report_id = $dracResult.feedback.report_metadata.report_id
                            $rptObj.StartData = $dracResult.feedback.report_metadata.date_range.begin
                            $rptObj.EndDate = $dracResult.feedback.report_metadata.date_range.end
                            $rptObj.source_ip = $row.row.source_ip
                            if($rptObj.source_ip -match $ipv4regexDef){
                                try{
                                    $geoipLookup = Invoke-RestMethod -Method Get -Uri ("http://ip-api.com/json/" + $rptObj.source_ip)
                                    $rptObj.GeoCountry = $geoipLookup.country
                                    $rptObj.GeoCity = $geoipLookup.city
                                    $rptObj.GeoRecord = $geoipLookup 
                                }catch{
                                    Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__ 
                                    Write-Host "StatusDescription:" $_.Exception.Response.StatusDescription
                                }
                            }
                            $rptObj.count = $row.row.count
                            if($row.auth_results.dkim){
                                $rptObj.auth_result_dkim = $row.auth_results.dkim.domain + " "  + $row.auth_results.dkim.result
                            }
                            if($row.auth_results.spf){
                                $rptObj.auth_result_spf = $row.auth_results.spf.domain + " "  + $row.auth_results.spf.result
                            }
                            if($row.row.policy_evaluated){
                                $rptObj.disposition = $row.row.policy_evaluated.disposition
                                $rptObj.dkim = $row.row.policy_evaluated.dkim
                                $rptObj.spf = $row.row.policy_evaluated.spf
                            }
                            $rptObj.record = $row
                            Write-Output $rptObj                                            
                        }
                    }     
                }
            }
        }
    }
}

function Get-MailBoxFolderFromPath {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string]
        $FolderPath,
		
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $MailboxName,

        [Parameter(Position = 2, Mandatory = $false)]
        [String]
        $WellKnownSearchRoot = "MsgFolderRoot"
    )

    process {
        if($FolderPath -eq '\'){
            return Get-MgUserMailFolder -UserId $MailboxName -MailFolderId msgFolderRoot 
        }
        $fldArray = $FolderPath.Split("\")
        #Loop through the Split Array and do a Search for each level of folder 
        $folderId = $WellKnownSearchRoot
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