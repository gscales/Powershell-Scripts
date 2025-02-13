function Invoke-SendMessageWithLargeAttachment {
    [CmdletBinding()]
    param (
        [Parameter(Position = 1, Mandatory = $true)]
        [String] $ToRecipient,
        [Parameter(Position = 2, Mandatory = $true)]
        [String] $UserId,
        [Parameter(Position = 3, Mandatory = $true)]
        [String] $Subject,
        [Parameter(Position = 4, Mandatory = $true)]
        [String] $Body,
        [Parameter(Position = 5, Mandatory = $true)]
        [String] $AttachmentPath
    )
    Process {
        #Create Draft Message
        Import-Module Microsoft.Graph.Mail
        Import-Module Microsoft.Graph.Users.Actions
        $uploadChunkSize = 327680
        $params = @{
            subject      = $subject
            body         = @{
                contentType = "HTML"
                content     = $body 
            }
            toRecipients = @(
                @{
                    emailAddress = @{
                        address = $ToRecipient
                    }
                }
            )
        }

        # A UPN can also be used as -UserId.
        $Draftmessage = New-MgUserMessage -UserId $userId -BodyParameter $params
        $FileStream = New-Object System.IO.StreamReader($attachmentpath)  
        $FileSize = $fileStream.BaseStream.Length 

        $params = @{
            AttachmentItem = @{
                attachmentType = "File"
                name           = [System.Io.Path]::GetFileName($attachmentpath)
                size           = $FileSize
            }
        }

        # A UPN can also be used as -UserId.
        $UploadSession = New-MgUserMessageAttachmentUploadSession -UserId $userId -MessageId $Draftmessage.Id -BodyParameter $params
        $FileOffsetStart = 0              
        $FileBuffer = [byte[]]::new($uploadChunkSize)
        do {            
            $FileChunkByteCount = $fileStream.BaseStream.Read($FileBuffer, 0, $FileBuffer.Length) 
            Write-Verbose ($fileStream.BaseStream.Position)
            $FileOffsetEnd = $fileStream.BaseStream.Position - 1
            if ($FileChunkByteCount -gt 0) {
                $UploadRangeHeader = "bytes " + $FileOffsetStart + "-" + $FileOffsetEnd + "/" + $FileSize
                Write-Verbose $UploadRangeHeader                
                $FileOffsetStart = $fileStream.BaseStream.Position
                $binaryContent = New-Object System.Net.Http.ByteArrayContent -ArgumentList @($FileBuffer, 0, $FileChunkByteCount)
                $FileBuffer = [byte[]]::new($uploadChunkSize)
                $headers = @{
                    'AnchorMailbox' = $userId
                    'Content-Range' = $UploadRangeHeader
                }
                $Result = (Invoke-RestMethod -Method Put -Uri $UploadSession.UploadUrl -UserAgent "UploadAgent" -Headers $headers -Body $binaryContent.ReadAsByteArrayAsync().Result -ContentType "application/octet-stream") 
                Write-Verbose $Result 
            }          

        }while ($FileChunkByteCount -ne 0)    

        Send-MgUserMessage -UserId $userId -MessageId $Draftmessage.Id
    }
}