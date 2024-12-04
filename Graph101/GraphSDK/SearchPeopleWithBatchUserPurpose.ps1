function Invoke-SearchPeopleWithBatch{
            [CmdletBinding()]
    param (
	[Parameter(Position = 1, Mandatory = $true)]
        [String]
        $UserId,
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $SearchPhrase
    )
    Process{
        $BatchRequestContent = @{}
        $BatchRequestContent.add("requests",@())
        $UserOutBatch = @()
        Get-MgUserPerson -User $UserId -Top 100 -Search $SearchPhrase -Headers @{"X-PeopleQuery-QuerySources"="Directory"} | foreach-object{
            $adGuid = $_.id
            $UserOutBatch += $_
            $BatchEntry = @{}
            $BatchEntry.Add("id",$_.id)
            $BatchEntry.Add("method","GET")
            $BatchEntry.Add("url","/users/$adGuid/MailboxSettings/UserPurpose")            
            $BatchHeaders = @{
                'Content-Type' =  "application/json"
            } 
            $BatchEntry.Add("headers",$BatchHeaders)
            $BatchRequestContent["requests"] += $BatchEntry
            if($UserOutBatch.Count -eq 10){
                $BatchResponse = Invoke-BatchGet -BatchRequestContent $BatchRequestContent               
                foreach($User in $UserOutBatch){
                    Write-Verbose $User.DisplayName
                    Write-Verbose $User.Id
                    if($BatchResponse[$User.Id].status -eq 200){
                        $User | Add-Member -NotePropertyName UserPurpose -NotePropertyValue $BatchResponse[$User.Id].body["value"]
                        Write-Verbose ($BatchResponse[$User.Id].body["value"])
                        
                    }else{
                        Write-Verbose ("Error")
                        Write-Verbose ($BatchResponse[$User.Id].status)
                        $User | Add-Member -NotePropertyName UserPurpose -NotePropertyValue $BatchResponse[$User.Id].status
                    }
                    Write-Output $User                 
                }
                $UserOutBatch = @()
                $BatchRequestContent = @{}
                $BatchRequestContent.add("requests",@())
            }
        }
        if($UserOutBatch.Count -gt 0){
                write-Verbose ("Execute Batch" + $batchCount)
                $BatchResponse = Invoke-BatchGet -BatchRequestContent $BatchRequestContent                
                foreach($User in $UserOutBatch){
                    Write-Verbose $User.DisplayName
                    Write-Verbose $User.Id
                    if($BatchResponse[$User.Id].status -eq 200){
                        $User | Add-Member -NotePropertyName UserPurpose -NotePropertyValue $BatchResponse[$User.Id].body["value"]
                        Write-Verbose ($BatchResponse[$User.Id].body["value"])
                        
                    }else{
                        Write-Verbose ("Error")
                        Write-Verbose $BatchResponse[$User.Id].status
                        $User | Add-Member -NotePropertyName UserPurpose -NotePropertyValue $BatchResponse[$User.Id].status
                    }
                    Write-Output $User
                }
        }       
    }
}

function Invoke-BatchGet{
    param (
        [Parameter(Position = 1, Mandatory = $true)]
        [PSObject]
        $BatchRequestContent
    )
    Process{
        $IndexedResponse = @{}
        $RequestURL = "https://graph.microsoft.com/v1.0/`$batch"
        $BatchResponse = Invoke-MgGraphRequest -Method POST -Uri $RequestURL -Body (ConvertTo-json  $BatchRequestContent -depth 10 -Compress) 
        if($BatchResponse.responses){
            foreach($Response in $BatchResponse.responses){     
                $IndexedResponse.Add($Response.Id,$Response)                   
            }
        }
        return, $IndexedResponse 
    }
}

