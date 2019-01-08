
function Invoke-TranslateExchangeIds {
    [CmdletBinding()]
    param (       
        [Parameter(Position = 2, Mandatory = $false)]
        [String]
        $SourceId, 
        
        [Parameter(Position = 3, Mandatory = $false)]
        [String]
        $SourceHexId,
        
        [Parameter(Position = 4, Mandatory = $false)]
        [String]
        $SourceEMSId,

        [Parameter(Position = 5, Mandatory = $false)]
        [String]
        $SourceFormat,  

        [Parameter(Position = 6, Mandatory = $false)]
        [String]
        $TargetFormat  
    )
    Begin {		
        Import-Module .\Microsoft.IdentityModel.Clients.ActiveDirectory.dll -Force
        $PromptBehavior = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters -ArgumentList Auto       
        $Context = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext("https://login.microsoftonline.com/common")
        $token = ($Context.AcquireTokenAsync("https://graph.microsoft.com", "d3590ed6-52b3-4102-aeff-aad2292ab01c", "urn:ietf:wg:oauth:2.0:oob", $PromptBehavior)).Result
        $ConvertRequest = @{}
        $ConvertRequest.Add("inputIds", @())
        if ($SourceHexId) {
            $byteArray = @($SourceHexId -split '([a-f0-9]{2})' | foreach-object { if ($_) {[System.Convert]::ToByte($_, 16)}})
            $urlSafeString = [Convert]::ToBase64String($byteArray).replace("/", "_").replace("+", "-")
            if ($urlSafeString.contains("==")) {$urlSafeString = $urlSafeString.replace("==", "2")}
            $ConvertRequest["inputIds"] += $urlSafeString

        }
        else {
            if ($SourceEMSId) {
                $HexEntryId = [System.BitConverter]::ToString([Convert]::FromBase64String($SourceEMSId)).Replace("-", "").Substring(2)  
                $HexEntryId = $HexEntryId.SubString(0, ($HexEntryId.Length - 2))
                $byteArray = @($HexEntryId -split '([a-f0-9]{2})' | foreach-object { if ($_) {[System.Convert]::ToByte($_, 16)}})
                $urlSafeString = [Convert]::ToBase64String($byteArray).replace("/", "_").replace("+", "-")
                if ($urlSafeString.contains("==")) {$urlSafeString = $urlSafeString.replace("==", "2")}
                $ConvertRequest["inputIds"] += $urlSafeString
            }
            else {
                $ConvertRequest["inputIds"] += $SourceId
            }

        }        
        $ConvertRequest.targetIdType = $TargetFormat
        $ConvertRequest.sourceIdType = $SourceFormat
        $Header = @{
            'Content-Type'  = 'application\json'
            'Authorization' = $token.CreateAuthorizationHeader()
        }        
        $JsonResult = Invoke-RestMethod -Headers $Header -Uri "https://graph.microsoft.com/beta/me/translateExchangeIds" -Method Post -ContentType "application/json" -Body (ConvertTo-Json $ConvertRequest -Depth 9)
        if($TargetFormat.ToLower() -eq "entryid"){
            $urldecodedstring = $JsonResult.value.targetId.replace("_", "/").replace("-", "+")
            $lastVal = $urldecodedstring.SubString($urldecodedstring.Length-1,1);
            if($lastVal -eq "2"){
                $urldecodedstring = $urldecodedstring.SubString(0,$urldecodedstring.Length-1) + "=="
            }
            return ([System.BitConverter]::ToString([Convert]::FromBase64String($urldecodedstring))).replace("-","")
        }else{
            return  $JsonResult.value.targetId
        }
       	
		
    }
}
