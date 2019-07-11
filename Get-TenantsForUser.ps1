function Get-TenantsForUser {
    [CmdletBinding()]
    param (       
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $UserName,
        [Parameter(Position = 2, Mandatory = $false)]
        [String]
        $ClientId
    )
    Begin {        
        if ([String]::IsNullOrEmpty($ClientId)) {
            $ClientId = "1950a258-227b-4e31-a9cf-717495945fc2"
        }		
        Import-Module .\Microsoft.IdentityModel.Clients.ActiveDirectory.dll -Force
       # $UserId = new-object Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifier -ArgumentList $UserName,UniqueId
        $PromptBehavior = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters -ArgumentList SelectAccount       
        $Context = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext("https://login.microsoftonline.com/common")
        $token = ($Context.AcquireTokenAsync("https://management.core.windows.net/", $ClientId , "urn:ietf:wg:oauth:2.0:oob", $PromptBehavior)).Result
        $Header = @{
            'Content-Type'  = 'application\json'
            'Authorization' = $token.CreateAuthorizationHeader()     
        }

        $UserResult = (Invoke-RestMethod -Headers $Header -Uri ("https://management.azure.com/tenants?api-version=2019-03-01&`$includeAllTenantCategories=true") -Method Get -ContentType "Application/json").value
        foreach($domain in $UserResult){            
            Write-Verbose $domain
            $TestPromptBehavior = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters -ArgumentList Auto  
            $TestContext = new-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext(("https://login.microsoftonline.com/" + $domain.tenantId))
            $TestToken = ($TestContext.AcquireTokenAsync("https://graph.microsoft.com/", "d3590ed6-52b3-4102-aeff-aad2292ab01c", "urn:ietf:wg:oauth:2.0:oob", $TestPromptBehavior)).Result
            $headers = @{
                'Content-Type'  = 'application\json'
                'Authorization' = $TestToken.CreateAuthorizationHeader()     
            }
            $Response = (Invoke-WebRequest -Uri "https://graph.microsoft.com/v1.0/me/memberOf" -Headers $headers) 
            $JsonResponse = ConvertFrom-Json $Response.Content
            $Groups = @()
            Add-Member -InputObject $domain -NotePropertyName Groups -NotePropertyValue  $Groups
            foreach ($Group in $JsonResponse.value) {                
                Write-Verbose $Group      
                $Conversations = @()       
                Add-Member -InputObject $Group  -NotePropertyName Conversations -NotePropertyValue $Conversations  
                if ($Group.groupTypes -eq "Unified") {
                    $cnvResponse = (Invoke-WebRequest -Uri ("https://graph.microsoft.com/v1.0/groups/" + $Group.id + "/conversations?Top=2") -Headers $headers) 
                    $cnvJsonResponse = ConvertFrom-Json $cnvResponse.Content
                    foreach ($cnv in $cnvJsonResponse.value) {
                        Write-Verbose $cnv
                        $Group.Conversations += $cnv
                    }	
                }
                $domain.Groups += $Group
                
            }
        }
        Write-Output $UserResult
       	
		
    }
}



