function Get-InboxRules {
    [CmdletBinding()]
    param (       
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $MailboxName,
        [Parameter(Position = 2, Mandatory = $false)]
        [String]
        $ClientId
    )
    Begin {
        if([String]::IsNullOrEmpty($ClientId)){
            $ClientId = "5471030d-f311-4c5d-91ef-74ca885463a7"
        }		
        Import-Module .\Microsoft.IdentityModel.Clients.ActiveDirectory.dll -Force
        $PromptBehavior = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters -ArgumentList Auto       
        $Context = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext("https://login.microsoftonline.com/common")
        $token = ($Context.AcquireTokenAsync("https://graph.microsoft.com", $ClientId , "urn:ietf:wg:oauth:2.0:oob", $PromptBehavior)).Result
        $Header = @{
            'Content-Type'  = 'application\json'
            'Authorization' = $token.CreateAuthorizationHeader()
        }        
        $Result =  Invoke-RestMethod -Headers $Header -Uri ("https://graph.microsoft.com/v1.0/users('" + $MailboxName + "')/mailFolders('inbox')/messageRules") -Method Get 
        if ($Result.value -ne $null) {
            foreach ($Message in $Result.value ) {
                write-output $Message
            }
        }
       	
		
    }
}

