$UserName = "user"
$password = "pass"
Import-Module .\Microsoft.IdentityModel.Clients.ActiveDirectory.dll -Force
$upCred = new-Object  Microsoft.IdentityModel.Clients.ActiveDirectory.UserPasswordCredential($UserName, $Password);
$EndpointUri = 'https://login.windows.net/common'
$Context = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext($EndpointUri)
$AADcredential = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserPasswordCredential" -ArgumentList $UserName, $Password
$authenticationResult = [Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContextIntegratedAuthExtensions]::AcquireTokenAsync($Context ,"https://graph.microsoft.com","d3590ed6-52b3-4102-aeff-aad2292ab01c",$AADcredential).result
$Header = @{
   'Content-Type'  = 'application\json'
   'Authorization' = $authenticationResult.CreateAuthorizationHeader()
}        
$Result =  Invoke-RestMethod -Headers $Header -Uri ("https://graph.microsoft.com/v1.0/me/drive/root") -Method Get
write-host $Result.webUrl

