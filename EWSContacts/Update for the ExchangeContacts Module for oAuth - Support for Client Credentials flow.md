# Update for the ExchangeContacts Module for oAuth - Support for Client Credentials flow #

Back in 2019 i updated my [ExchangeContacts](https://www.powershellgallery.com/packages/ExchangeContacts/1.7.0.0) module on the Powershell Gallery to support Modern Authentication using the ADAL library [https://gsexdev.blogspot.com/2019/12/update-to-exchangecontacts-module-to.html](https://gsexdev.blogspot.com/2019/12/update-to-exchangecontacts-module-to.html). The one thing i didn't include at the time was support for the [client credentials Authentication flow](https://docs.microsoft.com/en-us/azure/active-directory/develop/msal-authentication-flows#client-credentials) so a question from someone this week about switching from using basic Authenticaiton with that module prompted me to upgrade the ADAL code to MSAL and also to add support for the client credentials flow.

All the important code changes have been made in  all the changes where made in the [https://github.com/gscales/Powershell-Scripts/blob/master/EWSContacts/Module/functions/service/Connect-EXCExchange.ps1](https://github.com/gscales/Powershell-Scripts/blob/master/EWSContacts/Module/functions/service/Connect-EXCExchange.ps1)

The code supports 3 Authentication flows first Interactive Authentication Code (note you can tweek the prompt if you want it to be semi silent)

	$scope = "https://outlook.office.com/EWS.AccessAsUser.All";
	$Scopes = New-Object System.Collections.Generic.List[string]
	$Scopes.Add($Scope)				
	$pcaConfig = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($ClientId).WithRedirectUri$redirectUri)
	$token = $pcaConfig.Build().AcquireTokenInteractive($Scopes).WithPrompt([Microsoft.Identity.Client.Prompt]::SelectAccount).WithLoginHint($MailboxName).ExecuteAsync().Result
	$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($token.AccessToken)

to use this

    Get-EXCContacts -MailboxName gscales@datarumble.com -ModernAuth

ROPC using implicit credentials

	$scope = "https://outlook.office.com/EWS.AccessAsUser.All";
	$Scopes = New-Object System.Collections.Generic.List[string]
	$Scopes.Add($Scope)				
	$pcaConfig = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($ClientId).WithAuthority([Microsoft.Identity.Client.AadAuthorityAudience]::AzureAdMultipleOrgs)				
	$token = $pcaConfig.Build().AcquireTokenByUsernamePassword($Scopes,$Credentials.UserName,$Credentials.Password).ExecuteAsync().Result;
	$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($token.AccessToken)

to use this
	
    $creds = Get-Credentials
    Get-EXCContacts -MailboxName gscales@datarumble.com -ModernAuth -Credentials $creds

And the Client Credentials flow using a passed in certificate file and password

	$exVal = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable
	$certificateObject = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $CertificateFilePath, $CertificatePassword , $exVal
	$domain = $MailboxName.Split("@")[1]
	$Scope = "https://outlook.office365.com/.default"
	$TenantId = (Invoke-WebRequest https://login.windows.net/$domain/v2.0/.well-known/openid-configuration | ConvertFrom-Json).token_endpoint.Split('/')[3]
	$app =  [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($ClientId).WithCertificate($certificateObject).WithTenantId($TenantId).Build()
	$Scopes = New-Object System.Collections.Generic.List[string]
	$Scopes.Add($Scope)
	$token = $app.AcquireTokenForClient($Scopes).ExecuteAsync().Result
	$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($token.AccessToken)
	$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)	

For this you can use something like

    $certPassword = ConvertTo-SecureString -String "1234xxxx" -Force -AsPlainText
    Get-EXCContacts -MailboxName gscales@datarumble.com -ModernAuth -CertificateFilePath C:\temp\mp.pfx -CertificatePassword $certPassword -ClientId 2fe2dc9c-b746-4112-8a11-5054dce06af4

When using the client credentials flow it will use EWS Impersonation so the TargetMailbox passed in as the -MailboxName will be impersonated.