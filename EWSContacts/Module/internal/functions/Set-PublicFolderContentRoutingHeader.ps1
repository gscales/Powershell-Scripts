function Set-PublicFolderContentRoutingHeader
{
<#
	.SYNOPSIS
		A brief description of the Set-PublicFolderContentRoutingHeader function.
	
	.DESCRIPTION
		A detailed description of the Set-PublicFolderContentRoutingHeader function.
	
	.PARAMETER service
		A description of the service parameter.
	
	.PARAMETER Credentials
		A description of the Credentials parameter.
	
	.PARAMETER MailboxName
		A description of the MailboxName parameter.
	
	.PARAMETER pfAddress
		A description of the pfAddress parameter.
	
	.EXAMPLE
		PS C:\> Set-PublicFolderContentRoutingHeader -service $service -Credentials $Credentials -MailboxName 'value3' -pfAddress 'value4'
#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[Microsoft.Exchange.WebServices.Data.ExchangeService]
		$service,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[PSCredential]
		$Credentials,
		
		[Parameter(Position = 2, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[string]
		$pfAddress
	)
	process
	{
		$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
		$AutoDiscoverService = New-Object  Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverService($ExchangeVersion);
		$AutoDiscoverService.Credentials = $service.Credentials
		$AutoDiscoverService.EnableScpLookup = $false;
		$AutoDiscoverService.RedirectionUrlValidationCallback = { $true };
		$AutoDiscoverService.PreAuthenticate = $true;
		$AutoDiscoverService.KeepAlive = $false;
		$gsp = $AutoDiscoverService.GetUserSettings($MailboxName, [Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::AutoDiscoverSMTPAddress);
		#Write-Host $AutoDiscoverService.url
		$auDisXML = "<Autodiscover xmlns=`"http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006`"><Request>`r`n" +
		"<EMailAddress>" + $pfAddress + "</EMailAddress>`r`n" +
		"<AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema>`r`n" +
		"</Request>`r`n" +
		"</Autodiscover>`r`n";
		$AutoDiscoverRequest = [System.Net.HttpWebRequest]::Create($AutoDiscoverService.url.ToString().replace(".svc", ".xml"));
		$bytes = [System.Text.Encoding]::UTF8.GetBytes($auDisXML);
		$AutoDiscoverRequest.ContentLength = $bytes.Length;
		$AutoDiscoverRequest.ContentType = "text/xml";
		$AutoDiscoverRequest.UserAgent = "Microsoft Office/16.0 (Windows NT 6.3; Microsoft Outlook 16.0.6001; Pro)";
		$AutoDiscoverRequest.Headers.Add("Translate", "F");
		$AutoDiscoverRequest.Method = "POST";
		$AutoDiscoverRequest.Credentials = $creds;
		$RequestStream = $AutoDiscoverRequest.GetRequestStream();
		$RequestStream.Write($bytes, 0, $bytes.Length);
		$RequestStream.Close();
		$AutoDiscoverRequest.AllowAutoRedirect = $truee;
		$Response = $AutoDiscoverRequest.GetResponse().GetResponseStream()
		$sr = New-Object System.IO.StreamReader($Response)
		[XML]$xmlReponse = $sr.ReadToEnd()
		if ($xmlReponse.Autodiscover.Response.User.AutoDiscoverSMTPAddress -ne $null)
		{
			Write-Verbose "Public Folder Content Routing Information Header : $($xmlReponse.Autodiscover.Response.User.AutoDiscoverSMTPAddress)"
			$service.HttpHeaders["X-AnchorMailbox"] = $xmlReponse.Autodiscover.Response.User.AutoDiscoverSMTPAddress
			$service.HttpHeaders["X-PublicFolderMailbox"] = $xmlReponse.Autodiscover.Response.User.AutoDiscoverSMTPAddress
		}
		
	}
	
}
