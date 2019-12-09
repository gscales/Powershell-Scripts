function Set-PublicFolderRoutingHeader
{
<#
	.SYNOPSIS
		A brief description of the Set-PublicFolderRoutingHeader function.
	
	.DESCRIPTION
		A detailed description of the Set-PublicFolderRoutingHeader function.
	
	.PARAMETER Service
		A description of the Service parameter.
	
	.PARAMETER Credentials
		A description of the Credentials parameter.
	
	.PARAMETER MailboxName
		A description of the MailboxName parameter.
	
	.PARAMETER Header
		A description of the Header parameter.
	
	.EXAMPLE
		PS C:\> Set-PublicFolderRoutingHeader -Service $Service -Credentials $Credentials -MailboxName 'value3' -Header 'value4'
#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[Microsoft.Exchange.WebServices.Data.ExchangeService]
		$Service,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[PSCredential]
		$Credentials,
		
		[Parameter(Position = 2, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[string]
		$Header
	)
	process
	{
		$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
		$AutoDiscoverService = New-Object  Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverService($ExchangeVersion);
		$AutoDiscoverService.Credentials = $Service.Credentials
		$AutoDiscoverService.EnableScpLookup = $false;
		$AutoDiscoverService.RedirectionUrlValidationCallback = { $true };
		$AutoDiscoverService.PreAuthenticate = $true;
		$AutoDiscoverService.KeepAlive = $false;
		if ($Header -eq "X-AnchorMailbox")
		{
			$gsp = $AutoDiscoverService.GetUserSettings($MailboxName, [Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::PublicFolderInformation);
			$PublicFolderInformation = $null
			if ($gsp.Settings.TryGetValue([Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::PublicFolderInformation, [ref]$PublicFolderInformation))
			{
				Write-Verbose "Public Folder Routing Information Header : $PublicFolderInformation"
				if (!$Service.HttpHeaders.ContainsKey($Header))
				{
					$Service.HttpHeaders.Add($Header, $PublicFolderInformation)
				}
				else
				{
					$Service.HttpHeaders[$Header] = $PublicFolderInformation
				}
			}
		}
	}
}
