function Get-AutoDiscoverPhotoURL
{
<#
	.SYNOPSIS
		A brief description of the Get-AutoDiscoverPhotoURL function.
	
	.DESCRIPTION
		A detailed description of the Get-AutoDiscoverPhotoURL function.
	
	.PARAMETER EmailAddress
		A description of the EmailAddress parameter.
	
	.PARAMETER Credentials
		A description of the Credentials parameter.
	
	.EXAMPLE
		PS C:\> Get-AutoDiscoverPhotoURL -EmailAddress 'value1' -Credentials $Credentials
#>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]
		$EmailAddress,
		
		[Parameter(Mandatory = $true)]
		[Microsoft.Exchange.WebServices.Data.ExchangeService]
		$service
		
	)
	process
	{
		$version = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013
		$adService = New-Object Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverService($version);
		$adService.Credentials = $service.Credentials
		$adService.EnableScpLookup = $false;
		$adService.RedirectionUrlValidationCallback = { $true }
		$adService.PreAuthenticate = $true;
		$UserSettings = new-object Microsoft.Exchange.WebServices.Autodiscover.UserSettingName[] 1
		$UserSettings[0] = [Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::ExternalPhotosUrl
		$adResponse = $adService.GetUserSettings($EmailAddress, $UserSettings)
		$PhotoURI = $adResponse.Settings[[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::ExternalPhotosUrl]
		return $PhotoURI.ToString()
	}
}
