function Export-EXCGALContact
{
<#
	.SYNOPSIS
		Exports a Contact from the Global Address List on an Exchange Server using the  Exchange Web Services API to a VCF File
	
	.DESCRIPTION
		Exports a Contact from the Global Address List on an Exchange Server using the  Exchange Web Services API to a VCF File
		
		Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951
	
	.PARAMETER MailboxName
		A description of the MailboxName parameter.
	
	.PARAMETER EmailAddress
		A description of the EmailAddress parameter.
	
	.PARAMETER Credentials
		A description of the Credentials parameter.
	
	.PARAMETER IncludePhoto
		A description of the IncludePhoto parameter.
	
	.PARAMETER FileName
		A description of the FileName parameter.
	
	.PARAMETER Partial
		A description of the Partial parameter.
	
	.EXAMPLE
		Example 1 To export a GAL Entry to a vcf file
		Export-EXCGALContact -MailboxName user@domain.com -EmailAddress email@domain.com -FileName c:\export\export.vcf
		
	.EXAMPLE
		Example 2 To export a GAL Entry to vcf including the users photo
		Export-EXCGALContact -MailboxName user@domain.com -EmailAddress email@domain.com -FileName c:\export\export.vcf -IncludePhoto
#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$EmailAddress,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[System.Management.Automation.PSCredential]
		$Credentials,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[switch]
		$IncludePhoto,
		
		[Parameter(Position = 4, Mandatory = $true)]
		[string]
		$FileName,
		
		[Parameter(Position = 5, Mandatory = $false)]
		[switch]
		$Partial,
				
		[Parameter(Position = 6, Mandatory = $False)]
		[switch]
		$ModernAuth,
		
		[Parameter(Position = 7, Mandatory = $False)]
		[String]
		$ClientId
		
	)
	process
	{
		$service = Connect-EXCExchange -MailboxName $MailboxName -Credential $Credentials -ModernAuth:$ModernAuth.IsPresent -ClientId $ClientId
		$Error.Clear()
		$cnpsPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
		$ncCol = $service.ResolveName($EmailAddress, $ParentFolderIds, [Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryOnly, $true, $cnpsPropset)
		if ($Error.Count -eq 0)
		{
			foreach ($Result in $ncCol)
			{
				if (($Result.Mailbox.Address.ToLower() -eq $EmailAddress.ToLower()) -bor $Partial.IsPresent)
				{
					$ufilename = Get-UniqueFileName -FileName $FileName
					Set-content -path $ufilename "BEGIN:VCARD"
					add-content -path $ufilename "VERSION:2.1"
					$givenName = ""
					if ($ncCol.Contact.GivenName -ne $null)
					{
						$givenName = $ncCol.Contact.GivenName
					}
					$surname = ""
					if ($ncCol.Contact.Surname -ne $null)
					{
						$surname = $ncCol.Contact.Surname
					}
					add-content -path $ufilename ("N:" + $surname + ";" + $givenName)
					add-content -path $ufilename ("FN:" + $ncCol.Contact.DisplayName)
					$Department = "";
					if ($ncCol.Contact.Department -ne $null)
					{
						$Department = $ncCol.Contact.Department
					}
					
					$CompanyName = "";
					if ($ncCol.Contact.CompanyName -ne $null)
					{
						$CompanyName = $ncCol.Contact.CompanyName
					}
					add-content -path $ufilename ("ORG:" + $CompanyName + ";" + $Department)
					if ($ncCol.Contact.JobTitle -ne $null)
					{
						add-content -path $ufilename ("TITLE:" + $ncCol.Contact.JobTitle)
					}
					if ($ncCol.Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone] -ne $null)
					{
						add-content -path $ufilename ("TEL;CELL;VOICE:" + $ncCol.Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone])
					}
					if ($ncCol.Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::HomePhone] -ne $null)
					{
						add-content -path $ufilename ("TEL;HOME;VOICE:" + $ncCol.Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::HomePhone])
					}
					if ($ncCol.Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone] -ne $null)
					{
						add-content -path $ufilename ("TEL;WORK;VOICE:" + $ncCol.Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone])
					}
					if ($ncCol.Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessFax] -ne $null)
					{
						add-content -path $ufilename ("TEL;WORK;FAX:" + $ncCol.Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessFax])
					}
					if ($ncCol.Contact.BusinessHomePage -ne $null)
					{
						add-content -path $ufilename ("URL;WORK:" + $ncCol.Contact.BusinessHomePage)
					}
					if ($ncCol.Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business] -ne $null)
					{
						if ($ncCol.Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].CountryOrRegion -ne $null)
						{
							$Country = $ncCol.Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].CountryOrRegion.Replace("`n", "")
						}
						if ($ncCol.Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].City -ne $null)
						{
							$City = $ncCol.Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].City.Replace("`n", "")
						}
						if ($ncCol.Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].Street -ne $null)
						{
							$Street = $ncCol.Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].Street.Replace("`n", "")
						}
						if ($ncCol.Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].State -ne $null)
						{
							$State = $ncCol.Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].State.Replace("`n", "")
						}
						if ($ncCol.Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].PostalCode -ne $null)
						{
							$PCode = $ncCol.Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].PostalCode.Replace("`n", "")
						}
						$addr = "ADR;WORK;PREF:;" + $Country + ";" + $Street + ";" + $City + ";" + $State + ";" + $PCode + ";" + $Country
						add-content -path $ufilename $addr
					}
					if ($ncCol.Contact.ImAddresses[[Microsoft.Exchange.WebServices.Data.ImAddressKey]::ImAddress1] -ne $null)
					{
						add-content -path $ufilename ("X-MS-IMADDRESS:" + $ncCol.Contact.ImAddresses[[Microsoft.Exchange.WebServices.Data.ImAddressKey]::ImAddress1])
					}
					add-content -path $ufilename ("EMAIL;PREF;INTERNET:" + $ncCol.Mailbox.Address)
					
					
					if ($IncludePhoto)
					{
						$PhotoURL = Get-AutoDiscoverPhotoURL -EmailAddress $MailboxName -service $service
						$PhotoSize = "HR120x120"
						$PhotoURL = $PhotoURL + "/GetUserPhoto?email=" + $ncCol.Mailbox.Address + "&size=" + $PhotoSize;
						$wbClient = new-object System.Net.WebClient
						$creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(), $Credentials.GetNetworkCredential().password.ToString())
						$wbClient.Credentials = $creds
						$photoBytes = $wbClient.DownloadData($PhotoURL);
						add-content -path $ufilename "PHOTO;ENCODING=BASE64;TYPE=JPEG:"
						$ImageString = [System.Convert]::ToBase64String($photoBytes, [System.Base64FormattingOptions]::InsertLineBreaks)
						add-content -path $ufilename $ImageString
						add-content -path $ufilename "`r`n"
					}
					add-content -path $ufilename "END:VCARD"
					Write-Host "Contact exported to $ufilename"
				}
			}
		}
	}
}
