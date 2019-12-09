function New-EXCContact
{
<#
	.SYNOPSIS
		Creates a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API
	
	.DESCRIPTION
		Creates a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API
		
		Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951
	
	.PARAMETER MailboxName
		A description of the MailboxName parameter.
	
	.PARAMETER DisplayName
		A description of the DisplayName parameter.
	
	.PARAMETER FirstName
		A description of the FirstName parameter.
	
	.PARAMETER LastName
		A description of the LastName parameter.
	
	.PARAMETER EmailAddress
		A description of the EmailAddress parameter.
	
	.PARAMETER CompanyName
		A description of the CompanyName parameter.
	
	.PARAMETER Credentials
		A description of the Credentials parameter.
	
	.PARAMETER Department
		A description of the Department parameter.
	
	.PARAMETER Office
		A description of the Office parameter.
	
	.PARAMETER BusinssPhone
		A description of the BusinssPhone parameter.
	
	.PARAMETER MobilePhone
		A description of the MobilePhone parameter.
	
	.PARAMETER HomePhone
		A description of the HomePhone parameter.
	
	.PARAMETER IMAddress
		A description of the IMAddress parameter.
	
	.PARAMETER Street
		A description of the Street parameter.
	
	.PARAMETER City
		A description of the City parameter.
	
	.PARAMETER State
		A description of the State parameter.
	
	.PARAMETER PostalCode
		A description of the PostalCode parameter.
	
	.PARAMETER Country
		A description of the Country parameter.
	
	.PARAMETER JobTitle
		A description of the JobTitle parameter.
	
	.PARAMETER Notes
		A description of the Notes parameter.
	
	.PARAMETER Photo
		A description of the Photo parameter.
	
	.PARAMETER FileAs
		A description of the FileAs parameter.
	
	.PARAMETER WebSite
		A description of the WebSite parameter.
	
	.PARAMETER Title
		A description of the Title parameter.
	
	.PARAMETER Folder
		A description of the Folder parameter.
	
	.PARAMETER EmailAddressDisplayAs
		A description of the EmailAddressDisplayAs parameter.
	
	.PARAMETER useImpersonation
		A description of the useImpersonation parameter.
	
	.EXAMPLE
		Example 1 To create a contact in the default contacts folder
		New-EXCContact -Mailboxname mailbox@domain.com -EmailAddress contactEmai@domain.com -FirstName John -LastName Doe -DisplayName "John Doe"
		
	.EXAMPLE
		Example 2 To create a contact and add a contact picture
		New-EXCContact -Mailboxname mailbox@domain.com -EmailAddress contactEmai@domain.com -FirstName John -LastName Doe -DisplayName "John Doe" -photo 'c:\photo\Jdoe.jpg'
		
	.EXAMPLE
		Example 3 To create a contact in a user created subfolder
		New-EXCContact -Mailboxname mailbox@domain.com -EmailAddress contactEmai@domain.com -FirstName John -LastName Doe -DisplayName "John Doe" -Folder "\MyCustomContacts"
		
		This cmdlet uses the EmailAddress as unique key so it wont let you create a contact with that email address if one already exists.
#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$DisplayName,
		
		[Parameter(Position = 2, Mandatory = $true)]
		[string]
		$FirstName,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[string]
		$LastName,
		
		[Parameter(Position = 4, Mandatory = $true)]
		[string]
		$EmailAddress,
		
		[Parameter(Position = 5, Mandatory = $false)]
		[string]
		$CompanyName,
		
		[Parameter(Position = 6, Mandatory = $false)]
		[System.Management.Automation.PSCredential]
		$Credentials,
		
		[Parameter(Position = 7, Mandatory = $false)]
		[string]
		$Department,
		
		[Parameter(Position = 8, Mandatory = $false)]
		[string]
		$Office,
		
		[Parameter(Position = 9, Mandatory = $false)]
		[string]
		$BusinssPhone,
		
		[Parameter(Position = 10, Mandatory = $false)]
		[string]
		$MobilePhone,
		
		[Parameter(Position = 11, Mandatory = $false)]
		[string]
		$HomePhone,
		
		[Parameter(Position = 12, Mandatory = $false)]
		[string]
		$IMAddress,
		
		[Parameter(Position = 13, Mandatory = $false)]
		[string]
		$Street,
		
		[Parameter(Position = 14, Mandatory = $false)]
		[string]
		$City,
		
		[Parameter(Position = 15, Mandatory = $false)]
		[string]
		$State,
		
		[Parameter(Position = 16, Mandatory = $false)]
		[string]
		$PostalCode,
		
		[Parameter(Position = 17, Mandatory = $false)]
		[string]
		$Country,
		
		[Parameter(Position = 18, Mandatory = $false)]
		[string]
		$JobTitle,
		
		[Parameter(Position = 19, Mandatory = $false)]
		[string]
		$Notes,
		
		[Parameter(Position = 20, Mandatory = $false)]
		[string]
		$Photo,
		
		[Parameter(Position = 21, Mandatory = $false)]
		[string]
		$FileAs,
		
		[Parameter(Position = 22, Mandatory = $false)]
		[string]
		$WebSite,
		
		[Parameter(Position = 23, Mandatory = $false)]
		[string]
		$Title,
		
		[Parameter(Position = 24, Mandatory = $false)]
		[string]
		$Folder,
		
		[Parameter(Position = 25, Mandatory = $false)]
		[string]
		$EmailAddressDisplayAs,
		
		[Parameter(Position = 26, Mandatory = $false)]
		[switch]
		$useImpersonation,

		[Parameter(Position = 27, Mandatory = $False)]
		[switch]
		$ModernAuth,
		
		[Parameter(Position = 28, Mandatory = $False)]
		[String]
		$ClientId
		
		
		
	)
	Begin
	{
		#Connect
		$service = Connect-EXCExchange -MailboxName $MailboxName -Credential $Credentials -ModernAuth:$ModernAuth.IsPresent -ClientId $ClientId
		if ($useImpersonation.IsPresent)
		{
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}
		$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts, $MailboxName)
		if ($Folder)
		{
			$Contacts = Get-EXCContactFolder -Service $service -FolderPath $Folder -SmptAddress $MailboxName
		}
		else
		{
			$Contacts = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderid)
		}
		if ($service.URL)
		{
			$type = ("System.Collections.Generic.List" + '`' + "1") -as "Type"
			$type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.FolderId" -as "Type")
			$ParentFolderIds = [Activator]::CreateInstance($type)
			$ParentFolderIds.Add($Contacts.Id)
			$Error.Clear();
			$cnpsPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
			$ncCol = $service.ResolveName($EmailAddress, $ParentFolderIds, [Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryThenContacts, $true, $cnpsPropset);
			$createContactOkay = $false
			if ($Error.Count -eq 0)
			{
				if ($ncCol.Count -eq 0)
				{
					$createContactOkay = $true;
				}
				else
				{
					foreach ($Result in $ncCol)
					{
						if ($Result.Contact -eq $null)
						{
							Write-host "Contact already exists $($Result.Mailbox.Name)"
							throw "Contact already exists"
						}
						else
						{
							if ((Test-EmailAddress -EmailAddress $EmailAddress))
							{
								if ($Result.Mailbox.MailboxType -eq [Microsoft.Exchange.WebServices.Data.MailboxType]::Mailbox)
								{
									$UserDn = Get-UserDN -service $service -EmailAddress $Result.Mailbox.Address
									$cnpsPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
									$ncCola = $service.ResolveName($UserDn, $ParentFolderIds, [Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::ContactsOnly, $true, $cnpsPropset);
									if ($ncCola.Count -eq 0)
									{
										$createContactOkay = $true;
									}
									else
									{
										Write-Host -ForegroundColor Red -Object "Number of existing Contacts Found $($ncCola.Count)"
										foreach ($Result in $ncCola)
										{
											Write-Host -ForegroundColor Red -Object $Result.Mailbox.Name
										}
										throw "Contact already exists"
									}
								}
							}
							else
							{
								Write-Host -ForegroundColor Yellow "Email Address is not valid for GAL match"
							}
						}
					}
				}
				if ($createContactOkay)
				{
					$Contact = New-Object Microsoft.Exchange.WebServices.Data.Contact -ArgumentList $service
					#Set the GivenName
					$Contact.GivenName = $FirstName
					#Set the LastName
					$Contact.Surname = $LastName
					#Set Subject  
					$Contact.Subject = $DisplayName
					$Contact.FileAs = $DisplayName
					if ($Title -ne "")
					{
						$PR_DISPLAY_NAME_PREFIX_W = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3A45, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);
						$Contact.SetExtendedProperty($PR_DISPLAY_NAME_PREFIX_W, $Title)
					}
					$Contact.CompanyName = $CompanyName
					$Contact.DisplayName = $DisplayName
					$Contact.Department = $Department
					$Contact.OfficeLocation = $Office
					$Contact.CompanyName = $CompanyName
					$Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone] = $BusinssPhone
					$Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone] = $MobilePhone
					$Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::HomePhone] = $HomePhone
					$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business] = New-Object  Microsoft.Exchange.WebServices.Data.PhysicalAddressEntry
					$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].Street = $Street
					$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].State = $State
					$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].City = $City
					$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].CountryOrRegion = $Country
					$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].PostalCode = $PostalCode
					$Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1] = $EmailAddress
					if ([string]::IsNullOrEmpty($EmailAddressDisplayAs) -eq $false)
					{
						$Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Name = $EmailAddressDisplayAs
					}
					$Contact.ImAddresses[[Microsoft.Exchange.WebServices.Data.ImAddressKey]::ImAddress1] = $IMAddress
					$Contact.FileAs = $FileAs
					$Contact.BusinessHomePage = $WebSite
					#Set any Notes  
					$Contact.Body = $Notes
					$Contact.JobTitle = $JobTitle
					if ($Photo)
					{
						$fileAttach = $Contact.Attachments.AddFileAttachment($Photo)
						$fileAttach.IsContactPhoto = $true
					}
					$Contact.Save($Contacts.Id)
					Write-Host "Contact Created"
				}
			}
		}
	}
}
