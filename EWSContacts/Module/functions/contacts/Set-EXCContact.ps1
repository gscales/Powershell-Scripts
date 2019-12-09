function Set-EXCContact
{
<#
	.SYNOPSIS
		Updates a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API
	
	.DESCRIPTION
		Updates a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API
		
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
	
	.PARAMETER Partial
		A description of the Partial parameter.
	
	.PARAMETER force
		A description of the force parameter.
	
	.PARAMETER EmailAddressDisplayAs
		A description of the EmailAddressDisplayAs parameter.
	
	.PARAMETER useImpersonation
		A description of the useImpersonation parameter.
	
	.EXAMPLE
		Example1 Update the phone number of an existing contact
		Set-EXCContact  -Mailboxname mailbox@domain.com -EmailAddress contactEmai@domain.com -MobilePhone 023213421
		
	.EXAMPLE
		Example 2 Update the phone number of a contact in a users subfolder
		Set-EXCContact  -Mailboxname mailbox@domain.com -EmailAddress contactEmai@domain.com -MobilePhone 023213421 -Folder "\MyCustomContacts"
#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[string]
		$DisplayName,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[string]
		$FirstName,
		
		[Parameter(Position = 3, Mandatory = $false)]
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
		
		[Parameter(Mandatory = $false)]
		[switch]
		$Partial,
		
		[Parameter(Mandatory = $false)]
		[switch]
		$force,
		
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
		if ($Partial.IsPresent) { $force = $false }
		if ($Folder)
		{
			if ($Partial)
			{
				$Contacts = Get-EXCContact -MailboxName $MailboxName -EmailAddress $EmailAddress -Credentials $Credentials -Folder $Folder -Partial -ModernAuth:$ModernAuth.IsPresent -ClientId $ClientId
			}
			else
			{
				$Contacts =  Get-EXCContact -MailboxName $MailboxName -EmailAddress $EmailAddress -Credentials $Credentials -Folder $Folder -ModernAuth:$ModernAuth.IsPresent -ClientId $ClientId
			}
		}
		else
		{
			if ($Partial)
			{
				$Contacts = Get-EXCContact -MailboxName $MailboxName -EmailAddress $EmailAddress -Credentials $Credentials -Partial -ModernAuth:$ModernAuth.IsPresent -ClientId $ClientId
			}
			else
			{
				$Contacts = Get-EXCContact -MailboxName $MailboxName -EmailAddress $EmailAddress -Credentials $Credentials -ModernAuth:$ModernAuth.IsPresent -ClientId $ClientId
			}
		}
		
		foreach ($Contact in $Contacts)
		{
			if (($Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address.ToLower() -eq $EmailAddress.ToLower()) -bor $Partial.IsPresent)
			{
				$updateOkay = $false
				if ($force)
				{
					$updateOkay = $true
				}
				else
				{
					$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", ""
					$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", ""
					$choices = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
					$message = "Do you want to update contact with DisplayName $($Contact.DisplayName) : Subject-$($Contact.Subject) : Email-$($Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address)"
					$result = $Host.UI.PromptForChoice($caption, $message, $choices, 1)
					if ($result -eq 0)
					{
						$updateOkay = $true
					}
					else
					{
						Write-Host "No Action Taken"
					}
				}
				if ($updateOkay)
				{
					if ($FirstName -ne "")
					{
						$Contact.GivenName = $FirstName
					}
					if ($LastName -ne "")
					{
						$Contact.Surname = $LastName
					}
					if ($DisplayName -ne "")
					{
						$Contact.Subject = $DisplayName
					}
					if ($Title -ne "")
					{
						$PR_DISPLAY_NAME_PREFIX_W = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3A45, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);
						$Contact.SetExtendedProperty($PR_DISPLAY_NAME_PREFIX_W, $Title)
					}
					if ($CompanyName -ne "")
					{
						$Contact.CompanyName = $CompanyName
					}
					if ($DisplayName -ne "")
					{
						$Contact.DisplayName = $DisplayName
					}
					if ($Department -ne "")
					{
						$Contact.Department = $Department
					}
					if ($Office -ne "")
					{
						$Contact.OfficeLocation = $Office
					}
					if ($CompanyName -ne "")
					{
						$Contact.CompanyName = $CompanyName
					}
					if ($BusinssPhone -ne "")
					{
						$Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone] = $BusinssPhone
					}
					if ($MobilePhone -ne "")
					{
						$Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone] = $MobilePhone
					}
					if ($HomePhone -ne "")
					{
						$Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::HomePhone] = $HomePhone
					}
					if ($Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business] -eq $null)
					{
						$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business] = New-Object  Microsoft.Exchange.WebServices.Data.PhysicalAddressEntry
					}
					if ($Street -ne "")
					{
						$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].Street = $Street
					}
					if ($State -ne "")
					{
						$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].State = $State
					}
					if ($City -ne "")
					{
						$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].City = $City
					}
					if ($Country -ne "")
					{
						$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].CountryOrRegion = $Country
					}
					if ($PostalCode -ne "")
					{
						$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].PostalCode = $PostalCode
					}
					if ($EmailAddress -ne "")
					{
						$Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1] = $EmailAddress
					}
					if ([string]::IsNullOrEmpty($EmailAddressDisplayAs) -eq $false)
					{
						$Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Name = $EmailAddressDisplayAs
					}
					if ($IMAddress -ne "")
					{
						$Contact.ImAddresses[[Microsoft.Exchange.WebServices.Data.ImAddressKey]::ImAddress1] = $IMAddress
					}
					if ($FileAs -ne "")
					{
						$Contact.FileAs = $FileAs
					}
					if ($WebSite -ne "")
					{
						$Contact.BusinessHomePage = $WebSite
					}
					if ($Notes -ne "")
					{
						$Contact.Body = $Notes
					}
					if ($JobTitle -ne "")
					{
						$Contact.JobTitle = $JobTitle
					}
					if ($Photo)
					{
						$fileAttach = $Contact.Attachments.AddFileAttachment($Photo)
						$fileAttach.IsContactPhoto = $true
					}
					$Contact.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
					Write-Verbose "Contact updated $($Contact.Subject)"
					
				}
			}
		}
	}
}
