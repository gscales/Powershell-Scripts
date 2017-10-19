function Export-EXCContactFolder
{
	<#
		.SYNOPSIS
			Exports contacts from a mailbox/public folder to file.
		
		.DESCRIPTION
			Exports contacts from a mailbox/public folder to file.
	
			Currently, only export to CSV is being supported.
		
		.PARAMETER MailboxName
			The name of the mailbox to access.
		
		.PARAMETER Credentials
			Credentials that are authorized to access that mailbox.
		
		.PARAMETER Folder
			Default: Contacts
			The relative path to a folder with contacts within the mailbox.
			Example: Contacts\Private
		
		.PARAMETER PublicFolderPath
			The relative path to a public folder with contacts to export.
		
		.PARAMETER FileName
			Name of the file to export to.
			- File will be overwritten if it exists.
			- File needs not exist
			- Parent Folder must exist
			- User must have write access to the target path.
		
		.PARAMETER OutputType
			Default: CSV
			Currently, only CSV is supported as output type.
		
		.EXAMPLE
			PS C:\> Export-EXCContactFolder -MailboxName 'ben@example.com' -Credentials $Credentials -FileName 'C:\temp\contacts.csv'
	
			Exports the content of ben@example.com's default contacts folder to file.
	#>
	[CmdletBinding(DefaultParameterSetName = "Default")]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[System.Management.Automation.PSCredential]
		$Credentials,
		
		[Parameter(Position = 2, Mandatory = $false, ParameterSetName = "Default")]
		[string]
		$Folder,
		
		[Parameter(Position = 2, Mandatory = $true, ParameterSetName = "PublicFolder")]
		[string]
		$PublicFolderPath,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[string]
		$FileName,
		
		[ValidateSet('CSV')]
		[string]
		$OutputType = "CSV"
	)
	begin
	{
		#region Utility functions
		function Get-Contacts
		{
			[CmdletBinding()]
			param (
				[Parameter(Position = 1, Mandatory = $true)]
				[Microsoft.Exchange.WebServices.Data.Folder]
				$ContactFolder
			)
			process
			{
				$psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
				$PR_Gender = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(14925, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Short)
				$psPropset.Add($PR_Gender)
				#Define ItemView to retrive just 1000 Items      
				$ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
				$fiItems = $null
				do
				{
					$fiItems = $service.FindItems($ContactFolder.Id, $ivItemView)
					[Void]$ContactFolder.Service.LoadPropertiesForItems($fiItems, $psPropset)
					foreach ($Item in $fiItems.Items)
					{
						if ($Item -is [Microsoft.Exchange.WebServices.Data.Contact])
						{
							$expObj = "" | Select-Object DisplayName, GivenName, Surname, Gender, Email1DisplayName, Email1Type, Email1EmailAddress, BusinessPhone, MobilePhone, HomePhone, BusinessStreet, BusinessCity, BusinessState, HomeStreet, HomeCity, HomeState
							$expObj.DisplayName = $Item.DisplayName
							$expObj.GivenName = $Item.GivenName
							$expObj.Surname = $Item.Surname
							$expObj.Gender = ""
							$Gender = $null
							if ($item.TryGetProperty($PR_Gender, [ref]$Gender))
							{
								if ($Gender -eq 2)
								{
									$expObj.Gender = "Male"
								}
								if ($Gender -eq 1)
								{
									$expObj.Gender = "Female"
								}
							}
							$BusinessPhone = $null
							$MobilePhone = $null
							$HomePhone = $null
							if ($Item.PhoneNumbers -ne $null)
							{
								if ($Item.PhoneNumbers.TryGetValue([Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone, [ref]$BusinessPhone))
								{
									$expObj.BusinessPhone = $BusinessPhone
								}
								if ($Item.PhoneNumbers.TryGetValue([Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone, [ref]$MobilePhone))
								{
									$expObj.MobilePhone = $MobilePhone
								}
								if ($Item.PhoneNumbers.TryGetValue([Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::HomePhone, [ref]$HomePhone))
								{
									$expObj.HomePhone = $HomePhone
								}
							}
							if ($Item.EmailAddresses.Contains([Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1))
							{
								$expObj.Email1DisplayName = $Item.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Name
								$expObj.Email1Type = $Item.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].RoutingType
								$expObj.Email1EmailAddress = $Item.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address
							}
							$HomeAddress = $null
							$BusinessAddress = $null
							if ($item.PhysicalAddresses -ne $null)
							{
								if ($item.PhysicalAddresses.TryGetValue([Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Home, [ref]$HomeAddress))
								{
									$expObj.HomeStreet = $HomeAddress.Street
									$expObj.HomeCity = $HomeAddress.City
									$expObj.HomeState = $HomeAddress.State
								}
								if ($item.PhysicalAddresses.TryGetValue([Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business, [ref]$BusinessAddress))
								{
									$expObj.BusinessStreet = $BusinessAddress.Street
									$expObj.BusinessCity = $BusinessAddress.City
									$expObj.BusinessState = $BusinessAddress.State
								}
							}
							
							$expObj
						}
					}
					$ivItemView.Offset += $fiItems.Items.Count
				}
				while ($fiItems.MoreAvailable -eq $true)
			}
		}
		#endregion Utility functions
		
		# Connect
		$service = Connect-EXCExchange -MailboxName $MailboxName -Credentials $Credentials
	}
	process
	{
		if ($PublicFolderPath)
		{
			$service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName)
			$fldId = Get-PublicFolderIdFromPath -FolderPath $PublicFolderPath -SmtpAddress $MailboxName -service $service
			$contactsId = New-Object Microsoft.Exchange.WebServices.Data.FolderId($fldId)
			$contactFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $contactsId)
			Set-PublicFolderRoutingHeader -Service $service -Credentials $Credentials -MailboxName $MailboxName -Header "X-AnchorMailbox"
		}
		else
		{
			if ($Folder)
			{
				$contactFolder = Get-EXCContactFolder -Service $service -FolderPath $Folder -SmptAddress $MailboxName
			}
			else
			{
				$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts, $MailboxName)
				$contactFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderid)
			}
		}
		
		switch ($OutputType)
		{
			"CSV" { Get-Contacts -ContactFolder $contactFolder | Export-Csv -NoTypeInformation -Path $FileName }
			default { throw "Invalid output type: $OutputType" }
		}
	}
}
