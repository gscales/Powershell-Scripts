function Search-EXCContactsForCCNumbers
{
<#
	.SYNOPSIS
		Search Contacts in a Contact folder in a Mailbox using the  Exchange Web Services API
	
	.DESCRIPTION
		Searches Contact in a Contact folder in a Mailbox using the  Exchange Web Services API
		
		Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951
	
	.PARAMETER MailboxName
		A description of the MailboxName parameter.
	
	.PARAMETER Credentials
		A description of the Credentials parameter.
	
	.PARAMETER Folder
		A description of the Folder parameter.
	
	.PARAMETER useImpersonation
		A description of the useImpersonation parameter.
	
	.EXAMPLE
		PS C:\> Search-EXCContactsForCCNumbers -MailboxName 'value1' -Credentials $Credentials
#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[System.Management.Automation.PSCredential]
		$Credentials,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[string]
		$Folder,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[switch]
		$useImpersonation,

		[Parameter(Position = 4, Mandatory = $true)][String]$CrediCardValdatorDLLPath,

		[Parameter(Position = 5, Mandatory = $False)]
		[switch]
		$ModernAuth,
		
		[Parameter(Position = 6, Mandatory = $False)]
		[String]
		$ClientId

	)
	Begin
	{
		$Script:rptCollection = @()
		import-module path $CrediCardValdatorDLLPath
		#Connect
		$service = Connect-EXCExchange -MailboxName $MailboxName -Credential $Credentials -ModernAuth:$ModernAuth.IsPresent -ClientId $ClientId
		if ($useImpersonation.IsPresent)
		{
			$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}
		$folderid = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts, $MailboxName)
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
			$SfSearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass, "IPM.Contact")
			#Define ItemView to retrive just 1000 Items    
			$ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
			$fiItems = $null
			do
			{
				$fiItems = $service.FindItems($Contacts.Id, $SfSearchFilter, $ivItemView)
				if ($fiItems.Items.Count -gt 0)
				{
					$psPropset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
					[Void]$service.LoadPropertiesForItems($fiItems, $psPropset)
					foreach ($Contact in $fiItems.Items)
					{
						if ($Contact -is [Microsoft.Exchange.WebServices.Data.Contact])
						{
							$DnName = $Contact.DisplayName
							Write-Verbose "Processing $DnName"
							$BusinssPhone = $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone]
							$MobilePhone = $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone]
							$HomePhone = $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::HomePhone]
							if ($BusinssPhone -ne $null)
							{
								Write-Verbose 'Check BusinessPhone'
								$CheckObj = Find-CCNumber -Number $BusinssPhone -Property "BusinessPhone" -MailboxName $MailboxName -DisplayName $DnName
							}
							if ($MobilePhone -ne $null)
							{
								Write-Verbose 'Check MobilePhone'
								$CheckObj = Find-CCNumber -Number $MobilePhone -Property "MobilePhone" -MailboxName $MailboxName -DisplayName $DnName
								
							}
							if ($HomePhone -ne $null)
							{
								Write-Verbose 'Check HomePhone'
								$CheckObj = Find-CCNumber -Number $HomePhone -Property "HomePhone" -MailboxName $MailboxName -DisplayName $DnName
								
							}
							$Email1 = $Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1]
							$Email2 = $Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress2]
							$Email3 = $Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress3]
							if ($Email1 -ne $null)
							{
								Write-Verbose 'Check Email1'
								if (![string]::IsNullOrEmpty($Email1.Address))
								{
									if ([String]::IsNullOrEmpty($DnName))
									{
										$DnName = $Email1.Address
									}
									$CheckObj = Find-CCNumber -Number $Email1.Address -Property "Email" -MailboxName $MailboxName -DisplayName $DnName
								}
								
							}
							if ($Email2 -ne $null)
							{
								Write-Verbose 'Check Email2'
								if (![string]::IsNullOrEmpty($Email2.Address))
								{
									$CheckObj = Find-CCNumber -Number $Email2.Address -Property "Email2" -MailboxName $MailboxName -DisplayName $DnName
								}
								
							}
							if ($Email3 -ne $null)
							{
								Write-Verbose 'Check Email3'
								if (![string]::IsNullOrEmpty($Email3.Address))
								{
									$CheckObj = Find-CCNumber -Number $Email3.Address -Property "Email3" -MailboxName $MailboxName -DisplayName $DnName
								}
							}
							
						}
					}
				}
				$ivItemView.Offset += $fiItems.Items.Count
			}
			while ($fiItems.MoreAvailable -eq $true)
			
		}
		write-output $Script:rptCollection
	}
}
