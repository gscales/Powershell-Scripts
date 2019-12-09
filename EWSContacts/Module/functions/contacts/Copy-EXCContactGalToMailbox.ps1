function Copy-EXCContactGalToMailbox
{
<#
	.SYNOPSIS
		Copies a Contact from the Global Address List to a Local Mailbox Contacts folder using the  Exchange Web Services API
	
	.DESCRIPTION
		Copies a Contact from the Global Address List to a Local Mailbox Contacts folder using the  Exchange Web Services API
		
		Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951
	
	.PARAMETER MailboxName
		A description of the MailboxName parameter.
	
	.PARAMETER EmailAddress
		A description of the EmailAddress parameter.
	
	.PARAMETER Credentials
		A description of the Credentials parameter.
	
	.PARAMETER Folder
		A description of the Folder parameter.
	
	.PARAMETER IncludePhoto
		A description of the IncludePhoto parameter.
	
	.PARAMETER useImpersonation
		A description of the useImpersonation parameter.
	
	.EXAMPLE
		Example 1 To Copy a Gal contacts to local Contacts folder
		Copy-EXCContactGalToMailbox -MailboxName mailbox@domain.com -EmailAddress email@domain.com
		
	.EXAMPLE
		Example 2 Copy a GAL contact to a Contacts subfolder
		Copy-EXCContactGalToMailbox -MailboxName mailbox@domain.com -EmailAddress email@domain.com  -Folder \Contacts\UnderContacts
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
		[string]
		$Folder,
		
		[Parameter(Position = 4, Mandatory = $false)]
		[switch]
		$IncludePhoto,
		
		[Parameter(Position = 5, Mandatory = $false)]
		[switch]
		$useImpersonation,

		[Parameter(Position = 6, Mandatory = $False)]
		[switch]
		$ModernAuth,
		
		[Parameter(Position = 7, Mandatory = $False)]
		[String]
		$ClientId
	)
	Process
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
		$Error.Clear();
		$cnpsPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
		$ncCol = $service.ResolveName($EmailAddress, $ParentFolderIds, [Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryOnly, $true, $cnpsPropset);
		if ($Error.Count -eq 0)
		{
			foreach ($Result in $ncCol)
			{
				if ($Result.Mailbox.Address.ToLower() -eq $EmailAddress.ToLower())
				{
					$type = ("System.Collections.Generic.List" + '`' + "1") -as "Type"
					$type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.FolderId" -as "Type")
					$ParentFolderIds = [Activator]::CreateInstance($type)
					$ParentFolderIds.Add($Contacts.Id)
					$Error.Clear();
					$cnpsPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
					$ncCola = $service.ResolveName($EmailAddress, $ParentFolderIds, [Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryThenContacts, $true, $cnpsPropset);
					$createContactOkay = $false
					if ($Error.Count -eq 0)
					{
						if ($ncCola.Count -eq 0)
						{
							$createContactOkay = $true;
						}
						else
						{
							foreach ($aResult in $ncCola)
							{
								if ($aResult.Contact -eq $null)
								{
									Write-host "Contact already exists " + $aResult.Contact.DisplayName
									throw ("Contact already exists")
								}
								else
								{
									if ((Test-EmailAddress -EmailAddress $Result.Mailbox.Address))
									{
										if ($Result.Mailbox.MailboxType -eq [Microsoft.Exchange.WebServices.Data.MailboxType]::Mailbox)
										{
											$UserDn = Get-UserDN -service $service -EmailAddress $Result.Mailbox.Address
											$cnpsPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
											$ncColb = $service.ResolveName($UserDn, $ParentFolderIds, [Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::ContactsOnly, $true, $cnpsPropset);
											if ($ncColb.Count -eq 0)
											{
												$createContactOkay = $true;
											}
											else
											{
												Write-Host -ForegroundColor Red ("Number of existing Contacts Found " + $ncColb.Count)
												foreach ($Result in $ncColb)
												{
													Write-Host -ForegroundColor Red ($ncColb.Mailbox.Name)
												}
												throw ("Contact already exists")
											}
										}
									}
									else
									{
										Write-Host -ForegroundColor Yellow ("Email Address is not valid for GAL match")
									}
								}
							}
						}
						if ($createContactOkay)
						{
							#check for SipAddress
							$IMAddress = ""
							$emailVal = $null;
							if ($ncCol.Contact.EmailAddresses.TryGetValue([Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1, [ref]$emailVal))
							{
								$email1 = $emailVal.Address
								if ($email1.tolower().contains("sip:"))
								{
									$IMAddress = $email1
								}
							}
							$emailVal = $null;
							if ($ncCol.Contact.EmailAddresses.TryGetValue([Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress2, [ref]$emailVal))
							{
								$email2 = $emailVal.Address
								if ($email2.tolower().contains("sip:"))
								{
									$IMAddress = $email2
									$ncCol.Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress2] = $null
								}
							}
							$emailVal = $null;
							if ($ncCol.Contact.EmailAddresses.TryGetValue([Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress3, [ref]$emailVal))
							{
								$email3 = $emailVal.Address
								if ($email3.tolower().contains("sip:"))
								{
									$IMAddress = $email3
									$ncCol.Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress3] = $null
								}
							}
							if ($IMAddress -ne "")
							{
								$ncCol.Contact.ImAddresses[[Microsoft.Exchange.WebServices.Data.ImAddressKey]::ImAddress1] = $IMAddress
							}
							$ncCol.Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress2] = $null
							$ncCol.Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress3] = $null
							$ncCol.Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address = $ncCol.Mailbox.Address.ToLower()
							$ncCol.Contact.FileAs = $ncCol.Contact.DisplayName
							if ($IncludePhoto)
							{
								$PhotoURL = Get-AutoDiscoverPhotoURL -EmailAddress $MailboxName -service $service
								$PhotoSize = "HR120x120"
								$PhotoURL = $PhotoURL + "/GetUserPhoto?email=" + $ncCol.Mailbox.Address + "&size=" + $PhotoSize;
								$wbClient = new-object System.Net.WebClient
								$creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(), $Credentials.GetNetworkCredential().password.ToString())
								$wbClient.Credentials = $creds
								$photoBytes = $wbClient.DownloadData($PhotoURL);
								$fileAttach = $ncCol.Contact.Attachments.AddFileAttachment("contactphoto.jpg", $photoBytes)
								$fileAttach.IsContactPhoto = $true
							}
							$ncCol.Contact.Save($Contacts.Id);
							Write-Host ("Contact copied")
						}
					}
				}
			}
		}
	}
}
