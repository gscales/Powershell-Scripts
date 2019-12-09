function Get-EXCDumpsterContacts
{
<#
	.SYNOPSIS
		Gets Contacts from the RecoverItems Folder in a Mailbox using the  Exchange Web Services API
	
	.DESCRIPTION
		Gets Contacts from the RecoverItems Folder in a Mailbox using the  Exchange Web Services API
		
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
		Example 1 To get a Contact from a Mailbox's RecoverableItemsDeletions folder
		Get-EXCDumpsterContacts -MailboxName mailbox@domain.com
		
		Example 2 RecoverableItemsPurges
		Get-EXCDumpsterContacts -MailboxName mailbox@domain.com -Purges

#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[System.Management.Automation.PSCredential]
		$Credentials,
		

		[Parameter(Position = 4, Mandatory = $false)]
		[switch]
		$useImpersonation,

		[Parameter(Position = 5, Mandatory = $false)]
		[switch]
		$ForExportToVcf,

		[Parameter(Position = 6, Mandatory = $False)]
		[switch]
		$ModernAuth,

		[Parameter(Position = 7, Mandatory = $False)]
		[switch]
		$Purges,
		
		[Parameter(Position = 8, Mandatory = $False)]
		[String]
		$ClientId
	)
	Process
	{
		#Connect
		$service = Connect-EXCExchange -MailboxName $MailboxName -Credential $Credentials -ModernAuth:$ModernAuth.IsPresent -ClientId $ClientId 
		if ($useImpersonation)
		{
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}
		if($Purges.IsPresent){
			$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::RecoverableItemsPurges, $MailboxName)

		}else{
			$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::RecoverableItemsDeletions, $MailboxName)
		}
    	$Contacts = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)  
		if ($service.URL)
		{
			$SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass, "IPM.Contact")
			#Define ItemView to retrive just 1000 Items    			
			$ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
			if($ForExportToVcf.IsPresent){
				$ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(100)
			}
			$fiItems = $null
			do
			{
				$fiItems = $service.FindItems($Contacts.Id, $SfSearchFilter, $ivItemView)
				Write-Verbose("Retrieved " + $fiItems.Items.Count + " of " + $fiItems.TotalCount + " OffSet " + $ivItemView.Offset)
				if ($fiItems.Items.Count -gt 0)
				{
					$psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
					$psPropset.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text
					if($ForExportToVcf.IsPresent){
						$psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::MimeContent);
					}
					[Void]$service.LoadPropertiesForItems($fiItems, $psPropset)
					foreach ($Item in $fiItems.Items)
					{
						Write-Output $Item
					}
				}
				$ivItemView.Offset += $fiItems.Items.Count
			}
			while ($fiItems.MoreAvailable -eq $true)
			
		}
	}
}
