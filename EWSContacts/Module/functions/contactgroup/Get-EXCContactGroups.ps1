function Get-EXCContactGroups
{
<#
	.SYNOPSIS
		Gets a Contact Groups in a Contact folder in a Mailbox using the  Exchange Web Services API
	
	.DESCRIPTION
		Gets a Contact Group in a Contact folder in a Mailbox using the  Exchange Web Services API
		
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
		Example 1 To Get a Contact Group in the default contacts folder
		Get-EXCContactGroups  -Mailboxname mailbox@domain.com s

#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[System.Management.Automation.PSCredential]
		$Credentials,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[string]
		$Folder,		
	
		[Parameter(Position = 6, Mandatory = $false)]
		[switch]
		$useImpersonation,

		[Parameter(Position = 7, Mandatory = $False)]
		[switch]
		$ModernAuth,
		
		[Parameter(Position = 8, Mandatory = $False)]
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
			$SfSearchFilter2 = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass, "IPM.DistList")


			#Define ItemView to retrive just 1000 Items    
			$ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
			$fiItems = $null
			do
			{
				$fiItems = $service.FindItems($Contacts.Id, $SfSearchFilter2, $ivItemView)
				if ($fiItems.Item.Count -eq 0)
				{
					Write-Verbose "No Groups Found"
				}
				#[Void]$service.LoadPropertiesForItems($fiItems,$psPropset)  
				foreach ($Item in $fiItems.Items)
				{
					Write-Output $Item
				}
				$ivItemView.Offset += $fiItems.Items.Count
			}
			while ($fiItems.MoreAvailable -eq $true)
		}
	}
}
