function Get-EXCContacts
{
<#
	.SYNOPSIS
		Gets a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API
	
	.DESCRIPTION
		Gets a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API
		
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
		Example 1 To get a Contact from a Mailbox's default contacts folder
		Get-EXCContacts -MailboxName mailbox@domain.com
		
	.EXAMPLE
		Example 2 To get all the Contacts from subfolder of the Mailbox's default contacts folder
		Get-EXCContacts -MailboxName mailbox@domain.com -Folder \Contact\test
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
		[string]
		$RootFolderName,

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
		$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts, $MailboxName)
		if ($Folder)
		{
			$Contacts = Get-EXCContactFolder -Service $service -FolderPath $Folder -SmptAddress $MailboxName
		}
		else
		{
			if(![String]::IsNullOrEmpty($RootFolderName)){
				$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)   
				$rfRootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)  
				$fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1) 
				$SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$RootFolderName) 
				$findFolderResults = $service.FindFolders($rfRootFolder.Id,$SfSearchFilter,$fvFolderView) 
				if($findFolderResults.Folders.Count -eq 1){
					$Contacts = $findFolderResults.Folders[0]
				}else{
					throw "No Folder found"
				}
			}else{
				$Contacts = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderid)
			}
		}
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
