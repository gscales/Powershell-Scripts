function Get-EXCContactFolder
{
	<#
		.SYNOPSIS
			Returns the folder object of the specified path.
		
		.DESCRIPTION
			Returns the folder object of the specified path.
		
		.PARAMETER FolderPath
			The path to the folder, relative to the message folder base.
		
		.PARAMETER SmptAddress
			The email address of the mailbox to access
		
		.PARAMETER Service
			The established Service connection to use for this connection.
			Use 'Connect-EXCExchange' in order to establish a connection and obtain such an object.
		
		.EXAMPLE
			PS C:\> Get-EXCContactFolder -FolderPath 'Contacts\Private' -SmptAddress 'peter@example.com' -Service $Service
	
			Returns the 'Private' folder within the contacts folder for the mailbox peter@example.com
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$FolderPath,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$SmptAddress,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[Microsoft.Exchange.WebServices.Data.ExchangeService]
		$Service, 

		[Parameter(Position = 3, Mandatory = $false)]
		[String]
		$MailboxName,

		[Parameter(Position = 4, Mandatory = $False)]
		[switch]
		$ModernAuth,
		
		[Parameter(Position = 5, Mandatory = $False)]
		[String]
		$ClientId
	)
	process
	{
		if(![String]::IsNullOrEmpty($MailboxName)){
			$service = Connect-EXCExchange -MailboxName $MailboxName -Credentials $Credentials -ModernAuth:$ModernAuth.IsPresent -ClientId $ClientId
		}
		## Find and Bind to Folder based on Path  
		#Define the path to search should be seperated with \  
		#Bind to the MSGFolder Root  
		$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $SmptAddress)
		$tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $folderid)
		#Split the Search path into an array  
		$fldArray = $FolderPath.Split("\")
		#Loop through the Split Array and do a Search for each level of folder 
		for ($lint = 1; $lint -lt $fldArray.Length; $lint++)
		{
			#Perform search based on the displayname of each folder level 
			$fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1)
			$SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $fldArray[$lint])
			$findFolderResults = $Service.FindFolders($tfTargetFolder.Id, $SfSearchFilter, $fvFolderView)
			if ($findFolderResults.TotalCount -gt 0)
			{
				foreach ($folder in $findFolderResults.Folders)
				{
					$tfTargetFolder = $folder
				}
			}
			else
			{
				Write-host "Error Folder Not Found check path and try again"
				$tfTargetFolder = $null
				break
			}
		}
		if ($tfTargetFolder -ne $null)
		{
			return [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $tfTargetFolder.Id)
		}
		else
		{
			throw "Folder Not found"
		}
	}
}
