function Get-EXCAllContactFolders
{
	[CmdletBinding()]
	param (

		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[System.Management.Automation.PSCredential]
		$Credentials,	
	
		[switch]
		$useImpersonation,

		[Parameter(Position = 5, Mandatory = $False)]
		[switch]
		$ModernAuth,
		
		[Parameter(Position = 6, Mandatory = $False)]
		[String]
		$ClientId
	)
	process
	{
		$service = Connect-EXCExchange -MailboxName $MailboxName -Credentials $Credentials -ModernAuth:$ModernAuth.IsPresent -ClientId $ClientId
		if ($useImpersonation.IsPresent)
		{
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}
		## Find and Bind to Folder based on Path  
		#Define the path to search should be seperated with \  
		#Bind to the MSGFolder Root  
	    $PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
       
		$fvFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)  
        #Deep Transval will ensure all folders in the search path are returned  
		$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;
		$PR_Folder_Path = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
        #Add Properties to the  Property Set  
		$PropertySet.Add($PR_Folder_Path);   
		$fvFolderView.PropertySet = $PropertySet 
		$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $MailboxName)
		$tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $folderid)
		$fiResult = $null  
		$Folders = @()
        #The Do loop will handle any paging that is required if there are more the 1000 folders in a mailbox  
        do {  
            $fiResult = $tfTargetFolder.FindFolders($fvFolderView)  
            foreach ($ffFolder in $fiResult.Folders) {  
				if($ffFolder.FolderClass -ne $null){
					
					if($ffFolder.FolderClass -match "IPF.Contact"){
						$foldpathval = $null  
						#Try to get the FolderPath Value and then covert it to a usable String   
						if ($ffFolder.TryGetProperty($PR_Folder_Path, [ref] $foldpathval)) {  
							$binarry = [Text.Encoding]::UTF8.GetBytes($foldpathval)  
							$hexArr = $binarry | ForEach-Object { $_.ToString("X2") }  
							$hexString = $hexArr -join ''  
							$hexString = $hexString.Replace("FEFF", "5C00")  
							$fpath = ConvertToString($hexString)  
						}  
						$ffFolder | Add-Member -Name "FolderPath" -Value $fpath -MemberType NoteProperty
						$ffFolder | Add-Member -Name "Mailbox" -Value $ParentFolder.Mailbox -MemberType NoteProperty
						$Folders += $ffFolder
					}	
				}
			}
            $fvFolderView.Offset += $fiResult.Folders.Count
        }while ($fiResult.MoreAvailable -eq $true)  
        return, $Folders	
	}
}
function ConvertToString($ipInputString) {  
    $Val1Text = ""  
    for ($clInt = 0; $clInt -lt $ipInputString.length; $clInt++) {  
        $Val1Text = $Val1Text + [Convert]::ToString([Convert]::ToChar([Convert]::ToInt32($ipInputString.Substring($clInt, 2), 16)))  
        $clInt++  
    }  
    return $Val1Text  
} 