function Export-EXCRootContacts
{
<#
	.SYNOPSIS
		Exports all Contacts from any Non_IPM_Subtree (root) folder in a Mailbox using the  Exchange Web Services API to a VCF File
	
	.DESCRIPTION
		Exports all Contacts from any Non_IPM_Subtree (root) folderin a Mailbox using the  Exchange Web Services API to a VCF File
		
		Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951
	
	.PARAMETER MailboxName
		A description of the MailboxName parameter.
	
	.PARAMETER EmailAddress
		A description of the EmailAddress parameter.
	
	.PARAMETER Credentials
		A description of the Credentials parameter.
	
	.PARAMETER FileName
		A description of the FileName parameter.
	
	.PARAMETER Folder
		A description of the Folder parameter.
	
	.PARAMETER Partial
		A description of the Partial parameter.
	
	.EXAMPLE
		Example 1 To Export a contact to local file
		Export-EXCRootContacts -MailboxName mailbox@domain.com  -FolderName AllContacts -FileName c:\export\filename.vcf
		If the file already exists it will handle creating a unique filename
		
#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,		
	
		[Parameter(Position = 2, Mandatory = $false)]
		[System.Management.Automation.PSCredential]
		$Credentials,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[string]
		$FileName,
		
		[Parameter(Position = 4, Mandatory = $true)]
		[string]
		$FolderName,
		
		[Parameter(Position = 5, Mandatory = $false)]
		[switch]
		$ExportAsCSV,		
				
		[Parameter(Position = 7, Mandatory = $False)]
		[switch]
		$ModernAuth,
		
		[Parameter(Position = 8, Mandatory = $False)]
		[String]
		$ClientId
	)
	Begin
	{
		$ExportCollection = @()
		$DupTrack = @{}
		$FileName = Get-UniqueFileName -FileName $FileName
		if($ExportAsCSV.IsPresent){
			$Contacts =  Get-EXCContacts -MailboxName $MailboxName -Credentials $Credentials -RootFolderName $FolderName  -ModernAuth:$ModernAuth.IsPresent -ClientId $ClientId
		}else{
			$Contacts =  Get-EXCContacts -MailboxName $MailboxName -Credentials $Credentials -RootFolderName $FolderName -ForExportToVcf  -ModernAuth:$ModernAuth.IsPresent -ClientId $ClientId
			$AppendStream = new-object System.IO.FileStream($FileName,[System.IO.FileMode]::Append)
		}		
		$Contacts | ForEach-Object{
			$Contact = $_

				$emailVal = $null;
				if($Contact.EmailAddresses.TryGetValue([Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1, [ref]$emailVal)){
					if($emailVal.Address -ne $null){
						if(!$DupTrack.ContainsKey($emailVal.Address)){
							$DupTrack.Add($emailVal.Address,"")
							if($ExportAsCSV.IsPresent){
								$csvEntry = Invoke-ContactToCSVEntry -Contact $Contact
								$ExportCollection += $csvEntry
							}else{
								$AppendStream.Write($Contact.MimeContent.Content, 0, $Contact.MimeContent.Content.Length);
							}							
							write-host ("Exporting : " + $emailVal.Address)
						}else{
							write-Host ("Skipping duplicate" + $emailVal)
						}
					}
				}
				else{
					Write-Host ("No email found for " + $Contact.Subject)
				}
			

		}
		if($ExportAsCSV.IsPresent){
			$ExportCollection | export-csv -NoTypeInformation -Path $FileName -Encoding UTF8 
		}else{
			$AppendStream.Close();
			$AppendStream.Dispose()
		}
		write-host "Exported $FileName"		
		
	}
}
