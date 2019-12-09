function Export-EXCDumpsterContacts
{
<#
	.SYNOPSIS
		Exports all Contacts from a Dumpster folder in a Mailbox using the  Exchange Web Services API to a VCF File
	
	.DESCRIPTION
		Exports all Contacts from a Dumpster folder in a Mailbox using the  Exchange Web Services API to a VCF File
		
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
	Example 1 To Export the Mailbox's RecoverableItemsDeletions folder contacts to a CSV File
	Export-EXCDumpsterContacts -MailboxName mailbox@domain.com -ExportAsCSV
	
	Example 2 RecoverableItemsPurges
	Export-EXCDumpsterContacts -MailboxName mailbox@domain.com -ExportAsCSV -Purges

	Example 3 To Export the Mailbox's RecoverableItemsDeletions folder contacts to a single VCF File
	Export-EXCDumpsterContacts -MailboxName mailbox@domain.com -ExportAsCSV
	
	Example 4 RecoverableItemsPurges
	Export-EXCDumpsterContacts -MailboxName mailbox@domain.com -ExportAsCSV -Purges
	

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
		
		[Parameter(Position = 4, Mandatory = $False)]
		[switch]
		$Purges,

		[Parameter(Position = 5, Mandatory = $false)]
		[switch]
		$ExportAsCSV,		
		
		[Parameter(Position = 6, Mandatory = $False)]
		[switch]
		$ModernAuth,
		
		[Parameter(Position = 7, Mandatory = $False)]
		[String]
		$ClientId
		
	)
	Begin
	{
		if($ExportAsCSV.IsPresent){
			$Contacts = Get-EXCDumpsterContacts -MailboxName $MailboxName -Credentials $Credentials -Purges $Purges.IsPresent -ModernAuth:$ModernAuth.IsPresent -ClientId $ClientId

		}else{
			$Contacts = Get-EXCDumpsterContacts -MailboxName $MailboxName -Credentials $Credentials -Purges $Purges.IsPresent -ForExportToVcf -ModernAuth:$ModernAuth.IsPresent -ClientId $ClientId
		}
		$ExportCollection = @()
		$FileName = Get-UniqueFileName -FileName $FileName
		if($ExportAsCSV.IsPresent){
		}else{
			$AppendStream = new-object System.IO.FileStream($FileName,[System.IO.FileMode]::Append)
		}		
		$Contacts | ForEach-Object{
			$Contact = $_
			if($ExportAsCSV.IsPresent){
				$csvEntry = Invoke-ContactToCSVEntry -Contact $Contact
				$ExportCollection += $csvEntry
			}else{
				$AppendStream.Write($Contact.MimeContent.Content, 0, $Contact.MimeContent.Content.Length);
				write-host ("Exporting : " + $Contact.Subject)
			}

		}
		if($ExportAsCSV.IsPresent){
			$ExportCollection | export-csv -NoTypeInformation -Path $FileName
		}else{
			$AppendStream.Close();
			$AppendStream.Dispose()
		}
		write-host "Exported $FileName"		
		
		
	}
}
