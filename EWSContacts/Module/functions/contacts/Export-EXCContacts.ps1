function Export-EXCContacts
{
<#
	.SYNOPSIS
		Exports all Contacts from a Contact folder in a Mailbox using the  Exchange Web Services API to a VCF File
	
	.DESCRIPTION
		Exports all Contacts from a Contact folder in a Mailbox using the  Exchange Web Services API to a VCF File
		
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
		Example 1 To Export all contacts to local file vcf file
		Export-EXCContacts -MailboxName mailbox@domain.com  -FileName c:\export\filename.vcf
		If the file already exists it will handle creating a unique filename
		
	.EXAMPLE
		Example 2 To export from a contacts subfolder use
		Export-EXCContacts -MailboxName mailbox@domain.com  -FileName c:\export\filename.vcf -folder \contacts\subfolder

	.EXAMPLE
		Example 3 To Export a contact to local csv file
		Export-EXCContacts -MailboxName mailbox@domain.com  -FileName c:\export\filename.vcf -ExportAsCSV
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
		
		[Parameter(Position = 4, Mandatory = $false)]
		[string]
		$Folder,
		
		[Parameter(Position = 5, Mandatory = $false)]
		[switch]
		$Partial,

		[Parameter(Position = 6, Mandatory = $false)]
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
		if ($Folder)
		{
			if ($Partial.IsPresent)
			{
				$Contacts = Get-EXCContacts -MailboxName $MailboxName -Credentials $Credentials -Folder $Folder -Partial -ForExportToVcf -ModernAuth:$ModernAuth.IsPresent -ClientId $ClientId
			}
			else
			{
				$Contacts = Get-EXCContacts -MailboxName $MailboxName -Credentials $Credentials -Folder $Folder -ForExportToVcf -ModernAuth:$ModernAuth.IsPresent -ClientId $ClientId
			}
		}
		else
		{
			if ($Partial.IsPresent)
			{
				$Contacts = Get-EXCContacts -MailboxName $MailboxName -Credentials $Credentials -Partial -ForExportToVcf -ModernAuth:$ModernAuth.IsPresent -ClientId $ClientId
			}
			else
			{ 
				$Contacts =  Get-EXCContacts -MailboxName $MailboxName -Credentials $Credentials -ForExportToVcf -ModernAuth:$ModernAuth.IsPresent -ClientId $ClientId
			}
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
