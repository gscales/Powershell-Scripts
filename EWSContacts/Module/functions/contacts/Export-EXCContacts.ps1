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
		Example 1 To Export a contact to local file
		Export-EXCContacts -MailboxName mailbox@domain.com  -FileName c:\export\filename.vcf
		If the file already exists it will handle creating a unique filename
		
	.EXAMPLE
		Example 2 To export from a contacts subfolder use
		Export-EXCContacts -MailboxName mailbox@domain.com  -FileName c:\export\filename.vcf -folder \contacts\subfolder
#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,		
	
		[Parameter(Position = 2, Mandatory = $true)]
		[System.Management.Automation.PSCredential]
		$Credentials,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[string]
		$FileName,
		
		[Parameter(Position = 4, Mandatory = $false)]
		[string]
		$Folder
		

		
	)
	Begin
	{
		if ($Folder)
		{
			if ($Partial.IsPresent)
			{
				$Contacts = Get-EXCContact -MailboxName $MailboxName -Credentials $Credentials -Folder $Folder -Partial
			}
			else
			{
				$Contacts = Get-EXCContact -MailboxName $MailboxName -Credentials $Credentials -Folder $Folder
			}
		}
		else
		{
			if ($Partial.IsPresent)
			{
				$Contacts = Get-EXCContacts -MailboxName $MailboxName -Credentials $Credentials -Partial
			}
			else
			{
				$Contacts =  Get-EXCContacts -MailboxName $MailboxName -Credentials $Credentials
			}
		}
		$FileName = Get-UniqueFileName -FileName $FileName
		
		$Contacts | ForEach-Object{
			$Contact = $_
			$psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
			$psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::MimeContent);
			$Contact.load($psPropset)			
			
			$AppendStream.Write($Contact.MimeContent.Content, 0, $Contact.MimeContent.Content.Length);
		}
		$AppendStream.Close();
		$AppendStream.Dispose()
		write-host "Exported $FileName"
		
		
	}
}
