function Export-EXCContact
{
<#
	.SYNOPSIS
		Exports a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API to a VCF File
	
	.DESCRIPTION
		Exports a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API
		
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
		Export-EXCContact -MailboxName mailbox@domain.com -EmailAddress address@domain.com -FileName c:\export\filename.vcf
		If the file already exists it will handle creating a unique filename
		
	.EXAMPLE
		Example 2 To export from a contacts subfolder use
		Export-EXCContact -MailboxName mailbox@domain.com -EmailAddress address@domain.com -FileName c:\export\filename.vcf -folder \contacts\subfolder
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
		
		[Parameter(Position = 3, Mandatory = $true)]
		[string]
		$FileName,
		
		[Parameter(Position = 4, Mandatory = $false)]
		[string]
		$Folder,
		
		[Parameter(Position = 5, Mandatory = $false)]
		[switch]
		$Partial,
		

		[Parameter(Position = 6, Mandatory = $False)]
		[switch]
		$ModernAuth,
		
		[Parameter(Position = 7, Mandatory = $False)]
		[String]
		$ClientId
		
	)
	Begin
	{
		if ($Folder)
		{
			if ($Partial.IsPresent)
			{
				$Contacts = Get-EXCContact -MailboxName $MailboxName -EmailAddress $EmailAddress -Credentials $Credentials -Folder $Folder -Partial -ModernAuth:$ModernAuth.IsPresent -ClientId $ClientId
			}
			else
			{
				$Contacts = Get-EXCContact -MailboxName $MailboxName -EmailAddress $EmailAddress -Credentials $Credentials -Folder $Folder -ModernAuth:$ModernAuth.IsPresent -ClientId $ClientId
			}
		}
		else
		{
			if ($Partial.IsPresent)
			{
				$Contacts = Get-EXCContact -MailboxName $MailboxName -EmailAddress $EmailAddress -Credentials $Credentials -Partial -ModernAuth:$ModernAuth.IsPresent -ClientId $ClientId
			}
			else
			{
				$Contacts =  Get-EXCContact -MailboxName $MailboxName -EmailAddress $EmailAddress -Credentials $Credentials -ModernAuth:$ModernAuth.IsPresent -ClientId $ClientId
			}
		}
		
		$Contacts | ForEach-Object{
			$Contact = $_
			$psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
			$psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::MimeContent);
			$Contact.load($psPropset)
			$FileName = Get-UniqueFileName -FileName $FileName
			[System.IO.File]::WriteAllBytes($FileName, $Contact.MimeContent.Content)
			write-host "Exported $FileName"
		}
		
		
	}
}
