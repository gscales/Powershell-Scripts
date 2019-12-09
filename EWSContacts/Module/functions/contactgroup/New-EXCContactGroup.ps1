function New-EXCContactGroup
{
<#
	.SYNOPSIS
		Creates a Contact Group in a Contact folder in a Mailbox using the  Exchange Web Services API
	
	.DESCRIPTION
		Creates a Contact Group in a Contact folder in a Mailbox using the  Exchange Web Services API
		
		Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951
	
	.PARAMETER MailboxName
		A description of the MailboxName parameter.
	
	.PARAMETER Credentials
		A description of the Credentials parameter.
	
	.PARAMETER Folder
		A description of the Folder parameter.
	
	.PARAMETER GroupName
		A description of the GroupName parameter.
	
	.PARAMETER Members
		A description of the Members parameter.
	
	.PARAMETER useImpersonation
		A description of the useImpersonation parameter.
	
	.EXAMPLE
		Example 1 To create a Contact Group in the default contacts folder
		New-EXCContactGroup  -Mailboxname mailbox@domain.com -GroupName GroupName -Members ("member1@domain.com","member2@domain.com")
	
	.EXAMPLE
		Example 2 To create a Contact Group in a subfolder of default contacts folder
		New-EXCContactGroup  -Mailboxname mailbox@domain.com -GroupName GroupName -Folder \Contacts\Folder1 -Members ("member1@domain.com","member2@domain.com")
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
		
		[Parameter(Position = 3, Mandatory = $true)]
		[string]
		$GroupName,
		
		[Parameter(Position = 4, Mandatory = $true)]
		[PsObject]
		$Members,
		
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
		if ($useImpersonation)
		{
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}
		$folderId = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts, $MailboxName)
		if ($Folder)
		{
			$contactFolder = Get-EXCContactFolder -Service $service -FolderPath $Folder -SmptAddress $MailboxName
		}
		else
		{
			$contactFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderId)
		}
		if ($service.URL)
		{
			$contactGroup = New-Object Microsoft.Exchange.WebServices.Data.ContactGroup -ArgumentList $service
			$contactGroup.DisplayName = $GroupName
			foreach ($Member in $Members)
			{
				$contactGroup.Members.Add($Member)
			}
			$contactGroup.Save($contactFolder.Id)
			Write-Verbose "Contact Group created $GroupName"
		}
	}
}
