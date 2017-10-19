Function Test-EmailAddress
{
	<#
		.SYNOPSIS
			Tests, whether an email address is a legal email address.
		
		.DESCRIPTION
			Tests, whether an email address is a legal email address.
		
		.PARAMETER EmailAddress
			The address to verify.
		
		.EXAMPLE
			PS C:\> Test-EmailAddress -EmailAddress 'info@example.com'
	
			Tests whether 'info@example.com' is a legal email address.
			Hint: It probably is
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$EmailAddress
	)
	process
	{
		try
		{
			$check = New-Object System.Net.Mail.MailAddress($EmailAddress)
			return $true
		}
		catch
		{
			return $false
		}
	}
}
