function Find-CCNumber
{
<#
	.SYNOPSIS
		A brief description of the Find-CCNumber function.
	
	.DESCRIPTION
		A detailed description of the Find-CCNumber function.
	
	.PARAMETER Number
		A description of the Number parameter.
	
	.PARAMETER Property
		A description of the Property parameter.
	
	.PARAMETER MailboxName
		A description of the MailboxName parameter.
	
	.PARAMETER DisplayName
		A description of the DisplayName parameter.
	
	.EXAMPLE
		PS C:\> Find-CCNumber -Number 'value1' -Property 'value2' -MailboxName 'value3' -DisplayName 'value4'
#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$Number,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$Property,
		
		[Parameter(Position = 2, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[string]
		$DisplayName
	)
	Begin
	{
		$Number = Get-NumbersInString -InStr $Number
		$Number = $Number.Replace("-", "").Trim()
		if ($Number -ne 0)
		{
			write-host $Number
			if ($Number.Length -gt 4)
			{
				$detector = new-object CreditCardValidator.CreditCardDetector($Number)
				if ($detector.IsValid())
				{
					$rptObj = "" | Select-Object Mailbox, Contact, Property, Number, Brand, BrandName, IssuerCategory
					$rptObj.Mailbox = $MailboxName
					$rptObj.Contact = $DisplayName
					$rptObj.Property = $Property
					$rptObj.Number = $Number
					$rptObj.Brand = $detector.Brand
					$rptObj.BrandName = $detector.BrandName
					$rptObj.IssuerCategory = $detector.IssuerCategory
					$Script:rptCollection += $rptObj
				}
				else
				{
					$SSN_Regex = "^(?!000)([0-6]\d{2}|7([0-6]\d|7[012]))([ -]?)(?!00)\d\d\3(?!0000)\d{4}$"
					$Matches = $Number | Select-String -Pattern $SSN_Regex
					if ($Matches.Matches.Count -gt 0)
					{
						$rptObj = "" | Select-Object Mailbox, Contact, Property, Number, Brand, BrandName, IssuerCategory
						$rptObj.Mailbox = $MailboxName
						$rptObj.Contact = $DisplayName
						$rptObj.Property = $Property
						$rptObj.Number = $Number
						$rptObj.Brand = "Social Security Number"
						$Script:rptCollection += $rptObj
					}
				}
			}
		}
		return $detector
	}
}
