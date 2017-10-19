Function Get-NumbersInString
{
	<#
		.SYNOPSIS
			Utility function that strips non-numbers out of a string.
		
		.DESCRIPTION
			Utility function that strips non-numbers out of a string.
		
		.PARAMETER InStr
			The string that needs its non-numbers stripped
		
		.EXAMPLE
			PS C:\> Get-NumbersInString -InStr "abc 1234 def 5678"
	
			Will return 12345678
	#>
	[CmdletBinding()]
	Param (
		[string]
		$InStr
	)
	
	$Out = $InStr -replace "[^\d]"
	try { return [int]$Out }
	catch { }
	try { return [uint64]$Out }
	catch { return 0 }
}
