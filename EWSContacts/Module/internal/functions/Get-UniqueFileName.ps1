function Get-UniqueFileName
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$FileName
	)
	process
	{
		
		$directoryName = [System.IO.Path]::GetDirectoryName($FileName)
		$FileDisplayName = [System.IO.Path]::GetFileNameWithoutExtension($FileName);
		$FileExtension = [System.IO.Path]::GetExtension($FileName);
		for ($i = 1;; $i++)
		{
			
			if (-not (Test-Path $FileName))
			{
				return $FileName
			}
			else
			{
				$FileName = [System.IO.Path]::Combine($directoryName, $FileDisplayName + "(" + $i + ")" + $FileExtension);
			}
			
			if ($i -eq 10000) { throw "Out of Range" }
		}
	}
}
