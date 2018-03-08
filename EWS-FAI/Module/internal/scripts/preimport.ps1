# Add all things you want to run before importing the main code
		$EWSDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services' | Sort-Object Name -Descending | Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
		if (Test-Path $EWSDLL)
		{
			Import-Module $EWSDLL
		}
		else
		{
			"$(get-date -format yyyyMMddHHmmss):"
			"This script requires the EWS Managed API 1.2 or later."
			"Please download and install the current version of the EWS Managed API from"
			"http://go.microsoft.com/fwlink/?LinkId=255472"
			""
			"Exiting Script."
			$exception = New-Object System.Exception ("Managed Api missing")
			throw $exception
		}