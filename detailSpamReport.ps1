$rptCollection = @()
$Last24Junk = Get-MailDetailSpamReport -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date)
foreach($JunkMail in $Last24Junk){
	  $RptObject = New-Object PsObject
	  $JunkMail.psobject.properties | % {
    		$RptObject | Add-Member -NotePropertyName $_.Name -NotePropertyValue $_.Value 
	  }
	  $Details = $JunkMail | Get-MessageTraceDetail -StartDate (Get-Date).AddDays(-2) -EndDate (Get-Date) | Where-Object {$_.Event -eq "Spam"}
	  if(![String]::IsNullOrEmpty($Details.Data)){
	 	 $XMLDoc = [XML]$Details.Data
	 	 $MEPNodes = $XMLDoc.GetElementsByTagName("MEP")	  
		  for($nodeval=0;$nodeval -lt $MEPNodes.Count;$nodeval++){
			   	$Key = $MEPNodes[$nodeval].Attributes[0].Value.ToString() 	
				$Value = $MEPNodes[$nodeval].Attributes[1].Value.ToString() 	
				Add-Member -InputObject $RptObject -NotePropertyName ($Key) -NotePropertyValue ($Value)
					
	 	 }		  
	  
	  }
	  $rptCollection += $RptObject
     
}
$ScriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$filename = $ScriptPath + "\SpamReport-" + $(get-date -f yyyy-MM-dd-hh-mm-ss) + ".csv"
$rptCollection | Export-csv -NoTypeInformation -Path $filename
Write-Host ("Report written to " + $filename)
