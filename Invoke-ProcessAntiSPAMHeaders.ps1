function Invoke-ProcessAntiSPAMHeaders {
    [CmdletBinding()] 
    param (
        [Parameter(Position = 1, Mandatory = $false)]
        [psobject]
        $Item
    )
	
    process {
           if ([bool]($Item.PSobject.Properties.name -match "IndexedInternetMessageHeaders"))
            {
                if($Item.IndexedInternetMessageHeaders.ContainsKey("Authentication-Results")){
                    $AuthResultsText = $Item.IndexedInternetMessageHeaders["Authentication-Results"]
                    $SPFResults =  [regex]::Match($AuthResultsText,("spf=(.*?)dkim="))
                    if($SPFResults.Groups.Count -gt 0){
                       $SPF =  $SPFResults.Groups[1].Value    
                    }
                    $DKIMResults =  [regex]::Match($AuthResultsText,("dkim=(.*?)dmarc="))
                    if($DKIMResults.Groups.Count -gt 0){
                       $DKIM =  $DKIMResults.Groups[1].Value    
                    }
                    $DMARCResults =  [regex]::Match($AuthResultsText,("dmarc=(.*?)compauth="))
                    if($DMARCResults.Groups.Count -gt 0){
                       $DMARC =  $DMARCResults.Groups[1].Value    
                    }
                    $CompAuthResults =  [regex]::Match($AuthResultsText,("compauth=(.*)"))
                    if($CompAuthResults.Groups.Count -gt 0){
                       $CompAuth =  $CompAuthResults.Groups[1].Value    
                    }
                    Add-Member -InputObject $Item -NotePropertyName "SPF" -NotePropertyValue $SPF -Force
                    Add-Member -InputObject $Item -NotePropertyName "DKIM" -NotePropertyValue $DKIM  -Force
                    Add-Member -InputObject $Item -NotePropertyName "DMARC" -NotePropertyValue $DMARC  -Force
                    Add-Member -InputObject $Item -NotePropertyName "CompAuth" -NotePropertyValue $CompAuth  -Force
                 }
                if($Item.IndexedInternetMessageHeaders.ContainsKey("Authentication-Results-Original")){
                    $AuthResultsText = $Item.IndexedInternetMessageHeaders["Authentication-Results-Original"]
                    $SPFResults =  [regex]::Match($AuthResultsText,("spf=(.*?)\;"))
                    if($SPFResults.Groups.Count -gt 0){
                       $SPF =  $SPFResults.Groups[1].Value    
                    }
                    $DKIMResults =  [regex]::Match($AuthResultsText,("dkim=(.*?)\;"))
                    if($DKIMResults.Groups.Count -gt 0){
                       $DKIM =  $DKIMResults.Groups[1].Value    
                    }
                    $DMARCResults =  [regex]::Match($AuthResultsText,("dmarc=(.*?)\;"))
                    if($DMARCResults.Groups.Count -gt 0){
                       $DMARC =  $DMARCResults.Groups[1].Value    
                    }
                    $CompAuthResults =  [regex]::Match($AuthResultsText,("compauth=(.*)"))
                    if($CompAuthResults.Groups.Count -gt 0){
                       $CompAuth =  $CompAuthResults.Groups[1].Value    
                    }
                    Add-Member -InputObject $Item -NotePropertyName "Original-SPF" -NotePropertyValue $SPF -Force
                    Add-Member -InputObject $Item -NotePropertyName "Original-DKIM" -NotePropertyValue $DKIM  -Force
                    Add-Member -InputObject $Item -NotePropertyName "Original-DMARC" -NotePropertyValue $DMARC  -Force
                    Add-Member -InputObject $Item -NotePropertyName "Original-CompAuth" -NotePropertyValue $CompAuth  -Force
                }
                if ($Item.IndexedInternetMessageHeaders.ContainsKey("X-Microsoft-Antispam")){
                    $ASReport = $Item.IndexedInternetMessageHeaders["X-Microsoft-Antispam"]              
                    $PCLResults = [regex]::Match($ASReport,("PCL\:(.*?)\;"))
                    if($PCLResults.Groups.Count -gt 0){
                        $PCL =  $PCLResults.Groups[1].Value    
                    }
                    $BCLResults = [regex]::Match($ASReport,("BCL\:(.*?)\;"))
                    if($BCLResults.Groups.Count -gt 0){
                        $BCL =  $BCLResults.Groups[1].Value    
                    }
                    Add-Member -InputObject $Item -NotePropertyName "PCL" -NotePropertyValue $PCL  -Force
                    Add-Member -InputObject $Item -NotePropertyName "BCL" -NotePropertyValue $BCL  -Force
                }
                if ($Item.IndexedInternetMessageHeaders.ContainsKey("X-Forefront-Antispam-Report")){
                    $ASReport = $Item.IndexedInternetMessageHeaders["X-Forefront-Antispam-Report"]              
                    $CTRYResults = [regex]::Match($ASReport,("CTRY\:(.*?)\;"))
                    if($CTRYResults.Groups.Count -gt 0){
                        $CTRY =  $CTRYResults.Groups[1].Value    
                    }
                    $SFVResults = [regex]::Match($ASReport,("SFV\:(.*?)\;"))
                    if($SFVResults.Groups.Count -gt 0){
                        $SFV =  $SFVResults.Groups[1].Value    
                    }
                    $SRVResults = [regex]::Match($ASReport,("SRV\:(.*?)\;"))
                    if($SRVResults.Groups.Count -gt 0){
                        $SRV =  $SRVResults.Groups[1].Value    
                    }
                    Add-Member -InputObject $Item -NotePropertyName "CTRY" -NotePropertyValue $CTRY  -Force
                    Add-Member -InputObject $Item -NotePropertyName "SFV" -NotePropertyValue $SFV  -Force
                    Add-Member -InputObject $Item -NotePropertyName "SRV" -NotePropertyValue $SRV  -Force
                }
                
                if ($Item.IndexedInternetMessageHeaders.ContainsKey("X-MS-Exchange-Organization-SCL")){
                    Add-Member -InputObject $Item -NotePropertyName "SCL" -NotePropertyValue $Item.IndexedInternetMessageHeaders["X-MS-Exchange-Organization-SCL"]  -Force 
                }
                if ($Item.IndexedInternetMessageHeaders.ContainsKey("X-CustomSpam")){
                    Add-Member -InputObject $Item -NotePropertyName "ASF" -NotePropertyValue $Item.IndexedInternetMessageHeaders["X-CustomSpam"]  -Force 
                }
                
               
            }

            
               
		
    }
}