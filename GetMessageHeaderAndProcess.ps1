function Get-MessageHeaderAndProcess {
    [CmdletBinding()]
    param (       
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $MailboxName,

        [Parameter(Position = 2, Mandatory = $true)]
        [String]
        $MessageId,

        [Parameter(Position = 3, Mandatory = $true)]
        [String]
        $ClientId


    )
    Begin {
        if([String]::IsNullOrEmpty($ClientId)){
            $ClientId = "b3738173-a400-47f4-96f9-56163a84910f"
        }		
        Import-Module .\Microsoft.IdentityModel.Clients.ActiveDirectory.dll -Force
        $PromptBehavior = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters -ArgumentList Auto       
        $Context = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext("https://login.microsoftonline.com/common")
        $token = ($Context.AcquireTokenAsync("https://graph.microsoft.com", "b3738173-a400-47f4-96f9-56163a84910f", "urn:ietf:wg:oauth:2.0:oob", $PromptBehavior)).Result
        $Header = @{
            'Content-Type'  = 'application\json'
            'Authorization' = $token.CreateAuthorizationHeader()
        }        
        $Result =  Invoke-RestMethod -Headers $Header -Uri ("https://graph.microsoft.com/beta/users('" + $MailboxName + "')/Messages?`$Top=1000&`$select=ReceivedDateTime,Sender,Subject,IsRead,inferenceClassification,InternetMessageId,parentFolderId,hasAttachments,webLink,InternetMessageHeaders&`$filter=internetMessageId eq '" + $MessageId + "'") -Method Get 
        if ($Result.value -ne $null) {
            foreach ($Message in $Result.value ) {
                Invoke-ProcessAntiSPAMHeaders -Item $Message
                write-output $Message
            }
        }
       	
		
    }
}

function Invoke-ProcessAntiSPAMHeaders {
    [CmdletBinding()] 
    param (
        [Parameter(Position = 1, Mandatory = $false)]
        [psobject]
        $Item
    )
	
    process {
            $IndexedHeaders = New-Object 'system.collections.generic.dictionary[string,string]'
            if ([bool]($Item.PSobject.Properties.name -match "InternetMessageHeaders"))
            {
                
                foreach($header in $Item.InternetMessageHeaders){
                    if(!$IndexedHeaders.ContainsKey($header.name)){
                        $IndexedHeaders.Add($header.name,$header.value)
                    }
                }
            }
            if($IndexedHeaders.ContainsKey("Authentication-Results")){
                $AuthResultsText = $IndexedHeaders["Authentication-Results"]
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
            if($IndexedHeaders.ContainsKey("Authentication-Results-Original")){
                $AuthResultsText = $IndexedHeaders["Authentication-Results-Original"]
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
            if ($IndexedHeaders.ContainsKey("X-Microsoft-Antispam")){
                $ASReport = $IndexedHeaders["X-Microsoft-Antispam"]              
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
            if ($IndexedHeaders.ContainsKey("X-Forefront-Antispam-Report")){
                $ASReport = $IndexedHeaders["X-Forefront-Antispam-Report"]              
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
                $PTRResults = [regex]::Match($ASReport,("PTR\:(.*?)\;"))
                if($PTRResults.Groups.Count -gt 0){
                    $PTR =  $PTRResults.Groups[1].Value    
                }   
                $CIPResults = [regex]::Match($ASReport,("CIP\:(.*?)\;"))
                if($CIPResults.Groups.Count -gt 0){
                    $CIP =  $CIPResults.Groups[1].Value    
                }      
                $IPVResults = [regex]::Match($ASReport,("IPV\:(.*?)\;"))
                if($IPVResults.Groups.Count -gt 0){
                    $IPV =  $IPVResults.Groups[1].Value    
                }                   
                Add-Member -InputObject $Item -NotePropertyName "CTRY" -NotePropertyValue $CTRY  -Force
                Add-Member -InputObject $Item -NotePropertyName "SFV" -NotePropertyValue $SFV  -Force
                Add-Member -InputObject $Item -NotePropertyName "SRV" -NotePropertyValue $SRV  -Force
                Add-Member -InputObject $Item -NotePropertyName "PTR" -NotePropertyValue $PTR  -Force
                Add-Member -InputObject $Item -NotePropertyName "CIP" -NotePropertyValue $CIP  -Force
                Add-Member -InputObject $Item -NotePropertyName "IPV" -NotePropertyValue $IPV  -Force
            }
            
            if ($IndexedHeaders.ContainsKey("X-MS-Exchange-Organization-SCL")){
                Add-Member -InputObject $Item -NotePropertyName "SCL" -NotePropertyValue $IndexedHeaders["X-MS-Exchange-Organization-SCL"]  -Force 
            }
            if ($IndexedHeaders.ContainsKey("X-CustomSpam")){
                Add-Member -InputObject $Item -NotePropertyName "ASF" -NotePropertyValue $IndexedHeaders["X-CustomSpam"]  -Force 
            }
                
               
            

            
               
		
    }
}