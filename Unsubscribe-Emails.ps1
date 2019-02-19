
function Invoke-UnsubscribeEmail {
    [CmdletBinding()]
    param (       
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $MailboxName,

        [Parameter(Position = 2, Mandatory = $false)]
        [switch]
        $UnSubscribe
    )
    Begin {		
        $UnSubribeHash = @{}
        Import-Module .\Microsoft.IdentityModel.Clients.ActiveDirectory.dll -Force
        $PromptBehavior = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters -ArgumentList Auto       
        $Context = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext("https://login.microsoftonline.com/common")
        $token = ($Context.AcquireTokenAsync("https://graph.microsoft.com", "d3590ed6-52b3-4102-aeff-aad2292ab01c", "urn:ietf:wg:oauth:2.0:oob", $PromptBehavior)).Result
        $Header = @{
            'Content-Type'  = 'application\json'
            'Authorization' = $token.CreateAuthorizationHeader()
        }        
        $Result =  Invoke-RestMethod -Headers $Header -Uri ("https://graph.microsoft.com/beta/users('" + $MailboxName + "')/MailFolders/Inbox/Messages?`$Top=1000&`$select=ReceivedDateTime,Sender,Subject,IsRead,inferenceClassification,InternetMessageId,parentFolderId,hasAttachments,webLink,unsubscribeEnabled,unsubscribeData") -Method Get 
        if ($Result.value -ne $null) {
            foreach ( $Message in $Result.value ) {
                if($Message.unsubscribeEnabled){
                   foreach($Entry in $Message.unsubscribeData){
                       if($Entry.contains("mailto:")){
                           if(!$UnSubribeHash.ContainsKey($Entry))
                           {
                                $UnSubribeHash.Add($Entry,"")
                                if($UnSubscribe.IsPresent){
                                    $UnsubsribeResult = Invoke-RestMethod -Headers $Header -Uri ("https://graph.microsoft.com/beta/users('" + $MailboxName + "')/MailFolders/Inbox/Messages('" + $Message.id + "')/unsubscribe") -Method Post -ContentType "application/json" 
                                    write-host ("Unsubscribe : " + $Message.Subject)
                                    $UnsubsribeResult
                                }else{
                                    write-host ("ReportOnly - Unsubscribe : " + $Message.Subject)
                                }

                           } 
                       }
                   }
                }   
            }
        }
       	
		
    }
}
