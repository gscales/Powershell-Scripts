function Invoke-ListFAIItems {

    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string]
        $Folder,		
	
        [Parameter(Position = 1, Mandatory = $false)]
        [Microsoft.Exchange.WebServices.Data.ExchangeService]
        $Service, 

        [Parameter(Position = 2, Mandatory = $true)]
        [String]
        $MailboxName,  

        
     

        [Parameter(Position = 5, Mandatory = $false)]
	    [System.Management.Automation.PSCredential]
	    $Credentials


    )
    process {
        $service = Connect-FAIExchange -MailboxName $MailboxName -Credentials $Credentials 
        ## Find and Bind to Folder based on Path  
        #Define the path to search should be seperated with \  
        #Bind to the MSGFolder Root  
  
        $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$Folder, $MailboxName)
        
        if ($useImpersonation.IsPresent) {
            $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)        
        }
        $service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName);  
        #Check to see if it exists and display a better error if it doesn't
        $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000) 
        $ivItemView.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated
        $fiItems = $service.FindItems($folderid, $ivItemView) 
        if ($fiItems.Items.Count -gt 0) {
            foreach($Item in $fiItems){
                if($Item.ItemClass -match "IPM.Configuration"){
                    $rptObj = "" | Select ConfigItemName,ItemClass,Subject
                    $rptObj.ConfigItemName = $Item.ItemClass.Replace("IPM.Configuration.","")
                    $rptObj.ItemClass = $Item.ItemClass
                    $rptObj.Subject = $Item.Subject
                    Write-Output $rptObj
                }
            }
        }
        else {
            write-host ("No Objects in Folder")	
            return $null		
        }
    }
}
