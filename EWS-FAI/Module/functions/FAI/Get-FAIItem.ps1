function Get-FAIItem {

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
        
        [Parameter(Position = 3, Mandatory = $true)]
        [String]
        $ConfigItemName,        
       
        [Parameter(Position = 5, Mandatory = $false)]
		[System.Management.Automation.PSCredential]
        $Credentials,
        
        [Parameter(Position = 6, Mandatory = $false)]
		[switch]
        $ReturnConfigObject
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
        $sfFolderSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass, ("IPM.Configuration." + $ConfigItemName)) 
        $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1) 
        $ivItemView.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated
        $fiItems = $service.FindItems($folderid, $sfFolderSearchFilter, $ivItemView) 
        if ($fiItems.Items.Count -eq 1) {
            $UsrConfig = [Microsoft.Exchange.WebServices.Data.UserConfiguration]::Bind($service, $ConfigItemName, $folderid, [Microsoft.Exchange.WebServices.Data.UserConfigurationProperties]::All)
            if($ReturnConfigObject.IsPresent){
                return $UsrConfig
            }
            else{
                if($UsrConfig.Dictionary -ne $null){
                    return $UsrConfig.Dictionary
                } 
                else{
                    if($UsrConfig.XMLData -ne $null){
                        $XML = [System.Text.Encoding]::UTF8.GetString($UsrConfig.XmlData)  
                        $XML = $XML.Replace("\xEF\xBB\xBF", "")   
                        return [XML]$XML
                    }
                    else{
                        $UsrConfig
                    }                
                } 
            }            
        }
        else {
            write-host ("No Object in Folder")	
            return $null		
        }
    }
}
