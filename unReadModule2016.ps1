function Connect-Exchange { 
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 1, Mandatory = $true)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position = 2, Mandatory = $false)] [string]$url
    )  
    Begin {
        Load-EWSManagedAPI
		
        ## Set Exchange Version  
        if ([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016) {
            $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016
        }
        else {
            $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
        }        
		
		  
        ## Create Exchange Service Object  
        $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)  
		  
        ## Set Credentials to use two options are availible Option1 to use explict credentials or Option 2 use the Default (logged On) credentials  
		  
        #Credentials Option 1 using UPN for the windows Account  
        #$psCred = Get-Credential  
        $creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(), $Credentials.GetNetworkCredential().password.ToString())  
        $service.Credentials = $creds      
        #Credentials Option 2  
        #service.UseDefaultCredentials = $true  
        #$service.TraceEnabled = $true
        ## Choose to ignore any SSL Warning issues caused by Self Signed Certificates  
		  
        Handle-SSL	
		  
        ## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use  
		  
        #CAS URL Option 1 Autodiscover  
        if ($url) {
            $uri = [system.URI] $url
            $service.Url = $uri    
        }
        else {
            $service.AutodiscoverUrl($MailboxName, { $true })  
        }
        Write-host ("Using CAS Server : " + $Service.url)   
		   
        #CAS URL Option 2 Hardcoded  
		  
        #$uri=[system.URI] "https://casservername/ews/exchange.asmx"  
        #$service.Url = $uri    
		  
        ## Optional section for Exchange Impersonation  
		  
        #$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
        if (!$service.URL) {
            throw "Error connecting to EWS"
        }
        else {		
            return $service
        }
    }
}

function Load-EWSManagedAPI {
    param( 
    )  
    Begin {
        if (Test-Path ($script:ModuleRoot + "/Microsoft.Exchange.WebServices.dll")) {
            Import-Module ($script:ModuleRoot + "/Microsoft.Exchange.WebServices.dll")
            $Script:EWSDLL = $script:ModuleRoot + "/Microsoft.Exchange.WebServices.dll"
            write-verbose ("Using EWS dll from Local Directory")
        }
        else {

			
            ## Load Managed API dll  
            ###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
            $EWSDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
            if (Test-Path $EWSDLL) {
                Import-Module $EWSDLL
                $Script:EWSDLL = $EWSDLL 
            }
            else {
                "$(get-date -format yyyyMMddHHmmss):"
                "This script requires the EWS Managed API 1.2 or later."
                "Please download and install the current version of the EWS Managed API from"
                "http://go.microsoft.com/fwlink/?LinkId=255472"
                ""
                "Exiting Script."
                exit


            } 
        }
    }
}

function Handle-SSL {
    param( 
    )  
    Begin {
        ## Code From http://poshcode.org/624
        ## Create a compilation environment
        $Provider = New-Object Microsoft.CSharp.CSharpCodeProvider
        $Compiler = $Provider.CreateCompiler()
        $Params = New-Object System.CodeDom.Compiler.CompilerParameters
        $Params.GenerateExecutable = $False
        $Params.GenerateInMemory = $True
        $Params.IncludeDebugInformation = $False
        $Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

        $TASource = @'
  namespace Local.ToolkitExtensions.Net.CertificatePolicy{
    public class TrustAll : System.Net.ICertificatePolicy {
      public TrustAll() { 
      }
      public bool CheckValidationResult(System.Net.ServicePoint sp,
        System.Security.Cryptography.X509Certificates.X509Certificate cert, 
        System.Net.WebRequest req, int problem) {
        return true;
      }
    }
  }
'@ 
        $TAResults = $Provider.CompileAssemblyFromSource($Params, $TASource)
        $TAAssembly = $TAResults.CompiledAssembly

        ## We now create an instance of the TrustAll and attach it to the ServicePointManager
        $TrustAll = $TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
        [System.Net.ServicePointManager]::CertificatePolicy = $TrustAll

        ## end code from http://poshcode.org/624

    }
}

function CovertBitValue($String) {  
    $numItempattern = '(?=\().*(?=bytes)'  
    $matchedItemsNumber = [regex]::matches($String, $numItempattern)   
    $Mb = [INT64]$matchedItemsNumber[0].Value.Replace("(", "").Replace(",", "")  
    return [math]::round($Mb / 1048576, 0)  
}  

function Get-UnReadMessageCount {
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 1, Mandatory = $true)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position = 2, Mandatory = $false)] [switch]$useImpersonation,
        [Parameter(Position = 3, Mandatory = $false)] [string]$url,
        [Parameter(Position = 4, Mandatory = $true)] [Int32]$Months
    )  
    Begin {
        $eval1 = "Last" + $Months + "MonthsTotal"
        $eval2 = "Last" + $Months + "MonthsUnread"
        $eval3 = "Last" + $Months + "MonthsSent"
        $eval4 = "Last" + $Months + "MonthsReplyToSender"
        $eval5 = "Last" + $Months + "MonthsReplyToAll"
        $eval6 = "Last" + $Months + "MonthForward"
        $reply = 0;
        $replyall = 0
        $forward = 0
        $rptObj = "" | select  MailboxName, Mailboxsize, LastLogon, LastLogonAccount, $eval1, $eval2, $eval4, $eval5, $eval6, LastMailRecieved, $eval3, LastMailSent  
        $rptObj.MailboxName = $MailboxName  
        if ($url) {
            $service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials -url $url 
        }
        else {
            $service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials
        }
        if ($useImpersonation.IsPresent) {
            $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
        }
        $AQSString1 = "System.Message.DateReceived:>" + [system.DateTime]::Now.AddMonths(-$Months).ToString("yyyy-MM-dd")   
        $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $MailboxName)     
        $Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderid)  
        $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems, $MailboxName)     
        $SentItems = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderid)    
        $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)  
        $psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)  
        $psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived)
        $psPropset.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsRead)
        $PidTagLastVerbExecuted = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x1081, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer); 
        $psPropset.Add($PidTagLastVerbExecuted)
        $ivItemView.PropertySet = $psPropset
        $MailboxStats = Get-MailboxStatistics $MailboxName  
        $ts = CovertBitValue($MailboxStats.TotalItemSize.ToString())  
        write-host ("Total Size : " + $MailboxStats.TotalItemSize) 
        $rptObj.MailboxSize = $ts  
        write-host ("Last Logon Time : " + $MailboxStats.LastLogonTime) 
        $rptObj.LastLogon = $MailboxStats.LastLogonTime  
        write-host ("Last Logon Account : " + $MailboxStats.LastLoggedOnUserAccount ) 
        $rptObj.LastLogonAccount = $MailboxStats.LastLoggedOnUserAccount  
        $fiItems = $null
        $unreadCount = 0
        $settc = $true
        do { 
            $fiItems = $Inbox.findItems($AQSString1, $ivItemView)  
            if ($settc) {
                $rptObj.$eval1 = $fiItems.TotalCount  
                write-host ("Last " + $Months + " Months : " + $fiItems.TotalCount)
                if ($fiItems.TotalCount -gt 0) {  
                    write-host ("Last Mail Recieved : " + $fiItems.Items[0].DateTimeReceived ) 
                    $rptObj.LastMailRecieved = $fiItems.Items[0].DateTimeReceived  
                }		    
                $settc = $false
            }
            foreach ($Item in $fiItems.Items) {
                $unReadVal = $null
                if ($Item.TryGetProperty([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsRead, [ref]$unReadVal)) {
                    if (!$unReadVal) {
                        $unreadCount++
                    }
                } 
                $lastVerb = $null
                if ($Item.TryGetProperty($PidTagLastVerbExecuted, [ref]$lastVerb)) {
                    switch ($lastVerb) {
                        102 { $reply++ }
                        103 { $replyall++ }
                        104 { $forward++ }
                    }
                } 
            }    
            $ivItemView.Offset += $fiItems.Items.Count    
        }while ($fiItems.MoreAvailable -eq $true) 

        write-host ("Last " + $Months + " Months Unread : " + $unreadCount ) 
        $rptObj.$eval2 = $unreadCount  
        $rptObj.$eval4 = $reply
        $rptObj.$eval5 = $replyall
        $rptObj.$eval6 = $forward
        $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1)  
        $fiResults = $SentItems.findItems($AQSString1, $ivItemView)  
        write-host ("Last " + $Months + " Months Sent : " + $fiResults.TotalCount  )
        $rptObj.$eval3 = $fiResults.TotalCount  
        if ($fiResults.TotalCount -gt 0) {  
            write-host ("Last Mail Sent Date : " + $fiResults.Items[0].DateTimeSent  )
            $rptObj.LastMailSent = $fiResults.Items[0].DateTimeSent  
        }  
        Write-Output $rptObj  
    }
}

function Get-UnReadCountOnFolder {
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 1, Mandatory = $true)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position = 2, Mandatory = $false)] [switch]$useImpersonation,
        [Parameter(Position = 3, Mandatory = $false)] [string]$url
    )  
    Begin {
        if ($url) {
            $service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials -url $url 
        }
        else {
            $service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials
        }
        if ($useImpersonation.IsPresent) {
            $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
        }
        $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $MailboxName)     
        $Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderid) 
        Write-Host ("Total Message Count : " + $Inbox.TotalCount)
        Write-Host ("Total Unread Message Count : " + $Inbox.UnreadCount)
        $mbcomb = "" | select EmailAddress, Unread, TotalCount
        $mbcomb.EmailAddress = $MailboxName
        $mbcomb.Unread = $Inbox.UnreadCount
        $mbcomb.TotalCount = $Inbox.TotalCount
        return $mbcomb
    }
}

function Mark-AllMessagesUnread {
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 1, Mandatory = $true)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position = 2, Mandatory = $false)] [switch]$useImpersonation,
        [Parameter(Position = 3, Mandatory = $false)] [string]$url,
        [Parameter(Position = 4, Mandatory = $true)] [string]$FolderPath
    )  
    Begin {
        if ($url) {
            $service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials -url $url 
        }
        else {
            $service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials
        }
        if ($useImpersonation.IsPresent) {
            $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
        }
        $folderId = FolderIdFromPath -FolderPath $FolderPath -SmtpAddress $MailboxName
        if ($folderId -ne $null) {
            $Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderId) 
            Write-Host ("Total Message Count : " + $Inbox.TotalCount)
            Write-Host ("Total Unread Message Count : " + $Inbox.UnreadCount)
            Write-Host ("Marking all messags a unread in folder")
            $Folder.MarkAllItemsAsRead($true)
        }
    }	
}
function FolderIdFromPath {
    param (
        $FolderPath = "$( throw 'Folder Path is a mandatory Parameter' )",
        $SmtpAddress = "$( throw 'Folder Path is a mandatory Parameter' )"
		  )
    process {
        ## Find and Bind to Folder based on Path  
        #Define the path to search should be seperated with \  
        #Bind to the MSGFolder Root  
        $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $SmtpAddress)   
        $tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderid)  
        #Split the Search path into an array  
        $fldArray = $FolderPath.Split("\") 
        #Loop through the Split Array and do a Search for each level of folder 
        for ($lint = 1; $lint -lt $fldArray.Length; $lint++) { 
            #Perform search based on the displayname of each folder level 
            $fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1) 
            $SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $fldArray[$lint]) 
            $findFolderResults = $service.FindFolders($tfTargetFolder.Id, $SfSearchFilter, $fvFolderView) 
            if ($findFolderResults.TotalCount -gt 0) { 
                foreach ($folder in $findFolderResults.Folders) { 
                    $tfTargetFolder = $folder                
                } 
            } 
            else { 
                "Error Folder Not Found"  
                $tfTargetFolder = $null  
                break  
            }     
        }  
        if ($tfTargetFolder -ne $null) {
            return $tfTargetFolder.Id
        }
        else {
            throw "Folder not found"
        }
    }
}


function Mark-LastMessageReadinInbox {
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 1, Mandatory = $true)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position = 2, Mandatory = $false)] [switch]$useImpersonation,
        [Parameter(Position = 3, Mandatory = $false)] [string]$url
    )  
    Begin {
        if ($url) {
            $service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials -url $url 
        }
        else {
            $service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials
        }
        if ($useImpersonation.IsPresent) {
            $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
        }
        $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $MailboxName)     
        $Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderid) 
        $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1)
        $fiItems = $service.FindItems($Inbox.Id, $ivItemView)
        if ($fiItems.Items.Count -eq 1) {
            $fiItems.Items[0].isRead = $true
            $fiItems.Items[0].Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve);					
        }     
    }
}