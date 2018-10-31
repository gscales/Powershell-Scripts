function Connect-Exchange { 
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,        
        [Parameter(Position = 1, Mandatory = $false)] [string]$url,
        [Parameter(Position = 2, Mandatory = $false)] [string]$ClientId,
        [Parameter(Position = 3, Mandatory = $false)] [string]$redirectUrl,
        [Parameter(Position = 4, Mandatory = $false)] [string]$AccessToken,
        [Parameter(Position = 5, Mandatory = $false)] [switch]$basicAuth,
        [Parameter(Position = 6, Mandatory = $false)] [System.Management.Automation.PSCredential]$Credentials

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
        if ($basicAuth.IsPresent) {
            $creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(), $Credentials.GetNetworkCredential().password.ToString())  
            $service.Credentials = $creds    
        }
        else {
            if ([String]::IsNullOrEmpty($AccessToken)) {
                $Resource = "Outlook.Office365.com"    
                if ([String]::IsNullOrEmpty($ClientId)) {
                    $ClientId = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
                }
                if ([String]::IsNullOrEmpty($redirectUrl)) {
                    $redirectUrl = "urn:ietf:wg:oauth:2.0:oob"  
                }
                $Script:Token = Get-EWSAccessToken -MailboxName $MailboxName -ClientId $ClientId -redirectUrl $redirectUrl  -ResourceURL $Resource -Prompt $Prompt -CacheCredentials   
                $OAuthCredentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials((ConvertFrom-SecureStringCustom -SecureToken $Script:Token.access_token))
            }
            else {
                $OAuthCredentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($AccessToken)
            }    
            $service.Credentials = $OAuthCredentials
        }  
        #Credentials Option 1 using UPN for the windows Account  
        #$psCred = Get-Credential  
  
        #Credentials Option 2  
        #service.UseDefaultCredentials = $true  
        #$service.TraceEnabled = $true
        ## Choose to ignore any SSL Warning issues caused by Self Signed Certificates  
		  
        #Handle-SSL	
		  
        ## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use  

        #CAS URL Option 1 Autodiscover  
        if ($url) {
            $uri = [system.URI] $url
            $service.Url = $uri    
        }
        else {
            $service.AutodiscoverUrl($MailboxName, {$true})  
        }
        #Write-host ("Using CAS Server : " + $Service.url)   
		   
        #CAS URL Option 2 Hardcoded  
		  
        #$uri=[system.URI] "https://casservername/ews/exchange.asmx"  
        #$service.Url = $uri    
		  
        ## Optional section for Exchange Impersonation  
		  
        #$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
        if (!$service.URL) {
            throw "Error connecting to EWS"
        }
        else {		
            return, $service
        }
    }
}

function Load-EWSManagedAPI {
    param( 
    )  
    Begin {
        ## Load Managed API dll  
        ###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
        $EWSDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
        if (Test-Path $EWSDLL) {
            Import-Module $EWSDLL
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

        ## end code from http://poshcode.org/624 ##

    }
}

function ConvertId {    
    param (
        [Parameter(Position = 1, Mandatory = $false)] [String]$HexId,
        [Parameter(Position = 2, Mandatory = $false)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service
    )
    process {
        $aiItem = New-Object Microsoft.Exchange.WebServices.Data.AlternateId      
        $aiItem.Mailbox = $MailboxName      
        $aiItem.UniqueId = $HexId   
        $aiItem.Format = [Microsoft.Exchange.WebServices.Data.IdFormat]::HexEntryId      
        $convertedId = $service.ConvertId($aiItem, [Microsoft.Exchange.WebServices.Data.IdFormat]::EwsId) 
        return $convertedId.UniqueId
    }
}
   
function Get-TeamsIMLog {
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 1, Mandatory = $false)] [string]$AccessToken,
        [Parameter(Position = 2, Mandatory = $false)] [string]$url,
        [Parameter(Position = 3, Mandatory = $false)] [Int32]$Days = 30,
        [Parameter(Position = 4, Mandatory = $false)] [switch]$useImpersonation,
        [Parameter(Position = 5, Mandatory = $false)] [switch]$basicAuth,
        [Parameter(Position = 6, Mandatory = $false)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position = 8, Mandatory = $false)] [string]$sipAddress,        
        [Parameter(Position = 9, Mandatory = $false)] [switch]$reportbyDate,
        [Parameter(Position = 10, Mandatory = $false)] [switch]$reportFrom,
        [Parameter(Position = 11, Mandatory = $false)] [switch]$reportTo,
        [Parameter(Position = 12, Mandatory = $false)] [switch]$htmlReport,
        [Parameter(Position = 13, Mandatory = $false)] [switch]$GroupThreads


    )  
    Process {
        $LogCollection = @()
        $GraphUserCache = @{}
        $ThreadHash = @{}
        if ($basicAuth.IsPresent) {
            if (!$Credentials) {
                $Credentials = Get-Credential
            }
            $service = Connect-Exchange -MailboxName $MailboxName -url $url -basicAuth -Credentials $Credentials
        }
        else {
            $service = Connect-Exchange -MailboxName $MailboxName -url $url -AccessToken $AccessToken
        }
        $Script:SipAddress = $MailboxName
        if ($sipAddress) {
            $Script:SipAddress = $sipAddress
        }
        if ($useImpersonation.IsPresent) {
            $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)        
        }
        $service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName); 
        if ($AccessToken) {
            $Script:Token = $AccessToken
        }
        $Script:GraphToken = Invoke-RefreshAccessToken -MailboxName $MailboxName -ResourceURL graph.microsoft.com -AccessToken $Script:Token 
        
        $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root, $MailboxName)   
        $TeamChatFolderEntryId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([System.Guid]::Parse("{E49D64DA-9F3B-41AC-9684-C6E01F30CDFA}"), "TeamChatFolderEntryId", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
        $psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
        $psPropset.Add($TeamChatFolderEntryId)
        $RootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderid, $psPropset)
        $FolderIdVal = $null    

        
        if ($RootFolder.TryGetProperty($TeamChatFolderEntryId, [ref]$FolderIdVal)) {
            $TeamChatFolderId = new-object Microsoft.Exchange.WebServices.Data.FolderId((ConvertId -HexId ([System.BitConverter]::ToString($FolderIdVal).Replace("-", "")) -service $service))
            $TeamChatFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $TeamChatFolderId);
        }
       
        $QueryDate = (Get-Date).AddDays(-$Days)
        $sf1 = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, $QueryDate)
        #Define ItemView to retrive just 1000 Items    
        $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
        $FindItemPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
        $MessageType = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Guid]"{F6E4BA45-C83C-45DA-8F38-43BD3FE76D5C}", "IOpenTypedFacet.SkypeSpaces_ConversationPost_Extension#MessageType", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
        $ThreadType = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Guid]"{F6E4BA45-C83C-45DA-8F38-43BD3FE76D5C}", "IOpenTypedFacet.SkypeSpaces_ConversationPost_Extension#ThreadType", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
        $ThreadId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Guid]"{F6E4BA45-C83C-45DA-8F38-43BD3FE76D5C}", "IOpenTypedFacet.SkypeSpaces_ConversationPost_Extension#ThreadId", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
        $FromSkypeInternalId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Guid]"{F6E4BA45-C83C-45DA-8F38-43BD3FE76D5C}", "IOpenTypedFacet.SkypeSpaces_ConversationPost_Extension#FromSkypeInternalId", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
        $ToSkypeInternalIds = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Guid]"{F6E4BA45-C83C-45DA-8F38-43BD3FE76D5C}", "IOpenTypedFacet.SkypeSpaces_ConversationPost_Extension#ToSkypeInternalIds", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::StringArray)
        $ContainsExternalParticipants = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Guid]"{F6E4BA45-C83C-45DA-8F38-43BD3FE76D5C}", "IOpenTypedFacet.Com_Compliance_ExternalParticipants#ContainsExternalParticipants", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Boolean)
        $CreatedDateTime = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Guid]"{F6E4BA45-C83C-45DA-8F38-43BD3FE76D5C}", "IOpenTypedFacet.SkypeSpaces_ConversationPost_Extension#CreatedDateTime", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime)
        $PidTagInternetMessageId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(4149, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);
        $ItemPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
        $ItemPropset.Add($MessageType)
        $ItemPropset.Add($ThreadId)
        $ItemPropset.Add($ToSkypeInternalIds)
        $ItemPropset.Add($FromSkypeInternalId)
        $ItemPropset.Add($PidTagInternetMessageId)
        $ItemPropset.Add($ContainsExternalParticipants)
        $ItemPropset.Add($CreatedDateTime)
        $ItemPropset.Add($ThreadType)
        $ivItemView.PropertySet = $FindItemPropset
        $fiItems = $null    
        do { 
            $error.clear()
            try {
                $fiItems = $TeamChatFolder.FindItems($sf1, $ivItemView)                 
            }
            catch {
                write-host ("Error " + $_.Exception.Message)
                if ($_.Exception -is [Microsoft.Exchange.WebServices.Data.ServiceResponseException]) {
                    Write-Host ("EWS Error : " + ($_.Exception.ErrorCode))
                    Start-Sleep -Seconds 60 
                }	
                $fiItems = $TeamChatFolder.FindItems($sf1, $ivItemView)     
            }
            if ($fiItems.Items.Count -gt 0) {
                [Void]$TeamChatFolder.service.LoadPropertiesForItems($fiItems, $ItemPropset)  
            }
            if ($fiItems.Items.Count -gt 0) {
                Write-Host ("Processed " + $fiItems.Items.Count + " : " + $ItemClass)
            }  
                            
            foreach ($Item in $fiItems.Items) {
                $MessageidValue = $null;
                $MessageTypeValue = $null;
                $ThreadIdValue = $null;
                $ThreadTypeValue = $null
                $FromSkypeInternalIdValue = $null
                $ToSkypeInternalIdsValue = $null
                $ContainsExternalParticipantsValue = $null
                $CreatedDateTimeValue = $null
                [Void]$Item.TryGetProperty($PidTagInternetMessageId, [ref]$MessageidValue)
                [Void]$Item.TryGetProperty($MessageType, [ref]$MessageTypeValue)
                [Void]$Item.TryGetProperty($ThreadId, [ref]$ThreadIdValue)
                [Void]$Item.TryGetProperty($ContainsExternalParticipants, [ref]$ContainsExternalParticipantsValue)
                [Void]$Item.TryGetProperty($CreatedDateTime, [ref]$CreatedDateTimeValue)
                [Void]$Item.TryGetProperty($ThreadType, [ref]$ThreadTypeValue)
                if ($Item.ItemClass -eq "IPM.SkypeTeams.Message") {
                    $LogEntry = "" | Select ThreadId, MessageType,ThreadType, Sender, Recipients, RecipientCount, TimeSent, ExternalParticipants 
                    $LogEntry.ExternalParticipants = $false 
                    if ($ContainsExternalParticipantsValue) {
                        $LogEntry.ExternalParticipants = $ContainsExternalParticipantsValue
                    }                 
                    $LogEntry.MessageType = $MessageTypeValue   
                    $LogEntry.ThreadType = $ThreadTypeValue                     
                    $LogEntry.ThreadId = $ThreadIdValue
                    $LogEntry.TimeSent = $CreatedDateTimeValue.ToLocalTime()  
                    $LogEntry.RecipientCount = 0    
                    $LogEntry.Recipients = "";       
                    if ($Item.TryGetProperty($FromSkypeInternalId, [ref]$FromSkypeInternalIdValue)) {
                        $GraphUser = $null                        
                        if (!$GraphUserCache.ContainsKey($FromSkypeInternalIdValue.Split(':')[2])) {
                            $GraphUser = Get-UserFromGuid -MailboxName $MailboxName -GUID $FromSkypeInternalIdValue.Split(':')[2]
                            $GraphUserCache.Add($FromSkypeInternalIdValue.Split(':')[2], $GraphUser)
                        }
                        else {
                            $GraphUser = $GraphUserCache[$FromSkypeInternalIdValue.Split(':')[2]]
                        }
                        if ([String]::IsNullOrEmpty($GraphUser.mail)) {
                            $LogEntry.Sender = $GraphUser.userPrincipalName
                        }
                        else {
                            $LogEntry.Sender = $GraphUser.mail
                        }
                       
                    }
                    else {
                        $LogEntry.Sender = $Item.Sender.Address
                    }
                    if ($Item.TryGetProperty($ToSkypeInternalIds, [ref]$ToSkypeInternalIdsValue)) {
                        foreach ($rcp in $ToSkypeInternalIdsValue) {
                            $GraphUser = $null
                            if (!$GraphUserCache.ContainsKey($rcp.Split(':')[2])) {
                                $GraphUser = Get-UserFromGuid -MailboxName $MailboxName -GUID $rcp.Split(':')[2]
                                $GraphUserCache.Add($rcp.Split(':')[2], $GraphUser)
                            }
                            else {
                                $GraphUser = $GraphUserCache[$rcp.Split(':')[2]]
                            }
                            if ($GraphUser) {
                                if ($GraphUser.mail) {
                                    if (!$LogEntry.Recipients.Contains($GraphUser.mail)) {
                                        if ([String]::IsNullOrEmpty($GraphUser.mail)) {                                           
                                            $atn = $GraphUser.userPrincipalName + ";"  
                                        }
                                        else {
                                            $atn = $GraphUser.mail + ";"  
                                        }
                                        $LogEntry.RecipientCount++                                                                         		
                                        $LogEntry.Recipients += $atn   
                                    }
                                }                               
                            }                           
                        }
                    }
                    else {
                        foreach ($recp in $Item.ToRecipients) {
                            $atn = $recp.Address + ";"	
                            if (![String]::IsNullOrEmpty($recp.Address)) {
                                $LogEntry.Recipients += $atn
                                $LogEntry.RecipientCount++ 
                            }	
                            
                        }
                        foreach ($recp in $Item.CCRecipients) {
                            $atn = $recp.Address + ";"		
                            if (![String]::IsNullOrEmpty($recp.Address)) {
                                $LogEntry.Recipients += $atn
                                $LogEntry.RecipientCount++ 
                            }	
                        }
                    }
                    $LogCollection += $LogEntry      
                    if ($GroupThreads.IsPresent) {
                        if ($ThreadHash.ContainsKey($LogEntry.ThreadId)) {
                            $ThreadHash[$LogEntry.ThreadId].MessageCount++
                            if ($LogEntry.TimeSent -gt $ThreadHash[$LogEntry.ThreadId].LastMessage) {
                                $ThreadHash[$LogEntry.ThreadId].LastMessage = $LogEntry.TimeSent
                            }
                            if ($LogEntry.TimeSent -lt $ThreadHash[$LogEntry.ThreadId].FirstMessage) {
                                $ThreadHash[$LogEntry.ThreadId].FirstMessage = $LogEntry.TimeSent
                            }
                            $rcps = $LogEntry.Recipients.Split(';')
                            foreach ($rcp in $rcps) {
                                if (!$LogEntry.Recipients.Contains($rcp)) {
                                    $LogEntry.Recipients += $rcp + ";"
                                    $LogEntry.RecipientCount++
                                }
                            }
                            if (!$ThreadHash[$LogEntry.ThreadId].Senders.Contains($LogEntry.Sender)) {
                                $ThreadHash[$LogEntry.ThreadId].Senders += $LogEntry.Sender + ";"
                                $ThreadHash[$LogEntry.ThreadId].SenderCount++
                            }

                            
                        }
                        else {
                            Add-Member -InputObject $LogEntry -NotePropertyName MessageCount -NotePropertyValue 1
                            Add-Member -InputObject $LogEntry -NotePropertyName FirstMessage -NotePropertyValue $LogEntry.TimeSent
                            Add-Member -InputObject $LogEntry -NotePropertyName LastMessage -NotePropertyValue $LogEntry.TimeSent
                            Add-Member -InputObject $LogEntry -NotePropertyName Senders -NotePropertyValue ($LogEntry.Sender + ";")
                            Add-Member -InputObject $LogEntry -NotePropertyName SenderCount -NotePropertyValue 1 
                            $LogEntry.PSObject.Properties.Remove('Sender')
                            $LogEntry.PSObject.Properties.Remove('TimeSent')
                            $ThreadHash.Add($LogEntry.ThreadId, $LogEntry)
                        }
                    }            
               
                }
                $ivItemView.Offset += $fiItems.Items.Count    
                    
            }
        }while ($fiItems.MoreAvailable -eq $true)
        if ($GroupThreads.IsPresent) {
            return $ThreadHash.Values
           
        }
        else {
            if ($reportbyDate.IsPresent) {
                
                $rptHash = @{}
                $LogCollection | Select-Object @{n = "Date"; e = {$_.TimeSent.ToString("yyyy-MM-dd")}}, Sender, Recipients, RecipientCount, TimeSent, ExternalParticipants | Sort-Object Date -Descending | foreach-object {
                    if ($rptHash.ContainsKey($_.Date)) {
                        $rptObject = $rptHash[$_.Date]
                    }
                    else {
                        $rptObject = "" | Select Date, Sent, Received, PeopleMessaged, TotalCount
                        $rptObject.PeopleMessaged = @{}
                        $rptObject.Date = $_.Date
                        $rptObject.Sent = 0;
                        $rptObject.Received = 0
                        $rptObject.TotalCount = 0
                        $rptHash.Add($_.Date, $rptObject)
                    }
                    if ($_.Sender.Contains($Script:SipAddress)) {
                        $rptObject.Sent++
                        $rptObject.TotalCount++
                        $rcps = $_.Recipients.Split(';')
                        foreach ($rcp in $rcps) {
                            if (![String]::IsNullOrEmpty($rcp)) {
                                if (!$rptObject.PeopleMessaged.ContainsKey($rcp)) {
                                    if (!$rcp.Contains($Script:SipAddress)) {
                                        $rptObject.PeopleMessaged.Add($rcp, "");
                                    }                                
                                }
                            }
                        }
                    }
                    else {
                        $rcps = $_.Recipients.Split(';')
                        foreach ($rcp in $rcps) {
                            if (![String]::IsNullOrEmpty($rcp)) {
                                if (!$rptObject.PeopleMessaged.ContainsKey($rcp)) {
                                    if (!$_.Sender.Contains($Script:SipAddress)) {
                                        $rptObject.PeopleMessaged.Add($rcp, "");
                                    }                                    
                                }
                            }
                        }
                        $rptObject.Received++
                        $rptObject.TotalCount++                        
                    }
                }
                $OutputObject = $rptHash.Values | Sort-Object Date -Descending | Select-Object Date, Sent, Received, @{n = "PeopleMessagedCount"; e = {$_.PeopleMessaged.Count}}, @{n = "PeopleMessaged"; e = {[String]::Join(";", ($_.PeopleMessaged.Keys | % ToString))}}, TotalCount
                
            }
            else {
                if ($reportFrom.IsPresent) {
                    $rptHash = @{}
                    $LogCollection | ForEach-Object {
                        if ($rptHash.ContainsKey($_.Sender)) {
                            $rptObject = $rptHash[$_.Sender]
                        }
                        else {
                            $rptObject = "" | Select Sender, ReceivedFrom
                            $rptObject.Sender = $_.Sender                   
                            $rptHash.Add($_.Sender, $rptObject)
                        }
                        $rptObject.ReceivedFrom += 1
                    }
                    $OutputObject = $rptHash.Values | Sort-Object ReceivedFrom -Descending 
                }
                else {
                    if ($reportTo.IsPresent) {
                        $rptHash = @{}
                        $LogCollection | ForEach-Object {
                            $rcps = $_.Recipients.Split(';')
                            foreach ($rcp in $rcps) {
                                if (![String]::IsNullOrEmpty($rcp)) {
                                    if (!$rptHash.ContainsKey($rcp)) {
                                        if (!$rcp.Contains($Script:SipAddress)) {
                                            $rptObject = "" | Select Recipient, SentTo
                                            $rptObject.Recipient = $rcp
                                            $rptObject.SentTo = 1
                                            $rptHash.Add($rcp, $rptObject);
                                        }   
                                    }
                                    else {
                                        $rptHash[$rcp].SentTo += 1;
                                    }
                                }
                            }
                        }
                        $OutputObject = $rptHash.Values | Sort-Object SentTo -Descending 
                    }
                    else {
                        $OutputObject = $LogCollection
                    }
               
                }
              
            }
            $tableStyle = @"
<style>
BODY{background-color:white;}
TABLE{border-width: 1px;
  border-style: solid;
  border-color: black;
  border-collapse: collapse;
}
TH{border-width: 1px;
  padding: 10px;
  border-style: solid;
  border-color: black;
  background-color:#66CCCC
}
TD{border-width: 1px;
  padding: 2px;
  border-style: solid;
  border-color: black;
  background-color:white
}
</style>
"@
  
            $body = @"
<p style="font-size:25px;family:calibri;color:#ff9100">
$TableHeader
</p>
"@
            if ($htmlReport.IsPresent) {
                Write-Output ($OutputObject | ConvertTo-Html -Body $body -head $tableStyle)
            }
            else {
                Write-Output $OutputObject
            }            
           
        }
        
    
        
    }
}

function Get-UserFromGuid {
    param(
        [Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 2, Mandatory = $false)] [string]$GUID

    )
    Process {
 
        $HttpClient = Get-HTTPClient -MailboxName $MailboxName
        $HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", (ConvertFrom-SecureStringCustom -SecureToken $Script:GraphToken.access_token));
        $URL = "https://graph.microsoft.com/v1.0/users('" + $GUID + "')"
        $ClientResult = $HttpClient.GetAsync([Uri]$URL)
        return ConvertFrom-Json  $ClientResult.Result.Content.ReadAsStringAsync().Result
       
    }
}

function Invoke-ValidateToken {
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$service
    
    )
    begin {
        $MailboxName = $Script:Token.mailbox
        $minTime = new-object DateTime(1970, 1, 1, 0, 0, 0, 0, [System.DateTimeKind]::Utc);
        $expiry = $minTime.AddSeconds($Script:Token.expires_on)
        if ($expiry -le [DateTime]::Now.ToUniversalTime().AddMinutes(10)) {
            if ([bool]($Script:Token.PSobject.Properties.name -match "refresh_token")) {
                write-host "Refresh Token"
                $Script:Token = Invoke-RefreshAccessToken -MailboxName $MailboxName -AccessToken $Script:Token
                $OAuthCredentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials((ConvertFrom-SecureStringCustom -SecureToken $Script:Token.access_token))
                $service.Credentials = $OAuthCredentials
            }
            else {
                throw "App Token has expired"
            }        
        }
    }
}
function ConvertToString($ipInputString) {  
    $Val1Text = ""  
    for ($clInt = 0; $clInt -lt $ipInputString.length; $clInt++) {  
        $Val1Text = $Val1Text + [Convert]::ToString([Convert]::ToChar([Convert]::ToInt32($ipInputString.Substring($clInt, 2), 16)))  
        $clInt++  
    }  
    return $Val1Text  
} 


function Get-EWSAccessToken {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string]
        $MailboxName,
    
        [Parameter(Position = 1, Mandatory = $false)]
        [string]
        $ClientId,
    
        [Parameter(Position = 2, Mandatory = $false)]
        [string]
        $redirectUrl,
    
        [Parameter(Position = 3, Mandatory = $false)]
        [string]
        $ClientSecret,
    
        [Parameter(Position = 4, Mandatory = $false)]
        [string]
        $ResourceURL,
    
        [Parameter(Position = 5, Mandatory = $false)]
        [switch]
        $Beta,
    
        [Parameter(Position = 6, Mandatory = $false)]
        [String]
        $Prompt,

        [Parameter(Position = 7, Mandatory = $false)]
        [switch]
        $CacheCredentials
    
    )
    Begin {
        Add-Type -AssemblyName System.Web
        $HttpClient = Get-HTTPClient -MailboxName $MailboxName
        if ([String]::IsNullOrEmpty($ClientId)) {
            $ReturnToken = Get-ProfiledToken -MailboxName $MailboxName
            if ($ReturnToken -eq $null) {
                Write-Error ("No Access Token for " + $MailboxName)
            }
            else {				
                return $ReturnToken
            }
        }
        else {
            if ([String]::IsNullOrEmpty(($ClientSecret))) {
                $ClientSecret = $AppSetting.ClientSecret
            }
            if ([String]::IsNullOrEmpty($redirectUrl)) {
                $redirectUrl = [System.Web.HttpUtility]::UrlEncode("urn:ietf:wg:oauth:2.0:oob")
            }
            else {
                $redirectUrl = [System.Web.HttpUtility]::UrlEncode($redirectUrl)
            }
            if ([String]::IsNullOrEmpty($ResourceURL)) {
                $ResourceURL = $AppSetting.ResourceURL
            }
            if ([String]::IsNullOrEmpty($Prompt)) {
                $Prompt = "refresh_session"
            }
        
            $Phase1auth = Show-OAuthWindow -Url "https://login.microsoftonline.com/common/oauth2/authorize?resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&response_type=code&redirect_uri=$redirectUrl&prompt=$Prompt&domain_hint=organizations"
            $code = $Phase1auth["code"]
            $AuthorizationPostRequest = "resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&grant_type=authorization_code&code=$code&redirect_uri=$redirectUrl"
            if (![String]::IsNullOrEmpty($ClientSecret)) {
                $AuthorizationPostRequest = "resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&client_secret=$ClientSecret&grant_type=authorization_code&code=$code&redirect_uri=$redirectUrl"
            }
            $content = New-Object System.Net.Http.StringContent($AuthorizationPostRequest, [System.Text.Encoding]::UTF8, "application/x-www-form-urlencoded")
            $ClientReesult = $HttpClient.PostAsync([Uri]("https://login.windows.net/common/oauth2/token"), $content)
            $JsonObject = ConvertFrom-Json -InputObject $ClientReesult.Result.Content.ReadAsStringAsync().Result
            if ([bool]($JsonObject.PSobject.Properties.name -match "refresh_token")) {
                $JsonObject.refresh_token = (Get-ProtectedToken -PlainToken $JsonObject.refresh_token)
            }
            if ([bool]($JsonObject.PSobject.Properties.name -match "access_token")) {
                $JsonObject.access_token = (Get-ProtectedToken -PlainToken $JsonObject.access_token)
            }
            if ([bool]($JsonObject.PSobject.Properties.name -match "id_token")) {
                $JsonObject.id_token = (Get-ProtectedToken -PlainToken $JsonObject.id_token)
            }
            Add-Member -InputObject $JsonObject -NotePropertyName clientid -NotePropertyValue $ClientId
            Add-Member -InputObject $JsonObject -NotePropertyName redirectUrl -NotePropertyValue $redirectUrl
            Add-Member -InputObject $JsonObject -NotePropertyName mailbox -NotePropertyValue $MailboxName
            if ($Beta.IsPresent) {
                Add-Member -InputObject $JsonObject -NotePropertyName Beta -NotePropertyValue $True
            }
            return $JsonObject
        }
    }
}
function Show-OAuthWindow {
    [CmdletBinding()]
    param (
        [System.Uri]
        $Url
    
    )
    ## Start Code Attribution
    ## Show-AuthWindow function is the work of the following Authors and should remain with the function if copied into other scripts
    ## https://foxdeploy.com/2015/11/02/using-powershell-and-oauth/
    ## https://blogs.technet.microsoft.com/ronba/2016/05/09/using-powershell-and-the-office-365-rest-api-with-oauth/
    ## End Code Attribution
    Add-Type -AssemblyName System.Web
    Add-Type -AssemblyName System.Windows.Forms

    $form = New-Object -TypeName System.Windows.Forms.Form -Property @{ Width = 440; Height = 640 }
    $web = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{ Width = 420; Height = 600; Url = ($url) }
    $DocComp = {
        $Global:uri = $web.Url.AbsoluteUri
        if ($Global:Uri -match "error=[^&]*|code=[^&]*") { $form.Close() }
    }
    $web.ScriptErrorsSuppressed = $true
    $web.Add_DocumentCompleted($DocComp)
    $form.Controls.Add($web)
    $form.Add_Shown( { $form.Activate() })
    $form.ShowDialog() | Out-Null
    $queryOutput = [System.Web.HttpUtility]::ParseQueryString($web.Url.Query)
    $output = @{ }
    foreach ($key in $queryOutput.Keys) {
        $output["$key"] = $queryOutput[$key]
    }
    return $output
}

function Get-ProtectedToken {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [String]
        $PlainToken
    )
    begin {
        $SecureEncryptedToken = Protect-String -String $PlainToken
        return, $SecureEncryptedToken
    }
}
function Get-HTTPClient {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string]
        $MailboxName
    )
    process {
        Add-Type -AssemblyName System.Net.Http
        $handler = New-Object  System.Net.Http.HttpClientHandler
        $handler.CookieContainer = New-Object System.Net.CookieContainer
        $handler.AllowAutoRedirect = $true;
        $HttpClient = New-Object System.Net.Http.HttpClient($handler);
        #$HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", "");
        $Header = New-Object System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json")
        $HttpClient.DefaultRequestHeaders.Accept.Add($Header);
        $HttpClient.Timeout = New-Object System.TimeSpan(0, 0, 90);
        $HttpClient.DefaultRequestHeaders.TransferEncodingChunked = $false
        if (!$HttpClient.DefaultRequestHeaders.Contains("X-AnchorMailbox")) {
            $HttpClient.DefaultRequestHeaders.Add("X-AnchorMailbox", $MailboxName);
        }
        $Header = New-Object System.Net.Http.Headers.ProductInfoHeaderValue("RestClient", "1.1")
        $HttpClient.DefaultRequestHeaders.UserAgent.Add($Header);
        return $HttpClient
    }
}

function Protect-String {
    <#
.SYNOPSIS
    Uses DPAPI to encrypt strings.

.DESCRIPTION
    Uses DPAPI to encrypt strings.

.PARAMETER String
    The string to encrypt.

.EXAMPLE
    PS C:\> Protect-String -String $secret

    Encrypts the content stored in $secret and returns it.
#>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingConvertToSecureStringWithPlainText", "")]
    [CmdletBinding()]
    Param (
        [Parameter(ValueFromPipeline = $true)]
        [string[]]
        $String
    )

    begin {
        Add-Type -AssemblyName System.Security -ErrorAction Stop
    }
    process {
        foreach ($item in $String) {
            $stringBytes = [Text.Encoding]::UTF8.GetBytes($item)
            $encodedBytes = [System.Security.Cryptography.ProtectedData]::Protect($stringBytes, $null, 'CurrentUser')
            [System.Convert]::ToBase64String($encodedBytes) | ConvertTo-SecureString -AsPlainText -Force
        }
    }
}

function Invoke-RefreshAccessToken {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string]
        $MailboxName,
    
        [Parameter(Position = 1, Mandatory = $true)]
        [psobject]
        $AccessToken,

        [Parameter(Position = 2, Mandatory = $true)]
        [string]
        $ResourceURL
    )
    process {
        Add-Type -AssemblyName System.Web
        $HttpClient = Get-HTTPClient -MailboxName $MailboxName
        $ClientId = $AccessToken.clientid
        # $redirectUrl = [System.Web.HttpUtility]::UrlEncode($AccessToken.redirectUrl)
        $redirectUrl = $AccessToken.redirectUrl
        $RefreshToken = (ConvertFrom-SecureStringCustom -SecureToken $AccessToken.refresh_token)
        $AuthorizationPostRequest = "client_id=$ClientId&refresh_token=$RefreshToken&grant_type=refresh_token&redirect_uri=$redirectUrl"
        if ($ResourceURL) {
            $AuthorizationPostRequest = "resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&refresh_token=$RefreshToken&grant_type=refresh_token&redirect_uri=$redirectUrl"            
        }
        $content = New-Object System.Net.Http.StringContent($AuthorizationPostRequest, [System.Text.Encoding]::UTF8, "application/x-www-form-urlencoded")
        $ClientResult = $HttpClient.PostAsync([Uri]("https://login.windows.net/common/oauth2/token"), $content)
        if (!$ClientResult.Result.IsSuccessStatusCode) {
            Write-Output ("Error making REST POST " + $ClientResult.Result.StatusCode + " : " + $ClientResult.Result.ReasonPhrase)
            Write-Output $ClientResult.Result
            if ($ClientResult.Content -ne $null) {
                Write-Output ($ClientResult.Content.ReadAsStringAsync().Result);
            }
        }
        else {
            $JsonObject = ConvertFrom-Json -InputObject $ClientResult.Result.Content.ReadAsStringAsync().Result
            if ([bool]($JsonObject.PSobject.Properties.name -match "refresh_token")) {
                $JsonObject.refresh_token = (Get-ProtectedToken -PlainToken $JsonObject.refresh_token)
            }
            if ([bool]($JsonObject.PSobject.Properties.name -match "access_token")) {
                $JsonObject.access_token = (Get-ProtectedToken -PlainToken $JsonObject.access_token)
            }
            if ([bool]($JsonObject.PSobject.Properties.name -match "id_token")) {
                $JsonObject.id_token = (Get-ProtectedToken -PlainToken $JsonObject.id_token)
            }
            Add-Member -InputObject $JsonObject -NotePropertyName clientid -NotePropertyValue $ClientId
            Add-Member -InputObject $JsonObject -NotePropertyName redirectUrl -NotePropertyValue $redirectUrl
            Add-Member -InputObject $JsonObject -NotePropertyName mailbox -NotePropertyValue $MailboxName
            if ($AccessToken.Beta) {
                Add-Member -InputObject $JsonObject -NotePropertyName Beta -NotePropertyValue True
            }
        }
        return $JsonObject		
    }
}

function ConvertFrom-SecureStringCustom {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [System.Security.SecureString]
        $SecureToken
    )
    process {
        #$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureToken)
        $Token = Unprotect-String -String $SecureToken
        return, $Token
    }
}

function Unprotect-String {
    <#
.SYNOPSIS
    Uses DPAPI to decrypt strings.

.DESCRIPTION
    Uses DPAPI to decrypt strings.
    Designed to reverse encryption applied by Protect-String

.PARAMETER String
    The string to decrypt.

.EXAMPLE
    PS C:\> Unprotect-String -String $secret

    Decrypts the content stored in $secret and returns it.
#>
    [CmdletBinding()]
    Param (
        [Parameter(ValueFromPipeline = $true)]
        [System.Security.SecureString[]]
        $String
    )

    begin {
        Add-Type -AssemblyName System.Security -ErrorAction Stop
    }
    process {
        foreach ($item in $String) {
            $cred = New-Object PSCredential("irrelevant", $item)
            $stringBytes = [System.Convert]::FromBase64String($cred.GetNetworkCredential().Password)
            $decodedBytes = [System.Security.Cryptography.ProtectedData]::Unprotect($stringBytes, $null, 'CurrentUser')
            [Text.Encoding]::UTF8.GetString($decodedBytes)
        }
    }
}
$Script:Token = $null
$Script:GraphToken = $null
$Script:MaxCount = 0
$Script:UseMaxCount = $false
$Script:MaxCountExceeded = $false