function Connect-Exchange { 
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 1, Mandatory = $true)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position = 2, Mandatory = $false)] [string]$url
    )  
    Begin {
        Load-EWSManagedAPI
		
        ## Set Exchange Version  
        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
		  
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

        ## end code from http://poshcode.org/624

    }
}


function Send-EWSMessage {
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Mandatory = $true)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position = 2, Mandatory = $false)] [switch]$useImpersonation,
        [Parameter(Position = 3, Mandatory = $false)] [string]$url,
        [Parameter(Position = 6, Mandatory = $true)] [String]$To,
        [Parameter(Position = 7, Mandatory = $true)] [String]$Subject,
        [Parameter(Position = 8, Mandatory = $true)] [String]$Body,
        [Parameter(Position = 9, Mandatory = $false)] [String]$Attachment
    )  
    Begin {
        $service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials -url $url
        if ($useImpersonation.IsPresent) {
            $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)        
        }
        $service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName);  
        $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems, $MailboxName)   
        $SentItems = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderid)
        $EmailMessage = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $service  
        $EmailMessage.Subject = $Subject
        #Add Recipients    
        $EmailMessage.ToRecipients.Add($To)  
        $EmailMessage.Body = New-Object Microsoft.Exchange.WebServices.Data.MessageBody  
        $EmailMessage.Body.BodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::HTML  
        $EmailMessage.Body.Text = $Body
        $EmailMessage.From = $MailboxName
        if ($Attachment) {   
            $EmailMessage.Attachments.AddFileAttachment($Attachment)
        }
        $EmailMessage.SendAndSaveCopy($SentItems.Id) 
  
    }
}

function Invoke-ParseEmailBodyLinks {
    [CmdletBinding()] 
    param (
        [Parameter(Position = 1, Mandatory = $false)]
        [psobject]$Item

    )
	
    process {		
        $ParsedLinksObject = "" | Select HasBaseURL, ParsedBaseURL, Links, Images
        $ParsedLinksObject.HasBaseURL = $false
        $ParsedLinksObject.Links = @()      
        $ParsedLinksObject.Images = @()  
        $RegExHtmlLinks = "<`(.*?)>"  
        $PR_HTML = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x1013, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
        $HtmlBytes = $null
        if ($Item.TryGetProperty($PR_HTML, [ref]$HtmlBytes)) {          
            if ($HtmlBytes -ne $null) {                
                $matchedItems = [regex]::matches([System.Text.Encoding]::UTF8.GetString($HtmlBytes), $RegExHtmlLinks, [system.Text.RegularExpressions.RegexOptions]::Singleline)     
                foreach ($Match in $matchedItems) {   
                    if (!$Match.Value.StartsWith("</")) {
                        try {
                            if ($Match.Value.StartsWith("<base ", [System.StringComparison]::InvariantCultureIgnoreCase)) {
                                $ParsedLinksObject.HasBaseURL = $true
                                $Attributes = $Match.Value.Split(" ")
                                foreach ($Attribute in $Attributes) {
                                    if ($Attribute.Length -gt 10) {
                                        if ($Attribute.StartsWith('href=', [System.StringComparison]::InvariantCultureIgnoreCase)) {                                        
                                            $ParsedLinksObject.ParsedBaseURL = ([URI]($Attribute.Substring(6, $Attribute.Length - 7).Replace("`"", "").Replace("'", "").Replace("`r`n", "")))   
                                        }   
                                    }                                
                                }                                  
                            }
                            if ($Match.Value.StartsWith("<a "), [System.StringComparison]::InvariantCultureIgnoreCase) {
                                $Attributes = $Match.Value.Split(" ")
                                foreach ($Attribute in $Attributes) {
                                    if ($Attribute.StartsWith('href=', [System.StringComparison]::InvariantCultureIgnoreCase)) {     
                                        if ($Attribute.Length -gt 10) {
                                            $hrefVal = ([URI]($Attribute.Substring(6, $Attribute.Length - 7).Replace("`"", "").Replace("'", "").Replace("`r`n", "`n")))
                                            if ($ParsedLinksObject.HasBaseURL) {
                                                if ([String]::IsNullOrEmpty($hrefVal.DnsSafeHost)) {
                                                    $newHost = $ParsedLinksObject.ParsedBaseURL.OriginalString + $hrefVal.OriginalString
                                                    $hrefVal = ([URI]($newHost))
                                                }
                                            }
                                            $ParsedLinksObject.Links += $hrefVal   
                                        }                                   
                                
                                    }                                   
                                }                                
                            }
                            if ($Match.Value.StartsWith("<img ", [System.StringComparison]::InvariantCultureIgnoreCase)) {
                                $Attributes = $Match.Value.Split(" ")
                                foreach ($Attribute in $Attributes) {
                                    if ($Attribute.Length -gt 7) {
                                        if ($Attribute.StartsWith('src=', [System.StringComparison]::InvariantCultureIgnoreCase)) {                                        
                                            $ParsedLinksObject.Images += ([URI]($Attribute.Substring(5, $Attribute.Length - 6).Replace("`"", "").Replace("'", "").Replace("`r`n", "`n")))   
                                        } 
                                    }                                 
                        
                                }
                            } 
                        }
                        catch {
                            Write-host ("Parse exception " + $_.Exception.Message + " on Message " + $Item.Subject)
                            $Error.Clear()
                        }                       
                    }  
                }          

            }         

        }
        $Item | Add-Member -Name "ParsedLinks" -Value $ParsedLinksObject -MemberType NoteProperty -Force                 
        
      
    }

}
function Get-LinkReport {
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 1, Mandatory = $true)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position = 2, Mandatory = $false)] [string]$url,
        [Parameter(Position = 3, Mandatory = $false)] [switch]$useImpersonation,
        [Parameter(Position = 4, Mandatory = $true)] [string]$FolderPath, 
        [Parameter(Position = 7, Mandatory = $true)] [Int]$MessageCount 
    )  
    Process {
        $Messages = Get-EWSBodyLinks -MailboxName $MailboxName -Credentials $Credentials -url $url -useImpersonation:$useImpersonation.IsPresent -FolderPath $FolderPath -MessageCount $MessageCount
        $HrefPaths = New-Object 'system.collections.generic.dictionary[string,PsObject]'
        $Domains = New-Object 'system.collections.generic.dictionary[string,PsObject]'
        $Hrefs = New-Object 'system.collections.generic.dictionary[string,PsObject]'
        $BaseHrefs = New-Object 'system.collections.generic.dictionary[string,PsObject]'
        $Images = New-Object 'system.collections.generic.dictionary[string,PsObject]'
        $ImageDomains = New-Object 'system.collections.generic.dictionary[string,PsObject]'
        foreach ($Message in $Messages) {
            $MessageDomains = @{}      
            $MessageDomainsHrefPaths = @{}      
            $MessageDomainsHrefs = @{}  
            $MessageDomainsImages = @{}    
            $MessageImageDomains = @{} 
            if ($Message.ParsedLinks.HasBaseURL) {
                if (!$BaseHrefs.ContainsKey($Message.ParsedLinks.ParsedBaseURL)) {
                    $values = "" | Select BaseHref, Count
                    $values.BaseHref = $Message.ParsedLinks.ParsedBaseURL
                    $values.Count = 1
                    $BaseHrefs.add($Message.ParsedLinks.ParsedBaseURL, $values)
                }
                else {
                    $BaseHrefs[$Message.ParsedLinks.ParsedBaseURL].Count++
                }
            }
            foreach ($link in $Message.ParsedLinks.Links) {
                if (![String]::IsNullOrEmpty($link.DnsSafeHost)) {
                    if (![String]::IsNullOrEmpty($link.AbsolutePath)) {
                        $fpath = $link.host + "/" + $link.AbsolutePath
                        if (!$HrefPaths.ContainsKey($fpath)) {
                            $Counts = "" | select HrefPath, MessageCount, LinkCount
                            $Counts.HrefPath = $fpath
                            $Counts.MessageCount = 0
                            $Counts.LinkCount = 1
                            $HrefPaths.Add($fpath, $Counts)
                        }
                        else {
                            $HrefPaths[$fpath].LinkCount++
                        }
                        if (!$MessageDomainsHrefPaths.Contains($fpath)) {
                            $MessageDomainsHrefPaths.Add($fpath, 1)
                            $HrefPaths[$fpath].MessageCount++
                        }
                    }
                    if (![String]::IsNullOrEmpty($link.AbsoluteUri)) {
                        if (!$Hrefs.ContainsKey($link.AbsoluteUri)) {
                            $Counts = "" | select Href, MessageCount, LinkCount
                            $Counts.Href = $link.AbsoluteUri
                            $Counts.MessageCount = 0
                            $Counts.LinkCount = 1
                            $Hrefs.Add($link.AbsoluteUri, $Counts)
                        }
                        else {
                            $Hrefs[$link.AbsoluteUri].LinkCount++
                        }
                        if (!$MessageDomainsHrefs.Contains($link.AbsoluteUri)) {
                            $MessageDomainsHrefs.Add($link.AbsoluteUri, 1)
                            $Hrefs[$link.AbsoluteUri].MessageCount++
                        }
                    }
                    if (!$Domains.ContainsKey($link.DnsSafeHost)) {
                        $Counts = "" | select HostName, MessageCount, LinkCount
                        $Counts.MessageCount = 0
                        $Counts.LinkCount = 1
                        $Counts.HostName = $link.DnsSafeHost
                        $Domains.Add($link.DnsSafeHost, $Counts)
                    }
                    else {
                        $Domains[$link.DnsSafeHost].LinkCount++
                    }
                    if (!$MessageDomains.Contains($link.DnsSafeHost)) {
                        $MessageDomains.Add($link.DnsSafeHost, 1)
                        $Domains[$link.DnsSafeHost].MessageCount++
                    }
                }
            }
            foreach ($link in $Message.ParsedLinks.Images) {
                if (![String]::IsNullOrEmpty($link.AbsoluteUri)) {
                    if (!$Images.ContainsKey($link.AbsoluteUri)) {
                        $Counts = "" | select Src, MessageCount, LinkCount
                        $Counts.Src = $link.AbsoluteUri
                        $Counts.MessageCount = 0
                        $Counts.LinkCount = 1
                        $Images.Add($link.AbsoluteUri, $Counts)
                    }
                    else {
                        $Images[$link.AbsoluteUri].LinkCount++
                    }
                    if (!$MessageDomainsImages.Contains($link.AbsoluteUri)) {
                        $MessageDomainsImages.Add($link.AbsoluteUri, 1)
                        $Images[$link.AbsoluteUri].MessageCount++
                    }
                    if (![String]::IsNullOrEmpty($link.DnsSafeHost)) {
                        if (!$ImageDomains.ContainsKey($link.DnsSafeHost)) {
                            $Counts = "" | select HostName, MessageCount, LinkCount
                            $Counts.MessageCount = 0
                            $Counts.LinkCount = 1
                            $Counts.HostName = $link.DnsSafeHost
                            $ImageDomains.Add($link.DnsSafeHost, $Counts)
                        }
                        else {
                            $ImageDomains[$link.DnsSafeHost].LinkCount++
                        }
                        if (!$MessageImageDomains.Contains($link.DnsSafeHost)) {
                            $MessageImageDomains.Add($link.DnsSafeHost, 1)
                            $ImageDomains[$link.DnsSafeHost].MessageCount++
                        }
                    }
                }   
            }
        }
        $report = "" | Select Domains, hrefPaths, Hrefs, BaseHrefs, Images, ImageDomains
        $report.Domains = [Collections.Generic.List[PsObject]]$Domains.Values
        $report.HrefPaths = [Collections.Generic.List[PsObject]]$HrefPaths.Values
        $report.Hrefs = [Collections.Generic.List[PsObject]]$Hrefs.Values       
        $report.BaseHrefs = [Collections.Generic.List[PsObject]]$BaseHrefs.Values 
        $report.Images = [Collections.Generic.List[PsObject]]$Images.Values
        $report.ImageDomains = [Collections.Generic.List[PsObject]]$ImageDomains.Values
        return, $report
    }
}

function Get-EWSBodyLinks {
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 1, Mandatory = $true)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position = 2, Mandatory = $false)] [string]$url,
        [Parameter(Position = 3, Mandatory = $false)] [switch]$useImpersonation,
        [Parameter(Position = 4, Mandatory = $true)] [string]$FolderPath, 
        [Parameter(Position = 5, Mandatory = $false)] [Int]$MessageCount 
    )  
    Process {
        $MaxCount = $MessageCount
        $folders = Invoke-GenericFolderConnect -MailboxName $MailboxName -Credentials $Credentials -url $url -useImpersonation:$useImpersonation.IsPresent -FolderPath $FolderPath -Recurse:$false
        $PR_ENTRYID = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0FFF, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)  
        foreach ($Folder in $folders) {
            if ($MaxCount -gt 0) {
                $Script:MaxCount = $MaxCount
                $Script:UseMaxCount = $true
                $Script:MaxCountExceeded = $false
            }
            else {
                $Script:MaxCount = 0
                $Script:UseMaxCount = $false
                $Script:MaxCountExceeded = $false
            }
            if ($Script:UseMaxCount) {
                if ($Script:MaxCount -gt 50) {
                    $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(50)
                }
                else {
                    $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView($Script:MaxCount)
                }
            }
            else {
                $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(50)
            }                
            $ItemPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
            $ItemPropset.Add($PR_ENTRYID)
            $ItemPropset.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Preview)
            $itemCount = 0            
            $fiItems = $null    
            do { 
                $error.clear()
                try {
                    $fiItems = $Folder.service.FindItems($Folder.Id, $ivItemView) 
                }
                catch {
                    write-host ("Error " + $_.Exception.Message)
                    if ($_.Exception -is [Microsoft.Exchange.WebServices.Data.ServiceResponseException]) {
                        Write-Host ("EWS Error : " + ($_.Exception.ErrorCode))
                        Start-Sleep -Seconds 60 
                    }	
                    $fiItems = $Folder.service.FindItems($Folder.Id, $ivItemView) 
                }
                if ($fiItems.Items.Count -gt 0) {
                    $BodyPropSet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
                    $BodyPropSet.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::InternetMessageHeaders)
                    $BodyPropSet.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject)
                    $BodyPropSet.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Sender)
                    $BodyPropSet.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::From)
                    $BodyPropSet.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::DateTimeReceived)
                    $BodyPropSet.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::WebClientReadFormQueryString)
                    $BodyPropSet.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::HasAttachments)
                    $PR_HTML = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x1013, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
                    $BodyPropSet.Add($PR_HTML)
                    [Void]$Folder.service.LoadPropertiesForItems($fiItems, $BodyPropSet)  
                }
                if ($fiItems.Items.Count -gt 0) {
                    Write-Host ("Processed " + $fiItems.Items.Count + " : " + $ItemClass + " Total : " + $itemCount)
                }
                    
                foreach ($Item in $fiItems.Items) { 
                    $itemCount ++
                    $Okay = $true;
                    if ($Script:UseMaxCount) {
                        if ($itemCount -gt $Script:MaxCount) {
                            $okay = $False                               
                        }
                    }
                    if ($Okay) {
                        $Item | Add-Member -Name "FolderPath" -Value $Folder.FolderPath -MemberType NoteProperty
                        $Item | Add-Member -Name "SenderEmailAddress" -Value $Item.From.Address -MemberType NoteProperty
                        Invoke-ParseEmailBodyLinks -Item $Item
                        Write-Output $Item
                    }
                         
                }   
                if ($Script:UseMaxCount) {
                    if ($itemCount -ge $Script:MaxCount) {
                        $Script:MaxCountExceeded = $true
                    } 
                }
                $ivItemView.Offset += $fiItems.Items.Count    
            }while ($fiItems.MoreAvailable -eq $true -band (!$Script:MaxCountExceeded)) 
        }
    }
}
function Invoke-GenericFolderConnect {
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 1, Mandatory = $true)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position = 2, Mandatory = $false)] [string]$url,
        [Parameter(Position = 3, Mandatory = $false)] [switch]$useImpersonation,
        [Parameter(Position = 4, Mandatory = $false)] [string]$FolderPath,
        [Parameter(Position = 5, Mandatory = $false)] [switch]$Recurse,
        [Parameter(Position = 6, Mandatory = $false)] [switch]$RootFolder
    )  
    Process {
        $service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials -url $url
        if ($useImpersonation.IsPresent) {
            $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)        
        }
        $service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName);  
       
        $folders = Get-FolderFromPath -MailboxName $MailboxName -FolderPath $FolderPath -service $service -Recurse:$Recurse.IsPresent -RootFolder:$RootFolder.IsPresent
        return $folders 
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
function Get-FolderFromPath {
    param (
        [Parameter(Position = 0, Mandatory = $false)] [string]$FolderPath,
        [Parameter(Position = 1, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 2, Mandatory = $true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service,
        [Parameter(Position = 3, Mandatory = $false)] [Microsoft.Exchange.WebServices.Data.PropertySet]$PropertySet,
        [Parameter(Position = 5, Mandatory = $false)] [switch]$Recurse,
        [Parameter(Position = 6, Mandatory = $false)] [switch]$RootFolder
		  )
    process {
        ## Find and Bind to Folder based on Path  
        #Define the path to search should be seperated with \  
        #Bind to the MSGFolder Root  
        $SourceKey = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x65E0, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
        $psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
        $psPropset.Add($SourceKey)
        #PR_POLICY_TAG 0x3019
        $PolicyTag = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3019, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
        #PR_RETENTION_FLAGS 0x301D   
       	$RetentionFlags = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301D, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
       	#PR_RETENTION_PERIOD 0x301A
       	$RetentionPeriod = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301A, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
        $PidTagMessageSizeExtended = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0xe08, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long);
        $PR_ATTACH_ON_NORMAL_MSG_COUNT = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x66B1, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
        $psPropset.Add($PolicyTag)
        $psPropset.Add($RetentionFlags)
        $psPropset.Add($RetentionPeriod)
        $psPropset.Add($PidTagMessageSizeExtended)
        $psPropset.Add($PR_ATTACH_ON_NORMAL_MSG_COUNT)

        $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $MailboxName)   
        $tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderid)  
        if (!$RootFolder.IsPresent) {
            #Split the Search path into an array  
            $fldArray = $FolderPath.Split("\") 
            if ($fldArray.Length -lt 2) {throw "No Root Folder"}
            #Loop through the Split Array and do a Search for each level of folder 
            for ($lint = 1; $lint -lt $fldArray.Length; $lint++) { 
                #Perform search based on the displayname of each folder level 
                $fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1) 
                $fvFolderView.PropertySet = $psPropset
                $SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $fldArray[$lint]) 
                $findFolderResults = $service.FindFolders($tfTargetFolder.Id, $SfSearchFilter, $fvFolderView) 
                $tfTargetFolder = $null  
                if ($findFolderResults.TotalCount -gt 0) { 
                    foreach ($folder in $findFolderResults.Folders) { 
                        $tfTargetFolder = $folder                
                    } 
                } 
                else { 
                    Write-host ("Error Folder Not Found check path and try again") -ForegroundColor Red
                    $tfTargetFolder = $null  
                    break  
                }     
            }  
        }
        else {
            $FolderPath = "\"
        }
        if ($tfTargetFolder -ne $null) {
            $tfTargetFolder | Add-Member -Name "FolderPath" -Value $FolderPath -MemberType NoteProperty
            $tfTargetFolder | Add-Member -Name "Mailbox" -Value $MailboxName -MemberType NoteProperty
            $prop1Val = $null    
            if ($tfTargetFolder.TryGetProperty($PolicyTag, [ref] $prop1Val)) {     
                $rtnStringVal =	[System.BitConverter]::ToString($prop1Val).Replace("-", "");
                $rtnStringVal = $rtnStringVal.Substring(6, 2) + $rtnStringVal.Substring(4, 2) + $rtnStringVal.Substring(2, 2) + $rtnStringVal.Substring(0, 2) + "-" + $rtnStringVal.Substring(10, 2) + $rtnStringVal.Substring(8, 2) + "-" + $rtnStringVal.Substring(14, 2) + $rtnStringVal.Substring(12, 2) + "-" + $rtnStringVal.Substring(16, 2) + $rtnStringVal.Substring(18, 2) + "-" + $rtnStringVal.Substring(20, 12)           
                Add-Member -InputObject $tfTargetFolder -MemberType NoteProperty -Name PR_POLICY_TAG -Value $rtnStringVal      
            }
            $prop2Val = $null
            if ($tfTargetFolder.TryGetProperty($RetentionFlags, [ref] $prop2Val)) {
                Add-Member -InputObject $tfTargetFolder -MemberType NoteProperty -Name PR_RETENTION_FLAGS -Value $prop2Val
            }
            $prop3Val = $null
            if ($tfTargetFolder.TryGetProperty($RetentionPeriod, [ref] $prop3Val)) {
                Add-Member -InputObject $tfTargetFolder -MemberType NoteProperty -Name PR_RETENTION_PERIOD -Value $prop3Val
            }
            $prop4Val = $null
            if ($tfTargetFolder.TryGetProperty($PidTagMessageSizeExtended, [ref]  $prop4Val)) {
                Add-Member -InputObject $tfTargetFolder -MemberType NoteProperty -Name FolderSize -Value $prop4Val
            }
            $prop5Val = $null
            if ($tfTargetFolder.TryGetProperty($PR_ATTACH_ON_NORMAL_MSG_COUNT, [ref]  $prop5Val)) {
                Add-Member -InputObject $tfTargetFolder -MemberType NoteProperty -Name PR_ATTACH_ON_NORMAL_MSG_COUNT -Value $prop5Val
            }
            $GetLastItem = {
                $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1) 
                $fiItems = $this.FindItems($ivItemView)
                if ($fiItems.Items.Count -eq 1) {
                    $fiItems.Items[0].Load()
                    $fiItems.Items[0] | Add-Member -Name "FolderPath" -Value $this.FolderPath -MemberType NoteProperty
                    Add-Member -InputObject $this -MemberType NoteProperty -Name LastItem -Value $fiItems.Items[0] -Force
                    return $fiItems.Items[0]
                }                
            }
            Add-Member -InputObject $tfTargetFolder -MemberType ScriptMethod -Name GetLastItem -Value $GetLastItem
            if ($Recurse.IsPresent) {
                $Folders = @()
                $Folders += $tfTargetFolder
                $Folders = Get-SubFolders -ParentFolder $tfTargetFolder -Folders $Folders
                return, [PSObject]$Folders
            }
            else {
                return, [Microsoft.Exchange.WebServices.Data.Folder]$tfTargetFolder
            }
           
        }
        else {
            return, $null
        }
    }
}



$Script:MaxCount = 0
$Script:UseMaxCount = $false
$Script:MaxCountExceeded = $false