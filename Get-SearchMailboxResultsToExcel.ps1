function Connect-Exchange { 
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 1, Mandatory = $false)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position = 2, Mandatory = $false)] [string]$url,
        [Parameter(Position = 3, Mandatory = $False)]
        [switch]
        $ModernAuth,		
        [Parameter(Position = 4, Mandatory = $False)]
        [String]
        $ClientId,
        [Parameter(Position = 6, Mandatory = $False)]
        [String]
        $RedirectURL = "urn:ietf:wg:oauth:2.0:oob"
    )  
    Begin {
        Invoke-LoadEWSManagedAPI
		
        ## Set Exchange Version  
        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013
		  
        ## Create Exchange Service Object  
        $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)  
		  
        ## Set Credentials to use two options are availible Option1 to use explict credentials or Option 2 use the Default (logged On) credentials  
		  
        #Credentials Option 1 using UPN for the windows Account  
        #$psCred = Get-Credential  
        if ($ModernAuth.IsPresent) {
            Write-Verbose("Using Modern Auth")
            if ([String]::IsNullOrEmpty($ClientId)) {
                $ClientId = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
            }		
            Import-Module .\Microsoft.IdentityModel.Clients.ActiveDirectory.dll -Force
            $Context = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext("https://login.microsoftonline.com/common")
            if ($Credentials -eq $null) {
                $PromptBehavior = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters -ArgumentList SelectAccount			
                $token = ($Context.AcquireTokenAsync("https://outlook.office365.com", $ClientId , $RedirectURL, $PromptBehavior)).Result
                $service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($token.AccessToken)
            }
            else {
                $AADcredential = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserPasswordCredential" -ArgumentList  $Credentials.UserName.ToString(), $Credentials.GetNetworkCredential().password.ToString()
                $token = [Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContextIntegratedAuthExtensions]::AcquireTokenAsync($Context, "https://outlook.office365.com", $ClientId, $AADcredential).result
                $service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($token.AccessToken)
            }
        }
        else {
            Write-Verbose("Using Negotiate Auth")
            if (!$Credentials) { $Credentials = Get-Credential }
            $creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(), $Credentials.GetNetworkCredential().password.ToString())
            $service.Credentials = $creds
        } 
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

function Invoke-LoadEWSManagedAPI {
    param( 
    )  
    Begin {
        $EWSdllName = "/Microsoft.Exchange.WebServices.dll"
        if (Test-Path ($script:ModuleRoot + $EWSdllName)) {
            Import-Module ($script:ModuleRoot + $EWSdllName)
            $Script:EWSDLL = $script:ModuleRoot + $EWSdllName
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
            if ($fldArray.Length -lt 2) { throw "No Root Folder" }
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

function Get-ItemFromSubject {
    param (
        [Parameter(Position = 0, Mandatory = $true)] [Microsoft.Exchange.WebServices.Data.Folder]$Folder,
        [Parameter(Position = 1, Mandatory = $true)] [String]$Subject
    )
    process {
        $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1)  
        $SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject, $Subject) 
        $findItemResults = $Folder.FindItems($SfSearchFilter, $ivItemView)
        if ($findItemResults.Items.Count -eq 1) {
            return $findItemResults.Items[0]
        }
        else {
            throw "No Item found"
        }
    }

}

function Get-SearchMailboxResultToExcel {
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 1 , Mandatory = $false)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position = 2, Mandatory = $false)] [switch]$useImpersonation,
        [Parameter(Position = 3, Mandatory = $false)] [string]$url,
        [Parameter(Position = 4, Mandatory = $true)] [string]$SearchResultPath,
        [Parameter(Position = 5, Mandatory = $false)]
        [switch]
        $ModernAuth,		
        [Parameter(Position = 6, Mandatory = $False)]
        [String]
        $ClientId,
        [Parameter(Position = 7, Mandatory = $False)]
        [String]
        $RedirectURL = "urn:ietf:wg:oauth:2.0:oob"
        

       
    )  
    Begin {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials -url $url -ModernAuth:$ModernAuth.IsPresent -ClientId $ClientId -RedirectURL $RedirectURL
        if ($useImpersonation.IsPresent) {
            $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)        
        }
        $service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName);  
        $fldArray = $SearchResultPath.Split("\") 
        $ItemPath = $fldArray[($fldArray.Length - 1)]
        $ItemPath = $ItemPath.SubString($ItemPath.IndexOfAny('-'))
        #Loop through the Split Array and do a Search for each level of folder 
        for ($lint = 1; $lint -lt ($fldArray.Length - 1); $lint++) { 
            $FolderPath += ("\" + $fldArray[$lint])
        }
        Write-verbose ("FolderPath " + $FolderPath)
        Write-verbose ("ItemPath " + $ItemPath)
        $SearchFolder = Get-FolderFromPath -MailboxName $MailboxName -service $service -FolderPath $FolderPath       
        $SearchResultItem = Get-ItemFromSubject -Folder $SearchFolder -Subject $ItemPath
        if ($SearchResultItem.HasAttachments) {
            $SearchResultItem.Load();
            foreach ($Attachment in $SearchResultItem.Attachments) {
                $Attachment.Load()
                $downloadDirectory = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
                $fileName = ($downloadDirectory + "\" + $ItemPath.SubString(1).Replace("/", "-").Replace(":", "-") + "-" + $Attachment.Name.ToString())
                $fiFile = new-object System.IO.FileStream($fileName, [System.IO.FileMode]::Create)
                $fiFile.Write($Attachment.Content, 0, $Attachment.Content.Length)
                $fiFile.Close()
                write-host ("Downloaded Attachment : " + $FileName)
                if ($FileName.contains(".zip")) {                    
                    $Zip = [System.IO.Compression.ZipFile]::OpenRead($FileName)
                    try {
                        foreach ($file in $Zip.Entries) {
                            if ($file.Name.contains(".csv")) {
                                $ms = New-Object System.IO.MemoryStream
                                $ZipStream = $file.Open()
                                $ZipStream.CopyTo($ms);
                                $outputfile = $FileName.replace("zip", "")
                                [System.IO.File]::WriteAllBytes($outputfile, $ms.ToArray())
                                Invoke-Item $outputfile
                                
                            }
                        }
                    }
                    catch {
                        Write-Host "Error" + $_.ScriptStackTrace
                    }
                    $Zip.Dispose()
                }
            }
        }else{
            write-host ("Message has not attachment")
        }
        

    }
}