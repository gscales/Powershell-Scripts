function Connect-Exchange { 
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 1, Mandatory = $true)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position = 2, Mandatory = $false)] [string]$url
    )  
    Begin {
        Load-EWSManagedAPI
		
        ## Set Exchange Version  
        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016
		  
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
        ## Load Managed API dll  
        ###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
        $EWSDLL = $Script:SavePath + "\Microsoft.Exchange.WebServices.dll"
        if (Test-Path $EWSDLL) {
            Import-Module  $EWSDLL
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

Function Remove-InvalidFileNameChars {
    param(
        [Parameter(Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
        [String]$Name
    )

    $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
    $re = "[{0}]" -f [RegEx]::Escape($invalidChars)
    return ($Name -replace $re)
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

function Invoke-MailboxFolderAttachmentReport {
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Mandatory = $true)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position = 2, Mandatory = $false)] [switch]$useImpersonation,
        [Parameter(Position = 3, Mandatory = $false)] [string]$url,
        [Parameter(Position = 4, Mandatory = $false)] [string]$FolderPath,
        [Parameter(Position = 5, Mandatory = $false)] [switch]$RetainedItems,
        [Parameter(Position = 6, Mandatory = $false)] [switch]$IncludeHidden
       
    )  
    Begin {
        $service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials -url $url
        if ($useImpersonation.IsPresent) {
            $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)        
        }
        $service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName);  
        if ($RetainedItems.IsPresent) {
            $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::RecoverableItemsDeletions, $MailboxName)         
              
            $psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
            
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
            $ReportFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderid, $psPropset)

            $ReportFolder | Add-Member -Name "FolderPath" -Value "RecoverableItemsDeletions" -MemberType NoteProperty
            $ReportFolder | Add-Member -Name "Mailbox" -Value $MailboxName -MemberType NoteProperty
            $prop1Val = $null    
            if ($ReportFolder.TryGetProperty($PolicyTag, [ref] $prop1Val)) {     
                $rtnStringVal =	[System.BitConverter]::ToString($prop1Val).Replace("-", "");
                $rtnStringVal = $rtnStringVal.Substring(6, 2) + $rtnStringVal.Substring(4, 2) + $rtnStringVal.Substring(2, 2) + $rtnStringVal.Substring(0, 2) + "-" + $rtnStringVal.Substring(10, 2) + $rtnStringVal.Substring(8, 2) + "-" + $rtnStringVal.Substring(14, 2) + $rtnStringVal.Substring(12, 2) + "-" + $rtnStringVal.Substring(16, 2) + $rtnStringVal.Substring(18, 2) + "-" + $rtnStringVal.Substring(20, 12)           
                Add-Member -InputObject $ReportFolder -MemberType NoteProperty -Name PR_POLICY_TAG -Value $rtnStringVal      
            }
            $prop2Val = $null
            if ($ReportFolder.TryGetProperty($RetentionFlags, [ref] $prop2Val)) {
                Add-Member -InputObject $ReportFolder -MemberType NoteProperty -Name PR_RETENTION_FLAGS -Value $prop2Val
            }
            $prop3Val = $null
            if ($ReportFolder.TryGetProperty($RetentionPeriod, [ref] $prop3Val)) {
                Add-Member -InputObject $ReportFolder -MemberType NoteProperty -Name PR_RETENTION_PERIOD -Value $prop3Val
            }
            $prop4Val = $null
            if ($ReportFolder.TryGetProperty($PidTagMessageSizeExtended, [ref]  $prop4Val)) {
                Add-Member -InputObject $ReportFolder -MemberType NoteProperty -Name FolderSize -Value $prop4Val
            }
            $prop5Val = $null
            if ($ReportFolder.TryGetProperty($PR_ATTACH_ON_NORMAL_MSG_COUNT, [ref]  $prop5Val)) {
                Add-Member -InputObject $ReportFolder -MemberType NoteProperty -Name PR_ATTACH_ON_NORMAL_MSG_COUNT -Value $prop5Val
            }
            
        }
        else {
            $ReportFolder = Get-FolderFromPath -MailboxName $MailboxName -service $service -FolderPath $FolderPath       
        }
        
        $folderStats = Invoke-ProcessFolder -Folder $ReportFolder -IncludeHidden:$IncludeHidden.IsPresent -MailboxName $MailboxName 
        Invoke-WriteStatsToFile -Results $folderStats
        return $folderStats
    }
}

function Invoke-WriteStatsToFile {
    param(
        [Parameter(Position = 0, Mandatory = $true)] [PSObject]$Results
         

    )
    Process {
        $fileName = $Script:SavePath + "\" + $Results.FolderPath.Replace("\", "-") + "-" + [DateTime]::Now.ToString("yyyyMMdd-HHmmss") + ".html" 
        $Header = @"
<html>
<head>
<style>
.collapsible {
    background-color: #777;
    color: white;
    cursor: pointer;
    padding: 18px;
    width: 100%;
    border: none;
    text-align: left;
    outline: none;
    font-size: 15px;
}

.active, .collapsible:hover {
    background-color: #555;
}

.collapsible:after {
    content: '\002B';
    color: white;
    font-weight: bold;
    float: right;
    margin-left: 5px;
}

.active:after {
    content: "\2212";
}

.content {
    padding: 0 18px;
    max-height: 0;    
    overflow: hidden;
    transition: max-height 0.2s ease-out;
    background-color: #f1f1f1;
}
</style>
</head>
<button id='folderstatsButton' class="collapsible">Folder Statistics</button>
<div class="content">
<div id='FolderStats'></div>
<button class="collapsible">Item Age</button>
<div class="content">
<div id='ItemAge'></div>
</div>
<button class="collapsible">Attachment Age</button>
<div class="content">
<div id='AttachmentAge'></div>
</div>
<button class="collapsible">Attachment Extensions</button>
<div class="content">
<div id='Extensions'></div>
</div>
<button class="collapsible">Cloudy Attachment Extensions</button>
<div class="content">
<div id='CloudyExtensions'></div>
</div>
</div>
<script src="https://code.jquery.com/jquery-1.12.4.js"
            integrity="sha256-Qw82+bXyGq6MydymqBxNPYTaUXXq7c8v3CwiYwLLNXU="
           crossorigin="anonymous"></script>
<script
  src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"
  integrity="sha256-T0Vest3yCU7pafRw9r+settMBX6JkKN06dqBnpQ8d30="
  crossorigin="anonymous"></script>

<link href="https://cdnjs.cloudflare.com/ajax/libs/tabulator/3.5.3/css/tabulator.min.css" rel="stylesheet">
<script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/tabulator/3.5.3/js/tabulator.min.js"></script>
<script type="text/javascript">
var coll = document.getElementsByClassName("collapsible");
var i;

for (i = 0; i < coll.length; i++) {
  coll[i].addEventListener("click", function() {
    this.classList.toggle("active");
    var content = this.nextElementSibling;
    if (content.style.maxHeight){
      content.style.maxHeight = null;
    } else {
      content.style.maxHeight = "none";
    } 
  });
}

function bytesToSize(bytes) {
   var sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
   if (bytes == 0) return '0 Byte';
   var i = parseInt(Math.floor(Math.log(bytes) / Math.log(1024)));
   return Math.round(bytes / Math.pow(1024, i), 2) + ' ' + sizes[i];
};

"@
        $Json = ConvertTo-Json -InputObject $Results -Depth 8
        $Json = "var JSontext = " + $Json + ";`r`n"
        $Footer = @"
`$("#Extensions").tabulator({
  columns:[
    {title:"Extension", field:"Extension", sortable:true, width:200,sortable:true},
    {title:"Largest File Message Subject", field:"LargestFileMessageSubject", sortable:true},
    {title:"LargestFileMessageSender", field:"LargestFileMessageSender", sortable:true},    
    {title:"Largest File Name", field:"LargestFileName", sortable:true},
    {title:"LargestFileSize", field:"LargestFileSize", sortable:true, formatter:function(cell, formatterParams){
        var CellValue = cell.getValue();
        if(CellValue > 0){
            CellValue = bytesToSize(CellValue);
        }
        return CellValue; 
    },},    
    {title:"TotalNumber", field:"TotalNumber", sortable:true, sorter:"number"},
    {title:"TotalSize", field:"TotalSize", sortable:true, formatter:function(cell, formatterParams){
        var CellValue = cell.getValue();
        if(CellValue > 0){
            CellValue = bytesToSize(CellValue);
        }
        return CellValue; 
    },}
  ],
});
`$("#CloudyExtensions").tabulator({
  columns:[
    {title:"Extension", field:"Extension", sortable:true, width:200,sortable:true},  
    {title:"TotalNumber", field:"TotalNumber", sortable:true, sorter:"number"},
    {title:"TotalSize", field:"TotalSize", sortable:true, formatter:function(cell, formatterParams){
        var CellValue = cell.getValue();
        if(CellValue > 0){
            CellValue = bytesToSize(CellValue);
        }
        return CellValue; 
    },}
  ],
});
`$("#AttachmentAge").tabulator({
  columns:[
    {title:"Date", field:"Date", sortable:true, width:200},
    {title:"Total Number", field:"TotalNumber", sortable:true},
    {title:"Total Size", field:"TotalSize", sortable:true,formatter:function(cell, formatterParams){
        var CellValue = cell.getValue();
        if(CellValue > 0){
            CellValue = bytesToSize(CellValue);
        }
        return CellValue; 
    },}
  ],
});
`$("#ItemAge").tabulator({
  columns:[
    {title:"Date", field:"Date", sortable:true, width:200},
    {title:"Total Number", field:"TotalNumber", sortable:true},
    {title:"Total Size", field:"TotalSize", sortable:true,formatter:function(cell, formatterParams){
        var CellValue = cell.getValue();
        if(CellValue > 0){
            CellValue = bytesToSize(CellValue);
        }
        return CellValue; 
    },}
  ],
});
`$("#FolderStats").tabulator({
  columns:[
    {title:"Name", field:"Name", sortable:true, width:200},
    {title:"Value", field:"Value", sortable:true},
  ],
});
`$("#Extensions").tabulator("setData", JSontext.AttachmentExtensions);
`$("#CloudyExtensions").tabulator("setData", JSontext.CloudyAttachmentExtensions);
`$("#AttachmentAge").tabulator("setData", JSontext.AttachmentAgeStatistics);
`$("#ItemAge").tabulator("setData", JSontext.ItemAgeStatistics);
var FolderStatsJson = "[{\"Name\":  \"MailboxName\",\"Value\":  \"" + JSontext.MailboxName +  "\"},";
FolderStatsJson = FolderStatsJson + "{\"Name\":  \"FolderPath\",\"Value\":  \"" + JSontext.FolderPath.replace(String.fromCharCode(92),String.fromCharCode(92,92)) +  "\"},";
FolderStatsJson = FolderStatsJson + "{\"Name\":  \"TotalFolderItems\",\"Value\":  \"" + JSontext.TotalFolderItems +  "\"},";
FolderStatsJson = FolderStatsJson + "{\"Name\":  \"TotalFolderItemsSize\",\"Value\":  \"" + bytesToSize(JSontext.TotalFolderItemsSize) +  "\"},";
FolderStatsJson = FolderStatsJson + "{\"Name\":  \"TotalItemsNoAttach\",\"Value\":  \"" + JSontext.TotalItemsNoAttach +  "\"},";
FolderStatsJson = FolderStatsJson + "{\"Name\":  \"TotalItemsNoAttachSize\",\"Value\":  \"" + bytesToSize(JSontext.TotalItemsNoAttachSize) +  "\"},";
FolderStatsJson = FolderStatsJson + "{\"Name\":  \"TotalItemsAttach\",\"Value\":  \"" + JSontext.TotalItemsAttach +  "\"},";
FolderStatsJson = FolderStatsJson + "{\"Name\":  \"TotalItemsAttachSize\",\"Value\":  \"" + bytesToSize(JSontext.TotalItemsAttachSize) +  "\"},";
FolderStatsJson = FolderStatsJson + "{\"Name\":  \"TotalFileAttachments\",\"Value\":  \"" + JSontext.TotalFileAttachments +  "\"},";
FolderStatsJson = FolderStatsJson + "{\"Name\":  \"TotalFileAttachmentsSize\",\"Value\":  \"" + bytesToSize(JSontext.TotalFileAttachmentsSize) +  "\"},";
FolderStatsJson = FolderStatsJson + "{\"Name\":  \"TotalItemAttachments\",\"Value\":  \"" + JSontext.TotalItemAttachments +  "\"},";
FolderStatsJson = FolderStatsJson + "{\"Name\":  \"TotalItemAttachmentsSize\",\"Value\":  \"" + bytesToSize(JSontext.TotalItemAttachmentsSize) +  "\"},";
FolderStatsJson = FolderStatsJson + "{\"Name\":  \"TotalCloudyAttachments\",\"Value\":  \"" + JSontext.TotalCloudyAttachments +  "\"},";
FolderStatsJson = FolderStatsJson + "{\"Name\":  \"TotalCloudyAttachmentsSize\",\"Value\":  \"" + bytesToSize(JSontext.TotalCloudyAttachmentsSize) +  "\"},";
FolderStatsJson = FolderStatsJson + "{\"Name\":  \"LargestAttachmentSize\",\"Value\":  \"" + bytesToSize(JSontext.LargestAttachmentSize) +  "\"},";
FolderStatsJson = FolderStatsJson + "{\"Name\":  \"LargestAttachmentName\",\"Value\":  \"" + JSontext.LargestAttachmentName +  "\"}]";


var FolderStatobj = JSON.parse(FolderStatsJson);

`$("#folderstatsButton").html(("Folder Statistics " + JSontext.FolderPath));
`$("#FolderStats").tabulator("setData", FolderStatobj);
"@

        ($Header + $Json + "`r`n" + $Footer + "</script></html>") | Out-File $fileName

    }
}



function Invoke-ProcessFolder {
    param(
        [Parameter(Position = 0, Mandatory = $true)] [Microsoft.Exchange.WebServices.Data.Folder]$Folder,
        [Parameter(Position = 1, Mandatory = $false)] [switch]$IncludeHidden,
        [Parameter(Position = 2, Mandatory = $false)] [String]$MailboxName
      

    )
    Process {
        $Script:rptCollection = @()
        $rptObj = "" | select MailboxName, FolderPath, TotalFolderItems, TotalFolderItemsSize, TotalItemsNoAttach, TotalItemsNoAttachSize, TotalItemsAttach, TotalItemsAttachSize, TotalFileAttachments, TotalFileAttachmentsSize, TotalItemAttachments, TotalItemAttachmentsSize, TotalCloudyAttachments, TotalCloudyAttachmentsSize, LargestAttachmentSize, LargestAttachmentName, ItemAgeStatistics, AttachmentAgeStatistics, AttachmentExtensions, CloudyAttachmentExtensions
        $rptObj.MailboxName = $MailboxName
        $rptObj.FolderPath = $Folder.FolderPath
        $rptObj.TotalFolderItems = 0
        $rptObj.TotalFolderItemsSize = [Int64]0
        $rptObj.TotalItemsNoAttach = 0
        $rptObj.TotalItemsNoAttachSize = [Int64]0
        $rptObj.TotalItemsAttach = 0
        $rptObj.TotalItemsAttachSize = [Int64]0
        $rptObj.TotalFileAttachments = 0
        $rptObj.TotalCloudyAttachments = 0
        $rptObj.TotalCloudyAttachmentsSize = 0
        $rptObj.TotalFileAttachmentsSize = [Int64]0
        $rptObj.TotalItemAttachments = 0
        $rptObj.TotalItemAttachmentsSize = [Int64]0
        $rptObj.LargestAttachmentSize = [Int64]0
        $rptObj.LargestAttachmentName = ""
        $rptObj.AttachmentAgeStatistics = @()
        $rptObj.AttachmentExtensions = @()
        $rptObj.CloudyAttachmentExtensions = @()

        #Define ItemView to retrive just 1000 Items    
        $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)  
        $psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)  
        $psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Size)
        $psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived)
        $psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated)
        $ivItemView.PropertySet = $psPropset


        #Define Function to convert String to FolderPath  
        function ConvertToString($ipInputString) {  
            $Val1Text = ""  
            for ($clInt = 0; $clInt -lt $ipInputString.length; $clInt++) {  
                $Val1Text = $Val1Text + [Convert]::ToString([Convert]::ToChar([Convert]::ToInt32($ipInputString.Substring($clInt, 2), 16)))  
                $clInt++  
            }  
            return $Val1Text  
        } 
       
        $totalItemCnt = 1
        if ($Folder.TotalCount -ne $null) {
            $totalItemCnt = $Folder.TotalCount
            write-verbose ("Processing FolderPath : " + $Folder.FolderPath + " Item Count " + $totalItemCnt)
        }
        else {
            write-verbose ("Processing FolderPath : " + $Folder.FolderPath)
        }
        $runningCount = 0
        $aptrptHash = @{}
        $msgrptHash = @{}
        $attachmentNameHash = @{}
        $CloudyattachmentNameHash = @{}
        if ($totalItemCnt -gt 0) {
            #Define ItemView to retrive just 1000 Items    
            $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)  
            $psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)  
            $psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Size)
            $psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived)
            $psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated)
            $psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Attachments)
            $psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::HasAttachments)
            $psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject)
            $psPropset.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Sender)
            $fipsPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly) 
            $fipsPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Size)
            $fipsPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived)
            $fipsPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated)
            $fipsPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::HasAttachments)
            $PR_HASATTACH = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0E1B, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Boolean) 
            $fipsPropset.Add($PR_HASATTACH)
            $ivItemView.PropertySet = $fipsPropset		
            $fiItems = $null    
            do { 
                try {
                    $fiItems = $Folder.FindItems($ivItemView) 
                }
                catch {
                    Write-Verbose ("Error " + $_.Exception.InnerException.Message)
                    if ($_.Exception.InnerException -is [Microsoft.Exchange.WebServices.Data.ServiceResponseException]) {
                        if ($_.Exception.InnerException -is [Microsoft.Exchange.WebServices.Data.ServerBusyException]) {
                            $Seconds = [Math]::Round(($_.Exception.InnerException.BackOffMilliseconds / 1000), 0)
                            Write-Verbose ("Resume in " + $Seconds + " Miliseconds")
                            Start-Sleep -Milliseconds $_.Exception.InnerException.BackOffMilliseconds
                        }
                        else {
                            Write-Verbose ("Resume in 60 Seconds")
                            Start-Sleep -Seconds 60 
                        }                    
                       
                    }	
                    $fiItems = $Folder.FindItems($ivItemView) 
                }
                
                if ($fiItems.Items.Count -gt 0) {
                    $runningCount += $fiItems.Items.Count                 
                    write-verbose ("Processed : " + $runningCount + " Items of Total " + $totalItemCnt)
                    $type = ("System.Collections.Generic.List" + '`' + "1") -as "Type"
                    $type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.Item" -as "Type")
                    $AttachmentItems = [Activator]::CreateInstance($type)
                    foreach ($Item in $fiItems.Items) {
                        $rptObj.TotalFolderItems += 1
                        $rptObj.TotalFolderItemsSize += [Int64]$Item.Size
                        if ($IncludeHidden.IsPresent) {
                            $PR_HASATTACHValue = $null
                            if ($Item.TryGetProperty($PR_HASATTACH, [ref]$PR_HASATTACHValue)) {
                                if ($PR_HASATTACHValue) {
                                    $AttachmentItems.Add($Item)
                                }
                                else {
                                    $rptObj.TotalItemsNoAttach += 1
                                    $rptObj.TotalItemsNoAttachSize += [Int64]$Item.Size
                                }
                            }
                        }
                        else {
                            if ($Item.HasAttachments) {
                                $AttachmentItems.Add($Item)
                            }
                            else {
                                $rptObj.TotalItemsNoAttach += 1
                                $rptObj.TotalItemsNoAttachSize += [Int64]$Item.Size
                            }
                        }
                        $dateVal = $null
                        if ($Item.TryGetProperty([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, [ref]$dateVal ) -eq $false) {
                            $dateVal = $Item.DateTimeCreated
                        }
                        if ($msgrptHash.ContainsKey($dateVal.Year)) {
                            $msgrptHash[$dateVal.Year].TotalNumber += 1
                            $msgrptHash[$dateVal.Year].TotalSize += [Int64]$Item.Size
                        }
                        else {
                            $DateObj = "" | Select Date, TotalNumber, TotalSize
                            $DateObj.TotalNumber = 1
                            $DateObj.TotalSize = [Int64]$Item.Size
                            $DateObj.Date = $dateVal.Year
                            $msgrptHash.add($dateVal.Year, $DateObj)
                            if ($dateVal.Year -lt $minYear) {$minYear = $dateVal.Year}
                        }

                    }
                    if ($AttachmentItems.Count -gt 0) {
                        try {
                            [Void]$Folder.Service.LoadPropertiesForItems($AttachmentItems, $psPropset)
                        }
                        catch {
                            Write-Verbose ("Error " + $_.Exception.InnerException.Message)
                            if ($_.Exception.InnerException -is [Microsoft.Exchange.WebServices.Data.ServiceResponseException]) {
                                if ($_.Exception.InnerException -is [Microsoft.Exchange.WebServices.Data.ServerBusyException]) {
                                    $Seconds = [Math]::Round(($_.Exception.InnerException.BackOffMilliseconds / 1000), 0)
                                    Write-Verbose ("Resume in " + $Seconds + " Miliseconds")
                                    Start-Sleep -Milliseconds $_.Exception.InnerException.BackOffMilliseconds
                                }
                                else {
                                    Write-Verbose ("Resume in 60 Seconds")
                                    Start-Sleep -Seconds 60 
                                }                    
                            
                            }	 
                            [Void]$Folder.Service.LoadPropertiesForItems($AttachmentItems, $psPropset)
                        }
                        Write-Verbose ("Processing " + $AttachmentItems.Count + " with Attachments")
                        foreach ($Item in $AttachmentItems) {
                            if ($Item.Attachments.Count -gt 0) {
                                $rptObj.TotalItemsAttach += 1
                                $rptObj.TotalItemsAttachSize += [Int64]$Item.Size
                                foreach ($Attachment in $Item.Attachments) {							
                                    if ($Attachment -is [Microsoft.Exchange.WebServices.Data.FileAttachment]) {
                                        $rptObj.TotalFileAttachments += 1
                                        $rptObj.TotalFileAttachmentsSize += $Attachment.Size
                                        $attachSize = $Attachment.Size
                                        if ($attachSize -gt $rptobj.LargestAttachmentSize) {
                                            $rptobj.LargestAttachmentSize = $attachSize
                                            $rptobj.LargestAttachmentName = $Attachment.Name
                                        }
                                    }
                                    else {
                                        if ($Attachment -is [Microsoft.Exchange.WebServices.Data.ReferenceAttachment]) {
                                            $rptObj.TotalCloudyAttachmentsSize += $Attachment.Size
                                            $rptObj.TotalCloudyAttachments += 1
                                        }
                                        else {
                                            $rptObj.TotalItemAttachments += 1
                                            $rptObj.TotalItemAttachmentsSize += $Attachment.Size
                                        }
                                    }
                                    $dateVal = $null
                                    if ($Item.TryGetProperty([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, [ref]$dateVal ) -eq $false) {
                                        $dateVal = $Item.DateTimeCreated
                                    }
                                    if ($aptrptHash.ContainsKey($dateVal.Year)) {
                                        $aptrptHash[$dateVal.Year].TotalNumber += 1
                                        $aptrptHash[$dateVal.Year].TotalSize += [Int64]$Item.Size
                                    }
                                    else {
                                        $DateObj = "" | Select Date, TotalNumber, TotalSize
                                        $DateObj.TotalNumber = 1
                                        $DateObj.TotalSize = [Int64]$Item.Size
                                        $DateObj.Date = $dateVal.Year
                                        $aptrptHash.add($dateVal.Year, $DateObj)
                                        if ($dateVal.Year -lt $minYear) {$minYear = $dateVal.Year}
                                    }
                                    if (![String]::IsNullOrEmpty($Attachment.Name)) {
                                        $extension = [System.IO.Path]::GetExtension(($Attachment.Name | Remove-InvalidFileNameChars)).Replace(".", "")
                                        if ($attachmentNameHash.ContainsKey($extension)) {
                                            $attachmentNameHash[$extension].TotalNumber += 1
                                            $attachmentNameHash[$extension].TotalSize += [Int64]$Item.Size
                                            if ($attachmentNameHash[$extension].LargestFileSize -lt ([Int64]$Item.Size)) {
                                                $attachmentNameHash[$extension].LargestFileName = $Attachment.Name
                                                $attachmentNameHash[$extension].LargestFileSize = [Int64]$Item.Size
                                                $attachmentNameHash[$extension].LargestFileMessageSubject = $Item.Subject
                                                $attachmentNameHash[$extension].LargestFileMessageSender = $Item.Sender.Address
                                            }
                                        }
                                        else {
                                            $AttachExtObj = "" | Select Extension, LargestFileMessageSubject, LargestFileMessageSender, LargestFileName, LargestFileSize, TotalNumber, TotalSize
                                            $AttachExtObj.TotalNumber = 1
                                            $AttachExtObj.TotalSize = [Int64]$Item.Size
                                            $AttachExtObj.Extension = $extension
                                            $AttachExtObj.LargestFileName = $Attachment.Name 
                                            $AttachExtObj.LargestFileSize = [Int64]$Item.Size
                                            $AttachExtObj.LargestFileMessageSubject = $Item.Subject
                                            $AttachExtObj.LargestFileMessageSender = $Item.Sender.Address
                                            $attachmentNameHash.add($extension, $AttachExtObj)
                                            
                                        }
                                    }
                                    if (![String]::IsNullOrEmpty($Attachment.AttachLongPathName)) {
                                        if($Attachment.AttachLongPathName.Contains("/guestaccess.aspx")){
                                            $extension = "GuestInvite"
                                        }else{
                                            $AttachmentName = $Attachment.AttachLongPathName
                                            if($AttachmentName.Contains("?")){
                                                $indexoflasts= ($Attachment.AttachLongPathName).lastindexof('?')
                                                 $AttachmentName = $Attachment.AttachLongPathName.substring(0,$indexoflasts)
                                            }
                                            $extension = [System.IO.Path]::GetExtension(($AttachmentName | Remove-InvalidFileNameChars)).Replace(".", "")
                                        }                                        
                                        if ($CloudyattachmentNameHash.ContainsKey($extension)) {
                                            $CloudyattachmentNameHash[$extension].TotalNumber += 1
                                            $CloudyattachmentNameHash[$extension].TotalSize += [Int64]$Item.Size
                                        }else{
                                           $AttachExtObj = "" | Select Extension, TotalNumber, TotalSize
                                            $AttachExtObj.TotalNumber = 1
                                            $AttachExtObj.TotalSize = [Int64]$Item.Size
                                            $AttachExtObj.Extension = $extension
                                            $CloudyattachmentNameHash.add($extension, $AttachExtObj)
                                        }
                                    }
                                }
                            }
                            else {
                                $rptObj.TotalItemsNoAttach += 1
                                $rptObj.TotalItemsNoAttachSize += [Int64]$Item.Size
                            }
                        }
                    }                   
                }    
                $ivItemView.Offset += $fiItems.Items.Count    
            }while ($fiItems.MoreAvailable -eq $true)

            $rptObj.AttachmentAgeStatistics = $aptrptHash.Values | Sort-Object Date -Descending
            $rptObj.ItemAgeStatistics = $msgrptHash.Values  | Sort-Object Date -Descending
            $rptObj.AttachmentExtensions = $attachmentNameHash.Values  | Sort-Object Extension 
            $rptObj.CloudyAttachmentExtensions = $CloudyattachmentNameHash.Values  | Sort-Object Extension 
            $Script:rptCollection += $rptObj
        }
    
        return $Script:rptCollection
   
    }
}

$Script:SavePath = (Get-Item -Path ".\").FullName
