function Connect-Exchange{ 
    param( 
    		[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials
    )  
 	Begin
		 {
		Load-EWSManagedAPI
		
		## Set Exchange Version  
		$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1
		  
		## Create Exchange Service Object  
		$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)  
		  
		## Set Credentials to use two options are availible Option1 to use explict credentials or Option 2 use the Default (logged On) credentials  
		  
		#Credentials Option 1 using UPN for the windows Account  
		#$psCred = Get-Credential  
		$creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(),$Credentials.GetNetworkCredential().password.ToString())  
		$service.Credentials = $creds      
		#Credentials Option 2  
		#service.UseDefaultCredentials = $true  
		 #$service.TraceEnabled = $true
		## Choose to ignore any SSL Warning issues caused by Self Signed Certificates  
		  
		Handle-SSL	
		  
		## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use  
		  
		#CAS URL Option 1 Autodiscover  
		$service.AutodiscoverUrl($MailboxName,{$true})  
		Write-host ("Using CAS Server : " + $Service.url)   
		   
		#CAS URL Option 2 Hardcoded  
		  
		#$uri=[system.URI] "https://casservername/ews/exchange.asmx"  
		#$service.Url = $uri    
		  
		## Optional section for Exchange Impersonation  
		  
		#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
		if(!$service.URL){
			throw "Error connecting to EWS"
		}
		else
		{		
			return $service
		}
	}
}

function Load-EWSManagedAPI{
    param( 
    )  
 	Begin
	{
		## Load Managed API dll  
		###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
		$EWSDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
		if (Test-Path $EWSDLL)
		    {
		    Import-Module $EWSDLL
		    }
		else
		    {
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

function Handle-SSL{
    param( 
    )  
 	Begin
	{
		## Code From http://poshcode.org/624
		## Create a compilation environment
		$Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
		$Compiler=$Provider.CreateCompiler()
		$Params=New-Object System.CodeDom.Compiler.CompilerParameters
		$Params.GenerateExecutable=$False
		$Params.GenerateInMemory=$True
		$Params.IncludeDebugInformation=$False
		$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

$TASource=@'
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
		$TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
		$TAAssembly=$TAResults.CompiledAssembly

		## We now create an instance of the TrustAll and attach it to the ServicePointManager
		$TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
		[System.Net.ServicePointManager]::CertificatePolicy=$TrustAll

		## end code from http://poshcode.org/624

	}
}

function ConvertId{    
 param (
         $HexId = "$( throw 'HexId is a mandatory Parameter' )"
    )
 process{
     $aiItem = New-Object Microsoft.Exchange.WebServices.Data.AlternateId      
     $aiItem.Mailbox = $MailboxName      
     $aiItem.UniqueId = $HexId   
     $aiItem.Format = [Microsoft.Exchange.WebServices.Data.IdFormat]::HexEntryId      
     $convertedId = $service.ConvertId($aiItem, [Microsoft.Exchange.WebServices.Data.IdFormat]::EwsId) 
     return $convertedId.UniqueId
 }
}

function Get-BirthDayCalendar{
	param (
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [switch]$useImpersonation		  )
	process{
        
        $service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}
        $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)   
        $BirthdayCalendarFolderEntryId = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Common,"BirthdayCalendarFolderEntryId",[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary); 
        $psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
        $psPropset.Add($BirthdayCalendarFolderEntryId)
		$EWSRootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid,$psPropset)
        $BirthdayCalendarFolderEntryIdValue = $null
		$BirthdayCalendarFolderHexValue = $null
        if($EWSRootFolder.TryGetProperty($BirthdayCalendarFolderEntryId,[ref]$BirthdayCalendarFolderEntryIdValue)){
			$BirthdayFolderEWSId = new-object Microsoft.Exchange.WebServices.Data.FolderId((ConvertId -HexId ([System.BitConverter]::ToString($BirthdayCalendarFolderEntryIdValue).Replace("-",""))))
			$BirthdayFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$BirthdayFolderEWSId);
			return $BirthdayFolder
        }
		else
		{   
        	 throw [System.IO.FileNotFoundException] "folder not found."
		}
    }
}

function Get-Birthdays{
    	param (
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [switch]$useImpersonation		  )
	process{
        
        $service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}
        $strtime = (Get-date).Year.ToString() + "0101"    
        $endtime = (Get-date).AddYears(1).Year.ToString() + "0101"           
        $StartDate =  [datetime]::ParseExact($strtime,"yyyyMMdd",$null)
        $EndDate =  [datetime]::ParseExact($endtime,"yyyyMMdd",$null)
        $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)   
        $BirthdayCalendarFolderEntryId = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Common,"BirthdayCalendarFolderEntryId",[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary); 
        $psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
        $psPropset.Add($BirthdayCalendarFolderEntryId)
		$EWSRootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid,$psPropset)
        $BirthdayCalendarFolderEntryIdValue = $null
		$BirthdayCalendarFolderHexValue = $null
        if($EWSRootFolder.TryGetProperty($BirthdayCalendarFolderEntryId,[ref]$BirthdayCalendarFolderEntryIdValue)){
			$BirthdayFolderEWSId = new-object Microsoft.Exchange.WebServices.Data.FolderId((ConvertId -HexId ([System.BitConverter]::ToString($BirthdayCalendarFolderEntryIdValue).Replace("-",""))))
			$BirthdayFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$BirthdayFolderEWSId);
			$CalendarView = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($StartDate,$EndDate,1000)    
            $psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
            $BirthDayLocal = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Address,0x80DE, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime)
            $psPropset.Add($BirthDayLocal)
            $CalendarView.PropertySet = $psPropset
            $fiItems = $service.FindAppointments($BirthdayFolder.Id,$CalendarView)    
            foreach($Item in $fiItems.Items){      
                $exportObj = "" | Select subject,StartTime,EndTime,DateOfBirth,Age
                $exportObj.StartTime = $Item.Start
                $exportObj.EndTime = $Item.End
                $exportObj.Subject = $Item.Subject
                $BirthDavLocalValue = $null
                if($Item.TryGetProperty($BirthDayLocal,[ref]$BirthDavLocalValue)){
                    $exportObj.DateOfBirth = $BirthDavLocalValue
                    $exportObj.Age = [Math]::Truncate(($Item.Start – $BirthDavLocalValue).TotalDays / 365); 
                }
                Write-Output $exportObj
            }
        }
		else
		{   
        	 throw [System.IO.FileNotFoundException] "folder not found."
		}
    }
    
}