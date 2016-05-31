function Connect-Exchange{ 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [string]$url
    )  
 	Begin
		 {
		Load-EWSManagedAPI
		
		## Set Exchange Version  
		$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
		  
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
		if($url){
			$uri=[system.URI] $url
			$service.Url = $uri    
		}
		else{
			$service.AutodiscoverUrl($MailboxName,{$true})  
		}
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

function Get-CalendarClients
{
        [CmdletBinding()] 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [switch]$useImpersonation,
		[Parameter(Position=3, Mandatory=$false)] [string]$url
    )  
 	Begin
		 {
            if($url)
            {
                $service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials -url $url 
            }
            else{
                $service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials
            }
            if($useImpersonation.IsPresent)
            {
                $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
            }
            # Bind to Calendar
            $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$MailboxName)   
            $Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
            #Define ItemView to retrive just 1000 Items    
            $psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
            $calendarClientInfo = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::CalendarAssistant,0xB,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);
            $psPropset.Add($calendarClientInfo)
            $ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)   
            $ivItemView.PropertySet = $psPropset 
            $SummaryList = @{}
            $appointmentList = @()
            $fiItems = $null    
            do{    
                $fiItems = $service.FindItems($Calendar.Id,$ivItemView)    
                #[Void]$service.LoadPropertiesForItems($fiItems,$psPropset)  
                foreach($Item in $fiItems.Items){
                    $calendarClientInfoVal = $null
                    if($Item.TryGetProperty($calendarClientInfo,[ref]$calendarClientInfoVal)){
                         if(!$SummaryList.ContainsKey($calendarClientInfoVal)){
                             $SummaryObj = "" | Select Client,NumberOfAppointments
                             $SummaryObj.Client = $calendarClientInfoVal
                             $SummaryObj.NumberOfAppointments =1
                             $SummaryList.Add($calendarClientInfoVal,$SummaryObj)
                         }
                         else{
                             $SummaryList[$calendarClientInfoVal].NumberOfAppointments++
                         }
                        $rptObj = "" | Select Client,StartTime,EndTime,Duration,Subject,Type,Location
                        $rptObj.Client = $calendarClientInfoVal
                        $rptObj.StartTime = $Item.Start  
                        $rptObj.EndTime = $Item.End  
                        $rptObj.Duration = $Item.Duration
                        $rptObj.Subject  = $Item.Subject   
                        $rptObj.Type = $Item.AppointmentType
                        $rptObj.Location = $Item.Location
                        $appointmentList += $rptObj                       
                    }                              
                }    
                $ivItemView.Offset += $fiItems.Items.Count    
            }while($fiItems.MoreAvailable -eq $true)                        
            Write-Output $SummaryList.Values
            Write-Host
            $SummaryFileName = 'c:\temp\' + $MailboxName + '-Summary-AppointmentList.csv'
            $SummaryList.Values | Export-Csv -NoTypeInformation -Path $SummaryFileName 
            Write-Host ("Appointments Summary Report created in " + $SummaryFileName)
            $FileName = 'c:\temp\' + $MailboxName + '-AppointmentList.csv'
            $appointmentList | Export-Csv -NoTypeInformation -Path $FileName 
            Write-Host ("Appointments Report created in " + $FileName)
         }
    
}