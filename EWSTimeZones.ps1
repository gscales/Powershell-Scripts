function Connect-Exchange{ 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$false)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [string]$url,
		[Parameter(Position=3, Mandatory=$false)] [switch]$useCurrentLoggedOnCredentials,
		[Parameter(Position=4, Mandatory=$false)] [String]$TimeZone,
		[Parameter(Position=5, Mandatory=$false)] [PsObject]$ExistingService
    )  
 	Begin
		 {
		Load-EWSManagedAPI
		
		## Set Exchange Version  
		$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
		  
		## Create Exchange Service Object  
		if($TimeZone)
		{
			$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion,[System.TimeZoneInfo]::FindSystemTimeZoneById($TimeZone))
		}
		else
		{
			$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)  
		}
		## Set Credentials to use two options are availible Option1 to use explict credentials or Option 2 use the Default (logged On) credentials  
		  
		#Credentials Option 1 using UPN for the windows Account  
		#$psCred = Get-Credential  
		if($Credentials.IsPresent)
		{
			$creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(),$Credentials.GetNetworkCredential().password.ToString())  
			$service.Credentials = $creds
		}
		else
		{
			if($useCurrentLoggedOnCredentials.IsPresent)
			{
				$service.UseDefaultCredentials = $true  
			}
			else
			{
				if($ExistingService)
				{
					$service.Credentials = $ExistingService.Credentials
				}
				else
				{
					$psCred = Get-Credential  
					$creds = New-Object System.Net.NetworkCredential($psCred.UserName.ToString(),$psCred.GetNetworkCredential().password.ToString())  
					$service.Credentials = $creds
				}
			}
		}
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

function GetWorkHoursTimeZone
{
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [String]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service
    )
	Begin
	{
			# Bind to the Calendar Folder
			$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$MailboxName)   
			$Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
			$sfFolderSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass, "IPM.Configuration.WorkHours") 
			#Define ItemView to retrive just 1000 Items    
			$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1) 
			$ivItemView.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated
			$fiItems = $service.FindItems($Calendar.Id,$sfFolderSearchFilter,$ivItemView) 
			if($fiItems.Items.Count -eq 1)
			{
				$UsrConfig = [Microsoft.Exchange.WebServices.Data.UserConfiguration]::Bind($service, "WorkHours", $Calendar.Id, [Microsoft.Exchange.WebServices.Data.UserConfigurationProperties]::All)  
				[XML]$WorkHoursXMLString = [System.Text.Encoding]::UTF8.GetString($UsrConfig.XmlData) 
				$returnVal = $WorkHoursXMLString.Root.WorkHoursVersion1.TimeZone.Name
				write-host ("Parsed TimeZone : " + $returnVal)
				Write-Output $returnVal
			}
			else
			{
				write-host ("No Workhours Object in Mailbox")
				Write-Output $null
			}
	}
}
function GetOWATimeZone{
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [String]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service
    )
	Begin
	{
			# Bind to the RootFolder Folder
			$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)   
			$RootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
			$sfFolderSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass, "IPM.Configuration.OWA.UserOptions") 
			#Define ItemView to retrive just 1000 Items    
			$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1) 
			$ivItemView.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated
			$fiItems = $service.FindItems($RootFolder.Id,$sfFolderSearchFilter,$ivItemView) 
			if($fiItems.Items.Count -eq 1)
			{
				$UsrConfig = [Microsoft.Exchange.WebServices.Data.UserConfiguration]::Bind($service, "OWA.UserOptions", $RootFolder.Id, [Microsoft.Exchange.WebServices.Data.UserConfigurationProperties]::All)  
				if($UsrConfig.Dictionary.ContainsKey("timezone"))
				{
					$returnVal = $UsrConfig.Dictionary["timezone"]
					write-host ("OWA TimeZone : " + $returnVal)
					Write-Output $returnVal
				}
				else
				{
					write-host ("TimeZone not set")
					Write-Output $null
				}
				
				
			}
			else
			{
				write-host ("No Workhours OWAConfig for Mailbox")
				Write-Output $null
			}
	}
}

function GetTimeZoneFromCalendarEvents{
    param( 
		[Parameter(Position=0, Mandatory=$true)] [String]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service
    )
	Begin
	{
		$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$MailboxName)   
		$Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
		$tzList = [TimeZoneInfo]::GetSystemTimeZones()
		$AppointmentStateFlags = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Appointment,0x8217,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);  
		$ResponseStatus = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Appointment,0x8218,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);  
		$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(150)  
		$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
		$psPropset.Add($AppointmentStateFlags)
		$psPropset.Add($ResponseStatus)
		$ivItemView.PropertySet = $psPropset
		$fiItems = $service.FindItems($Calendar.Id,$ivItemView)
		$tzCount = @{}
		if($fiItems.Items.Count -gt 0){
			foreach($Item in $fiItems.Items){
				$AppointmentStateFlagsVal = $null
				[Void]$Item.TryGetProperty($AppointmentStateFlags,[ref]$AppointmentStateFlagsVal)
				$ResponseStatusVal = $null
				[Void]$Item.TryGetProperty($ResponseStatus,[ref]$ResponseStatusVal)
				if($ResponseStatusVal -eq "Organizer" -bor $AppointmentStateFlagsVal -eq 0)
				{
					if($tzCount.ContainsKey($Item.TimeZone))
					{
						$tzCount[$Item.TimeZone]++
					}
					else
					{
						$tzCount.Add($Item.TimeZone,1)
					}
				}
			}    
		}
		$returnVal = $null
		if($tzCount.Count -gt 0){
			$fav = ""
			$tzCount.GetEnumerator() | sort -Property Value -Descending | foreach-object {
			   if($fav -eq "")
			   {
			      $fav = $_.Key
			   }
			}
 			foreach($tz in $tzList)
			{
				if($tz.DisplayName -eq $fav)
				{
					$returnVal = $tz.Id
				}
			}
		}
		Write-Host ("TimeZone From Calendar Appointments : " + $returnVal)
		Write-Output $returnVal
	}
}

function Show-MailboxTimeZone{
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$false)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [switch]$useImpersonation
	)
	Begin
	{
	    $service = Connect-Exchange -MailboxName $MailboxName
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
		}
		Write-Host "Outlook settings"
		GetWorkHoursTimeZone -service $service -MailboxName $MailboxName
		Write-Host "OWA settings"
		GetOWATimeZone -service $service -MailboxName $MailboxName
		Write-Host "Calendar Appointment Settings"
		GetTimeZoneFromCalendarEvents -service $service -MailboxName $MailboxName
	}
}



