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

function CovertBitValue($String){  
    $numItempattern = '(?=\().*(?=bytes)'  
    $matchedItemsNumber = [regex]::matches($String, $numItempattern)   
    $Mb = [INT64]$matchedItemsNumber[0].Value.Replace("(","").Replace(",","")  
    return [math]::round($Mb/1048576,0)  
}  

function Get-UnReadMessageCount{
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [switch]$useImpersonation,
		[Parameter(Position=3, Mandatory=$false)] [string]$url,
		[Parameter(Position=4, Mandatory=$true)] [Int32]$Months
    )  
 	Begin
	{
		$eval1 = "Last" + $Months + "MonthsTotal"
		$eval2 = "Last" + $Months + "MonthsUnread"
		$eval3 = "Last" + $Months + "MonthsSent"
		$eval4 = "Last" + $Months + "MonthsReplyToSender"
		$eval5 = "Last" + $Months + "MonthsReplyToAll"
		$eval6 = "Last" + $Months + "MonthForward"
		$reply = 0;
		$replyall = 0
		$forward = 0
		$rptObj = "" | select  MailboxName,Mailboxsize,LastLogon,LastLogonAccount,$eval1,$eval2,$eval4,$eval5,$eval6,LastMailRecieved,$eval3,LastMailSent  
		$rptObj.MailboxName = $MailboxName  
		if($url){
			$service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials -url $url 
		}
		else{
			$service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials
		}
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
		}
		$AQSString1 = "System.Message.DateReceived:>" + [system.DateTime]::Now.AddMonths(-$Months).ToString("yyyy-MM-dd")   
  		$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxName)     
		$Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)  
  		$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems,$MailboxName)     
		$SentItems = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)    
		$ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)  
		$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)  
		$psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived)
		$psPropset.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsRead)
		$PidTagLastVerbExecuted = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x1081,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
		#Find Items Greater then 35 MB
		$sfItemSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,[system.DateTime]::Now.AddMonths(-$Months)) 
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
	   	do{ 
			$fiItems = $Inbox.findItems($sfItemSearchFilter,$ivItemView)  
			if($settc){
				$rptObj.$eval1 = $fiItems.TotalCount  
				write-host ("Last " + $Months + " Months : " + $fiItems.TotalCount)
				if($fiItems.TotalCount -gt 0){  
			    	write-host ("Last Mail Recieved : " + $fiItems.Items[0].DateTimeReceived ) 
			    	$rptObj.LastMailRecieved = $fiItems.Items[0].DateTimeReceived  
				}		    
				$settc = $false
			}
			    foreach($Item in $fiItems.Items){
					$unReadVal = $null
					if($Item.TryGetProperty([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsRead,[ref]$unReadVal)){
						if(!$unReadVal){
							$unreadCount++
						}
					} 
				   $lastVerb = $null
					if($Item.TryGetProperty($PidTagLastVerbExecuted,[ref]$lastVerb)){
						switch($lastVerb){
							102 { $reply++ }
							103 { $replyall++}
							104 { $forward++}
						}
					} 
			    }    
			    $ivItemView.Offset += $fiItems.Items.Count    
			}while($fiItems.MoreAvailable -eq $true) 

		write-host ("Last " + $Months + " Months Unread : " + $unreadCount ) 
		$rptObj.$eval2 = $unreadCount  
		$rptObj.$eval4 = $reply
		$rptObj.$eval5 = $replyall
		$rptObj.$eval6 = $forward
		$ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1)  
		$fiResults = $SentItems.findItems($sfItemSearchFilter,$ivItemView)  
		write-host ("Last " + $Months + " Months Sent : " + $fiResults.TotalCount  )
		$rptObj.$eval3 = $fiResults.TotalCount  
		if($fiResults.TotalCount -gt 0){  
		    write-host ("Last Mail Sent Date : " + $fiResults.Items[0].DateTimeSent  )
		    $rptObj.LastMailSent = $fiResults.Items[0].DateTimeSent  
		}  
		Write-Output $rptObj  
	}
}