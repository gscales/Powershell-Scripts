function Connect-Exchange{ 
    param( 
    		[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials
    )  
 	Begin
		 {
		Load-EWSManagedAPI
		
		## Set Exchange Version  
		$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
		  
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
function Get-ScriptDirectory
{
  $Invocation = (Get-Variable MyInvocation -Scope 1).Value
  Split-Path $Invocation.MyCommand.Path
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
		 [Parameter(Position=1, Mandatory=$false)] [String]$HexId,
    	 [Parameter(Position=2, Mandatory=$false)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service
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
function Get-TeamChatStats{
        param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position=2, Mandatory=$false)] [switch]$useImpersonation
    )  
 	Begin
	 {
        $service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
        }
		# Bind to the Root Folder
		$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)   
		$TeamChatFolderEntryId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([System.Guid]::Parse("{E49D64DA-9F3B-41AC-9684-C6E01F30CDFA}"), "TeamChatFolderEntryId", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
		$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
		$psPropset.Add($TeamChatFolderEntryId)
		$RootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid,$psPropset)
		$FolderIdVal = $null
		$PR_MESSAGE_SIZE_EXTENDED = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3592, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
		$PR_ATTACH_ON_NORMAL_MSG_COUNT = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x66B1, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long);
		$Propset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
		$Propset.add($PR_MESSAGE_SIZE_EXTENDED)
		$Propset.add($PR_ATTACH_ON_NORMAL_MSG_COUNT)
		$rptObject = "" | Select Mailbox,RecipientCount,Recipients,ConversationCount,TotalAttachmentCount,TotalItemCount,TotalFolderSizeMB
		if ($RootFolder.TryGetProperty($TeamChatFolderEntryId,[ref]$FolderIdVal))
		{
			$TeamChatFolderId= new-object Microsoft.Exchange.WebServices.Data.FolderId((ConvertId -HexId ([System.BitConverter]::ToString($FolderIdVal).Replace("-","")) -service $service))
			$TeamChatFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$TeamChatFolderId,$Propset);
			$rptObject.TotalItemCount = $TeamChatFolder.TotalCount
			$rptObject.Mailbox = $MailboxName
			$folderSize = $null
			[Void]$TeamChatFolder.TryGetProperty($PR_MESSAGE_SIZE_EXTENDED, [ref]$folderSize)
			[Int64]$attachcnt = 0
			[Void]$TeamChatFolder.TryGetProperty($PR_ATTACH_ON_NORMAL_MSG_COUNT,[ref] $attachcnt)
			if($attachcnt -eq $null){
				$attachcnt = 0
			}
			$rptObject.TotalAttachmentCount = $attachcnt
			$rptObject.TotalFolderSizeMB  = [math]::round($folderSize /1Mb, 3)
			$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000) 
			$SkypeMessagePropertyBag = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([System.Guid]::Parse("{0A63905C-12A3-42E8-9441-793FC61F4670}"), "SkypeMessagePropertyBag", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);
			$ItemPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
			$ItemPropset.Add($SkypeMessagePropertyBag)
			$rptCollection = @{}
			$rptCollection2 = @{}
			$fiItems = $null    
			do{    
				$fiItems = $service.FindItems($TeamChatFolder.Id,$ivItemView)   
				Write-Host ("Processed " + $fiItems.Items.Count)
				[Void]$service.LoadPropertiesForItems($fiItems,$ItemPropset)  
				foreach($Item in $fiItems.Items){      
					#Process Item
					foreach($rcp in $Item.ToRecipients){

						if(!$rptCollection.ContainsKey($rcp.Address)){
							$rptObject.RecipientCount++
							$rptCollection.add($rcp.Address,1)							
						}
						else{							
							$rptCollection[$rcp.Address]++
						}
						$SkypeMessagePropertyBagValue = $null
						if($Item.TryGetProperty($SkypeMessagePropertyBag,[ref]$SkypeMessagePropertyBagValue)){
							$SkypeProps = ConvertFrom-Json -InputObject $SkypeMessagePropertyBagValue
							if(![String]::IsNullOrEmpty($SkypeProps.conversationId)){
								if(!$rptCollection2.ContainsKey($SkypeProps.conversationId)){
									$rptObject.ConversationCount++
									$rptCollection2.add($SkypeProps.conversationId,1)							
								}
								else{							
									$rptCollection2[$SkypeProps.conversationId]++
								}
							}
						}
					}
				}    
				$ivItemView.Offset += $fiItems.Items.Count    
			}while($fiItems.MoreAvailable -eq $true) 
		}
		else{
			write-host ("Team Chat Folder not found")
		}
		$rptObject.Recipients = $rptCollection
		return $rptObject
  
     }
}

function Get-TeamChatSkypeMessagePropertyBag{
        param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position=2, Mandatory=$false)] [switch]$useImpersonation
    )  
 	Begin
	 {
        $service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
        }
		# Bind to the Root Folder
		$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)   
		$TeamChatFolderEntryId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([System.Guid]::Parse("{E49D64DA-9F3B-41AC-9684-C6E01F30CDFA}"), "TeamChatFolderEntryId", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
		$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
		$psPropset.Add($TeamChatFolderEntryId)
		$RootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid,$psPropset)
		$FolderIdVal = $null
		if ($RootFolder.TryGetProperty($TeamChatFolderEntryId,[ref]$FolderIdVal))
		{
			$TeamChatFolderId= new-object Microsoft.Exchange.WebServices.Data.FolderId((ConvertId -HexId ([System.BitConverter]::ToString($FolderIdVal).Replace("-","")) -service $service))
			$TeamChatFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$TeamChatFolderId);
			$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000) 
			$SkypeMessagePropertyBag = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([System.Guid]::Parse("{0A63905C-12A3-42E8-9441-793FC61F4670}"), "SkypeMessagePropertyBag", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);
			$ItemPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
			$ItemPropset.Add($SkypeMessagePropertyBag)
			$fiItems = $null    
			do{    
				$fiItems = $service.FindItems($TeamChatFolder.Id,$ivItemView)   
				Write-Host ("Processed " + $fiItems.Items.Count)
				[Void]$service.LoadPropertiesForItems($fiItems,$ItemPropset)  
				foreach($Item in $fiItems.Items){      
					#Process Item
					foreach($rcp in $Item.ToRecipients){
						$SkypeMessagePropertyBagValue = $null
						if($Item.TryGetProperty($SkypeMessagePropertyBag,[ref]$SkypeMessagePropertyBagValue)){
								$SkypeProps = ConvertFrom-Json -InputObject $SkypeMessagePropertyBagValue
								Write-Output $SkypeProps
						}
					}
				}    
				$ivItemView.Offset += $fiItems.Items.Count    
			}while($fiItems.MoreAvailable -eq $true) 
		}
		else{
			write-host ("Team Chat Folder not found")
		}  
     }
}

