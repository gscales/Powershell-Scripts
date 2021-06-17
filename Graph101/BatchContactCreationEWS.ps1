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
    $Navigated = {
      if($web.DocumentText -match "document.location.replace"){
        $Script:oAuthCode = [regex]::match($web.DocumentText, "code=(.*?)\\u0026").Groups[1].Value
        $form.Close();
      }
    }    
    $web.ScriptErrorsSuppressed = $true
    $web.Add_Navigated($Navigated)
    $form.Controls.Add($web)
    $form.Add_Shown( { $form.Activate() })
    $form.ShowDialog() | Out-Null
    return $Script:oAuthCode
}

function Get-AccessTokenForGraph {
    [CmdletBinding()]
    param (   
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $MailboxName,
        [Parameter(Position = 2, Mandatory = $false)]
        [String]
        $ClientId,
        [Parameter(Position = 3, Mandatory = $false)]
        [String]
        $RedirectURI = "urn:ietf:wg:oauth:2.0:oob",
        [Parameter(Position = 4, Mandatory = $false)]
        [String]
        $scopes = "User.Read.All Mail.Read",
        [Parameter(Position = 5, Mandatory = $false)]
        [switch]
        $Prompt

    )
    Process {
 
        if ([String]::IsNullOrEmpty($ClientId)) {
            $ClientId = "5471030d-f311-4c5d-91ef-74ca885463a7"
        }		
        $Domain = $MailboxName.Split('@')[1]
        $TenantId = (Invoke-WebRequest ("https://login.windows.net/" + $Domain + "/v2.0/.well-known/openid-configuration") | ConvertFrom-Json).token_endpoint.Split('/')[3]
        Add-Type -AssemblyName System.Web, PresentationFramework, PresentationCore
        $state = Get-Random
        $authURI = "https://login.microsoftonline.com/$TenantId"
        $authURI += "/oauth2/v2.0/authorize?client_id=$ClientId"
        $authURI += "&response_type=code&redirect_uri= " + [System.Web.HttpUtility]::UrlEncode($RedirectURI)
        $authURI += "&response_mode=query&scope=" + [System.Web.HttpUtility]::UrlEncode($scopes) + "&state=$state"
        if ($Prompt.IsPresent) {
            $authURI += "&prompt=select_account"
        }else{
            $authURI += "&prompt=none&login_hint=" + $MailboxName
        }    

        # Extract code from query string
        $authCode = Show-OAuthWindow -Url $authURI
        $Body = @{"grant_type" = "authorization_code"; "scope" = $scopes; "client_id" = "$ClientId"; "code" = $authCode; "redirect_uri" = $RedirectURI }
        $tokenRequest = Invoke-RestMethod -Method Post -ContentType application/x-www-form-urlencoded -Uri "https://login.microsoftonline.com/$tenantid/oauth2/token" -Body $Body 
        $AccessToken = $tokenRequest.access_token
        return $AccessToken
		
    }
    
}

function Get-AccessTokenForGraphFromCertificate{
    param(
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $TenantDomain,
        [Parameter(Position = 2, Mandatory = $true)]
        [String]
        $ClientId,
		[Parameter(Position = 3, Mandatory = $false)]
		[System.Security.Cryptography.X509Certificates.X509Certificate2]
        $Certificate,
        [Parameter(Position = 2, Mandatory = $false)]
        [String]
        $Scope = "https://outlook.office365.com/.default"
         
    )
    Process{       
        
        # Create base64 hash of certificate
        $CertificateBase64Hash = [System.Convert]::ToBase64String($Certificate.GetCertHash()) -replace '\+','-' -replace '/','_' -replace '='
        
        # Create Token Timestamps
        $StartDate = (Get-Date "1970-01-01T00:00:00Z" ).ToUniversalTime()
        $TokenExpiration = [math]::Round(((New-TimeSpan -Start $StartDate -End (Get-Date).ToUniversalTime().AddMinutes(2)).TotalSeconds),0)
        $NotBefore = [math]::Round(((New-TimeSpan -Start $StartDate -End ((Get-Date).ToUniversalTime())).TotalSeconds),0)
        
        $ClientAssertionheader = @{
            alg = "RS256"
            typ = "JWT"           
            x5t = $CertificateBase64Hash 
        }        
        $ClientAssertionPayLoad = @{           
            aud = "https://login.microsoftonline.com/$TenantDomain/oauth2/token"        
            exp = $TokenExpiration
            iss = $ClientId
            jti = [guid]::NewGuid()
            nbf = $NotBefore
            sub = $ClientId
        }
        $CAEncodedHeader = [System.Convert]::ToBase64String(([System.Text.Encoding]::UTF8.GetBytes(($ClientAssertionheader | ConvertTo-Json)))) -replace '\+','-' -replace '/','_' -replace '='
        $CAEncodedPayload = [System.Convert]::ToBase64String(([System.Text.Encoding]::UTF8.GetBytes(($ClientAssertionPayLoad | ConvertTo-Json)))) -replace '\+','-' -replace '/','_' -replace '=' 
        
        # Get the private key object of your certificate
        $PrivateKey = ([System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate))
        $RSAPadding = [Security.Cryptography.RSASignaturePadding]::Pkcs1
        $HashAlgorithm = [Security.Cryptography.HashAlgorithmName]::SHA256
        
        # Sign the Assertion
        $Signature = [Convert]::ToBase64String(
            $PrivateKey.SignData([System.Text.Encoding]::UTF8.GetBytes(($CAEncodedHeader + "." + $CAEncodedPayload)),$HashAlgorithm,$RSAPadding)
        ) -replace '\+','-' -replace '/','_' -replace '='
        
        # Create the assertion token
        $ClientAssertion = $CAEncodedHeader + "." + $CAEncodedPayload + "." + $Signature
        
        # Create a hash with body parameters
        $Body = @{
            client_id = $ClientId
            client_assertion = $ClientAssertion
            client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
            scope = $Scope
            grant_type = "client_credentials"        
        }
        
        $AuthUrl = "https://login.microsoftonline.com/$TenantDomain/oauth2/v2.0/token"
        
        return Invoke-RestMethod -Headers $Header -Method POST -Uri $AuthUrl -Body $Body -ContentType 'application/x-www-form-urlencoded'
     
    }
}

function Import-ContactsFromCSV(){
    param(
    [Parameter(Position = 1, Mandatory = $true)]
    [String]
    $CertificateThumbPrint,
    
    [Parameter(Position = 2, Mandatory = $true)]
    [string]
    $ClientId,		   
        
    [Parameter(Position =3, Mandatory = $true)]
    [string]
    $TenantDomain,

    [Parameter(Position = 4,Mandatory = $true)]
    [string]
    $TargetMailbox,

    [Parameter(Position = 5,Mandatory = $true)]
    [string]
    $CSVFile)
    Process{
        $rptObject = "" | Select MailboxName,ContactSucess,ErrorCount,Errors,TimeToRun
        $ScriptStartTime = Get-Date
        $rptObject.MailboxName = $TargetMailbox
        $Certificate = Get-Item ("Cert:\CurrentUser\My\$CertificateThumbPrint")
        $token = Get-AccessTokenForGraphFromCertificate -TenantDomain $TenantDomain -ClientId $ClientId -Certificate $Certificate        
        $AccessToken = $token.access_token
        $service = Connect-Exchange -MailboxName $TargetMailbox -AccessToken $AccessToken -url "https://outlook.office365.com/ews/exchange.asmx" 
        $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $TargetMailbox)
        $service.HttpHeaders.Add("X-AnchorMailbox", $TargetMailbox)
        $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::contacts,$TargetMailbox)         
        $TargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderid) 
        $type = ("System.Collections.Generic.List" + '`' + "1") -as "Type"
        $type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.Contact" -as "Type")
        $ContactBatch = [Activator]::CreateInstance($type)
        Import-Csv -Path $CSVFile | ForEach-Object{
            $User = $_            
            $Contact = New-EXCBatchContact -service $service -FirstName $User.FirstName -LastName $User.LastName -EmailAddress $User.Email -EmailAddressDisplayAs $User.DisplayName -fileAs ($User.LastName + "," +  $User.FirstName) -JobTitle $User.Title -Department $User.Dept -CompanyName $User.Company -MobilePhone $User.Mobile -BusinssPhone $User.TelephoneNumber
            $ContactBatch.Add($Contact)
            if($ContactBatch.Count -gt 60){
                $createResult = $service.CreateItems($ContactBatch,$TargetFolder.Id, $null, $null)
                if($createResult){
                    foreach ($Response in $createResult) {       
                        if ($Response.Result -eq [Microsoft.Exchange.WebServices.Data.ServiceResult]::Success) {
                            $rptObject.ContactSucess++
                        }else{
                            $rptObject.ErrorCount++
                            $rptObject.Errors += $Response.ErrorCode
                        }
                    }
                }
                write-verbose("Created " + $ContactBatch.Count + " contacts")
                $ContactBatch = [Activator]::CreateInstance($type)
               
            }
            $User = $null
        }
        if($ContactBatch.Count -gt 0){
            $createResult = $service.CreateItems($ContactBatch,$TargetFolder.Id, $null, $null)           
            if($createResult){
                foreach ($Response in $createResult) {       
                    if ($Response.Result -eq [Microsoft.Exchange.WebServices.Data.ServiceResult]::Success) {
                        $rptObject.ContactSucess++
                    }else{
                        $rptObject.ErrorCount++
                        $rptObject.Errors += $Response.ErrorCode
                    }
                }
            }
        }
        $rptObject.TimeToRun = [Math]::Round((New-TimeSpan -Start $ScriptStartTime -End (Get-Date)).TotalSeconds,2)   
        return $rptObject
        
    }
}

function New-EXCBatchContact
{
<#
	.SYNOPSIS
		Creates a Contact to be used in a Batch request Exchange Web Services API
	
	.DESCRIPTION
		Creates a Contact to be used in a Batch request Exchange Web Services API
		
		Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951
	

	.PARAMETER FirstName
		A description of the FirstName parameter.
	
	.PARAMETER LastName
		A description of the LastName parameter.
	
	.PARAMETER EmailAddress
		A description of the EmailAddress parameter.
	
	.PARAMETER CompanyName
		A description of the CompanyName parameter.
	
	.PARAMETER Credentials
		A description of the Credentials parameter.
	
	.PARAMETER Department
		A description of the Department parameter.
	
	.PARAMETER Office
		A description of the Office parameter.
	
	.PARAMETER BusinssPhone
		A description of the BusinssPhone parameter.
	
	.PARAMETER MobilePhone
		A description of the MobilePhone parameter.
	
	.PARAMETER HomePhone
		A description of the HomePhone parameter.
	
	.PARAMETER IMAddress
		A description of the IMAddress parameter.
	
	.PARAMETER Street
		A description of the Street parameter.
	
	.PARAMETER City
		A description of the City parameter.
	
	.PARAMETER State
		A description of the State parameter.
	
	.PARAMETER PostalCode
		A description of the PostalCode parameter.
	
	.PARAMETER Country
		A description of the Country parameter.
	
	.PARAMETER JobTitle
		A description of the JobTitle parameter.
	
	.PARAMETER Notes
		A description of the Notes parameter.
	
	.PARAMETER Photo
		A description of the Photo parameter.
	
	.PARAMETER FileAs
		A description of the FileAs parameter.
	
	.PARAMETER WebSite
		A description of the WebSite parameter.
	
	.PARAMETER Title
		A description of the Title parameter.
	
	.PARAMETER Folder
		A description of the Folder parameter.
	
	.PARAMETER EmailAddressDisplayAs
		A description of the EmailAddressDisplayAs parameter.
	
	.PARAMETER useImpersonation
		A description of the useImpersonation parameter.
	
	
#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 1, Mandatory = $true)]
		[Microsoft.Exchange.WebServices.Data.ExchangeService]
		$service,	

		[Parameter(Position = 2, Mandatory = $true)]
		[string]
		$FirstName,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[string]
		$LastName,
		
		[Parameter(Position = 4, Mandatory = $true)]
		[string]
		$EmailAddress,
		
		[Parameter(Position = 5, Mandatory = $false)]
		[string]
		$CompanyName,
		
		[Parameter(Position = 6, Mandatory = $false)]
		[System.Management.Automation.PSCredential]
		$Credentials,
		
		[Parameter(Position = 7, Mandatory = $false)]
		[string]
		$Department,
		
		[Parameter(Position = 8, Mandatory = $false)]
		[string]
		$Office,
		
		[Parameter(Position = 9, Mandatory = $false)]
		[string]
		$BusinssPhone,
		
		[Parameter(Position = 10, Mandatory = $false)]
		[string]
		$MobilePhone,
		
		[Parameter(Position = 11, Mandatory = $false)]
		[string]
		$HomePhone,
		
		[Parameter(Position = 12, Mandatory = $false)]
		[string]
		$IMAddress,
		
		[Parameter(Position = 13, Mandatory = $false)]
		[string]
		$Street,
		
		[Parameter(Position = 14, Mandatory = $false)]
		[string]
		$City,
		
		[Parameter(Position = 15, Mandatory = $false)]
		[string]
		$State,
		
		[Parameter(Position = 16, Mandatory = $false)]
		[string]
		$PostalCode,
		
		[Parameter(Position = 17, Mandatory = $false)]
		[string]
		$Country,
		
		[Parameter(Position = 18, Mandatory = $false)]
		[string]
		$JobTitle,
		
		[Parameter(Position = 19, Mandatory = $false)]
		[string]
		$Notes,
		
		[Parameter(Position = 20, Mandatory = $false)]
		[string]
		$Photo,
		
		[Parameter(Position = 21, Mandatory = $false)]
		[string]
		$FileAs,
		
		[Parameter(Position = 22, Mandatory = $false)]
		[string]
		$WebSite,
		
		[Parameter(Position = 23, Mandatory = $false)]
		[string]
		$Title,
		
		[Parameter(Position = 24, Mandatory = $false)]
		[string]
		$EmailAddressDisplayAs		
		
		
	)
	Begin
	{
			$Contact = New-Object Microsoft.Exchange.WebServices.Data.Contact -ArgumentList $service
			#Set the GivenName
			$Contact.GivenName = $FirstName
			#Set the LastName
			$Contact.Surname = $LastName
			#Set Subject  
			$Contact.Subject = $DisplayName
			$Contact.FileAs = $DisplayName
			if ($Title -ne "")
			{
				$PR_DISPLAY_NAME_PREFIX_W = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3A45, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);
				$Contact.SetExtendedProperty($PR_DISPLAY_NAME_PREFIX_W, $Title)
			}
			$Contact.CompanyName = $CompanyName
			$Contact.DisplayName = $DisplayName
			$Contact.Department = $Department
			$Contact.OfficeLocation = $Office
			$Contact.CompanyName = $CompanyName
			$Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone] = $BusinssPhone
			$Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone] = $MobilePhone
			$Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::HomePhone] = $HomePhone
			$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business] = New-Object  Microsoft.Exchange.WebServices.Data.PhysicalAddressEntry
			$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].Street = $Street
			$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].State = $State
			$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].City = $City
			$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].CountryOrRegion = $Country
			$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].PostalCode = $PostalCode
			$Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1] = $EmailAddress
			if ([string]::IsNullOrEmpty($EmailAddressDisplayAs) -eq $false)
			{
				$Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Name = $EmailAddressDisplayAs
			}
			$Contact.ImAddresses[[Microsoft.Exchange.WebServices.Data.ImAddressKey]::ImAddress1] = $IMAddress
			$Contact.FileAs = $FileAs
			$Contact.BusinessHomePage = $WebSite
			#Set any Notes  
			$Contact.Body = $Notes
			$Contact.JobTitle = $JobTitle
			if ($Photo)
			{
				$fileAttach = $Contact.Attachments.AddFileAttachment($Photo)
				$fileAttach.IsContactPhoto = $true
			}
			return $Contact	
	}
}

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
        Invoke-LoadEWSManagedAPI
		
        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016     
		
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
		  
        Invoke-HandleSSL	
		  
        ## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use  

        #CAS URL Option 1 Autodiscover  
        if ($url) {
            $uri = [system.URI] $url
            $service.Url = $uri    
        }
        else {
            $service.AutodiscoverUrl($MailboxName, { $true })  
        }
        #Write-host ("Using CAS Server : " + $Service.url)   
		   
        #CAS URL Option 2 Hardcoded  
		  
        #$uri=[system.URI] "https://casservername/ews/exchange.asmx"  
        #$service.Url = $uri    
		  
        ## Optional section for Exchange Impersonation  
		  
        #$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
        if (!$service.Url) {
            throw "Error connecting to EWS"
        }
        else {	
            return, $service
        }
    }
}

function Invoke-LoadEWSManagedAPI {
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

function Invoke-HandleSSL {
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

function Convert-ToString($ipInputString) {  
    $Val1Text = ""  
    for ($clInt = 0; $clInt -lt $ipInputString.length; $clInt++) {  
        $Val1Text = $Val1Text + [Convert]::ToString([Convert]::ToChar([Convert]::ToInt32($ipInputString.Substring($clInt, 2), 16)))  
        $clInt++  
    }  
    return $Val1Text  
} 