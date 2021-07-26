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
        $Scope = "https://graph.microsoft.com/.default"
         
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




function Invoke-CreateCategorySearchFolder {
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
        $AutoPrompt,
        [Parameter(Position = 6, Mandatory = $true)]
        [String]
        $SearchFolderName,	
        [Parameter(Position = 7, Mandatory = $true)]
        [String]
        $CategoryName,
        [Parameter(Position = 8, Mandatory = $false)]
        [switch]
        $ServicePrincipalAuthentication,
        [Parameter(Position = 9, Mandatory = $true)]
        [String]
        $CertificateThumbPrint,  
        [Parameter(Position =10, Mandatory = $true)]
        [string]
        $TenantDomain  		
    )

    process {
        
        $prompt = $true
        if($AutoPrompt.IsPresent){
            $prompt = $false
        }
        $EndPoint = "https://graph.microsoft.com/v1.0/users"
        $RequestURL = $EndPoint + "('$MailboxName')/MailFolders('SearchFolders')/childfolders"
        if($ServicePrincipalAuthentication.IsPresent){
            $Certificate = Get-Item ("Cert:\CurrentUser\My\$CertificateThumbPrint")
            $token = Get-AccessTokenForGraphFromCertificate -TenantDomain $TenantDomain -ClientId $ClientId -Certificate $Certificate
            $AccessToken = $token.access_token      
        }else{
            $AccessToken = Get-AccessTokenForGraph -MailboxName $Mailboxname -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -Prompt:$prompt
        }     
        $JsonBody = @"
{
    "@odata.type": "microsoft.graph.mailSearchFolder",
    "displayName": "$SearchFolderName",
    "includeNestedFolders": true,
    "sourceFolderIds": ["MsgFolderRoot"],
    "filterQuery": "Categories/any(a:a+eq+'$CategoryName')"
}
"@
        $headers = @{
            'Authorization' = "Bearer $AccessToken"
            'AnchorMailbox' = "$MailboxName"
        }
        return (Invoke-RestMethod -Method POST -Uri $RequestURL -UserAgent "GraphBasicsPs101" -Headers $headers -Body $JsonBody -ContentType "application/json" )  
 
    }
}

function Invoke-CreateSearchFolder {
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
        $AutoPrompt,
        [Parameter(Position = 6, Mandatory = $true)]
        [String]
        $SearchFolderName,	
        [Parameter(Position = 7, Mandatory = $true)]
        [String]
        $Filter,        
        [Parameter(Position = 8, Mandatory = $false)]
        [switch]
        $ServicePrincipalAuthentication,
        [Parameter(Position = 9, Mandatory = $true)]
        [String]
        $CertificateThumbPrint,  
        [Parameter(Position =10, Mandatory = $true)]
        [string]
        $TenantDomain  

    )

    process {
        
        $prompt = $true
        if($AutoPrompt.IsPresent){
            $prompt = $false
        }
        $EndPoint = "https://graph.microsoft.com/v1.0/users"
        $RequestURL = $EndPoint + "('$MailboxName')/MailFolders('SearchFolders')/childfolders"
        if($ServicePrincipalAuthentication.IsPresent){
            $Certificate = Get-Item ("Cert:\CurrentUser\My\$CertificateThumbPrint")
            $token = Get-AccessTokenForGraphFromCertificate -TenantDomain $TenantDomain -ClientId $ClientId -Certificate $Certificate
            $AccessToken = $token.access_token      
        }else{
            $AccessToken = Get-AccessTokenForGraph -MailboxName $Mailboxname -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -Prompt:$prompt
        }        
        $JsonBody = @"
{
    "@odata.type": "microsoft.graph.mailSearchFolder",
    "displayName": "$SearchFolderName",
    "includeNestedFolders": true,
    "sourceFolderIds": ["MsgFolderRoot"],
    "filterQuery": "$Filter"
}
"@
        $headers = @{
            'Authorization' = "Bearer $AccessToken"
            'AnchorMailbox' = "$MailboxName"
        }
        return (Invoke-RestMethod -Method POST -Uri $RequestURL -UserAgent "GraphBasicsPs101" -Headers $headers -Body $JsonBody -ContentType "application/json" )  
 
    }
}

function Get-SearchFolders {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $false)]
		[psobject]
        $AccessToken,   
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
        $AutoPrompt,
        [Parameter(Position = 6, Mandatory = $false)]
        [String]
        $Filter,
        [Parameter(Position = 7, Mandatory = $false)]
        [switch]
        $ServicePrincipalAuthentication,
        [Parameter(Position = 8, Mandatory = $true)]
        [String]
        $CertificateThumbPrint,  
        [Parameter(Position = 9, Mandatory = $true)]
        [string]
        $TenantDomain  	 

    )

    process {        
        $prompt = $true
        if($AutoPrompt.IsPresent){
            $prompt = $false
        }
        $EndPoint = "https://graph.microsoft.com/v1.0/users"
        $RequestURL = $EndPoint + "('$MailboxName')/MailFolders('SearchFolders')/childfolders?`$Top=999"
        if(![String]::IsNullOrEmpty($Filter)){
            $RequestURL += "&`$Filter=" + $Filter
        }
        if([String]::IsNullOrEmpty($AccessToken)){
            if($ServicePrincipalAuthentication.IsPresent){
                $Certificate = Get-Item ("Cert:\CurrentUser\My\$CertificateThumbPrint")
                $token = Get-AccessTokenForGraphFromCertificate -TenantDomain $TenantDomain -ClientId $ClientId -Certificate $Certificate
                $AccessToken = $token.access_token      
            }else{
                $AccessToken = Get-AccessTokenForGraph -MailboxName $Mailboxname -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -Prompt:$prompt
            }  
        }  
		do
		{
            $headers = @{
                'Authorization' = "Bearer $AccessToken"
                'AnchorMailbox' = "$MailboxName"
            }
            $Folders = (Invoke-RestMethod -Method Get -Uri $RequestURL -UserAgent "GraphBasicsPs101" -Headers $headers).value  
			foreach ($Folder in $Folders)
			{				
				Write-Output $Folder
    		}
			$RequestURL = $JSONOutput.'@odata.nextLink'
        }
        while (![String]::IsNullOrEmpty($RequestURL))          
 
    }
}

function Invoke-RemoveSearchFolder {
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
        $AutoPrompt,
        [Parameter(Position = 6, Mandatory = $true)]
        [String]
        $SearchFolderName	
    )

    process {
        if([String]::IsNullOrEmpty($SearchFolderName)){throw "Search FolderName can't be blank"}
        $prompt = $true
        if($AutoPrompt.IsPresent){
            $prompt = $false
        }
        $AccessToken = Get-AccessTokenForGraph -MailboxName $Mailboxname -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -Prompt:$prompt
        $Folder = Get-SearchFolders -MailboxName $MailboxName -Filter "displayName eq '$SearchFolderName'" -AccessToken $AccessToken
        $headers = @{
            'Authorization' = "Bearer $AccessToken"
            'AnchorMailbox' = "$MailboxName"
        }
        if($Folder.displayName -eq $SearchFolderName){
            $EndPoint = "https://graph.microsoft.com/v1.0/users"
            if(![String]::IsNullOrEmpty($Folder.id)){
                $RequestURL = $EndPoint + "('$MailboxName')/MailFolders/" + $Folder.id
                return (Invoke-RestMethod -Method Delete -Uri $RequestURL -UserAgent "GraphBasicsPs101" -Headers $headers)   
            }else{
                throw "Folder Id invalid"
            }
          
        }else{
            Write-Host "No Folder Found"
        }

    }
}

function Get-FolderFromPath {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string]
        $FolderPath,
		
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
        $AutoPrompt	
    )

    process {
        
        $prompt = $true
        if($AutoPrompt.IsPresent){
            $prompt = $false
        }
        $EndPoint = "https://graph.microsoft.com/v1.0/users"
        $RequestURL = $EndPoint + "('$MailboxName')/MailFolders('MsgFolderRoot')/childfolders?"
        $AccessToken = Get-AccessTokenForGraph -MailboxName $Mailboxname -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -Prompt:$prompt
        $fldArray = $FolderPath.Split("\")
        $PropList = @()
        $FolderSizeProp = Get-TaggedProperty -DataType "Long" -Id "0x66b3"
        $EntryId = Get-TaggedProperty -DataType "Binary" -Id "0xfff"
        $PropList += $FolderSizeProp 
        $PropList += $EntryId
        $Props = Get-ExtendedPropList -PropertyList $PropList 
        $RequestURL += "`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
        #Loop through the Split Array and do a Search for each level of folder 
        for ($lint = 1; $lint -lt $fldArray.Length; $lint++) {
            #Perform search based on the displayname of each folder level
            $FolderName = $fldArray[$lint];
            $headers = @{
                'Authorization' = "Bearer $AccessToken"
                'AnchorMailbox' = "$MailboxName"
            }
            $RequestURL = $RequestURL += "`&`$filter=DisplayName eq '$FolderName'"
            $tfTargetFolder = (Invoke-RestMethod -Method Get -Uri $RequestURL -UserAgent "GraphBasicsPs101" -Headers $headers).value  
            if ($tfTargetFolder.displayname -match $FolderName) {
                $folderId = $tfTargetFolder.Id.ToString()
                $RequestURL = $EndPoint + "('$MailboxName')/MailFolders('$folderId')/childfolders?"
                $RequestURL += "`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
            }
            else {
                throw ("Folder Not found")
            }
        }
        if ($tfTargetFolder.singleValueExtendedProperties) {
            foreach ($Prop in $tfTargetFolder.singleValueExtendedProperties) {
                Switch ($Prop.Id) {
                    "Long 0x66b3" {      
                        $tfTargetFolder | Add-Member -NotePropertyName "FolderSize" -NotePropertyValue $Prop.value 
                    }
                    "Binary 0xfff" {
                        $tfTargetFolder | Add-Member -NotePropertyName "PR_ENTRYID" -NotePropertyValue ([System.BitConverter]::ToString([Convert]::FromBase64String($Prop.Value)).Replace("-", ""))
                        $tfTargetFolder | Add-Member -NotePropertyName "ComplianceSearchId" -NotePropertyValue ("folderid:" + $tfTargetFolder.PR_ENTRYID.SubString(($tfTargetFolder.PR_ENTRYID.length - 48)))
                    }
                }
            }
        }
        return $tfTargetFolder 
    }
}

function Get-TaggedProperty {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [String]
        $DataType,
		
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $Id,
		
        [Parameter(Position = 2, Mandatory = $false)]
        [String]
        $Value
    )
    Begin {
        $Property = "" | Select-Object Id, DataType, PropertyType, Value
        $Property.Id = $Id
        $Property.DataType = $DataType
        $Property.PropertyType = "Tagged"
        if (![String]::IsNullOrEmpty($Value)) {
            $Property.Value = $Value
        }
        return, $Property
    }
}


function Get-ExtendedPropList {
    [CmdletBinding()]
    param (
        [Parameter(Position = 1, Mandatory = $false)]
        [PSCustomObject]
        $PropertyList
    )
    Begin {
        $rtString = "";
        $PropName = "Id"
        foreach ($Prop in $PropertyList) {
            if ($Prop.PropertyType -eq "Tagged") {
                if ($rtString -eq "") {
                    $rtString = "($PropName%20eq%20'" + $Prop.DataType + "%20" + $Prop.Id + "')"
                }
                else {
                    $rtString += " or ($PropName%20eq%20'" + $Prop.DataType + "%20" + $Prop.Id + "')"
                }
            }
            else {
                if ($Prop.Type -eq "String") {
                    if ($rtString -eq "") {
                        $rtString = "($PropName%20eq%20'" + $Prop.DataType + "%20{" + $Prop.Guid + "}%20Name%20" + $Prop.Id + "')"
                    }
                    else {
                        $rtString += " or ($PropName%20eq%20'" + $Prop.DataType + "%20{" + $Prop.Guid + "}%20Name%20" + $Prop.Id + "')"
                    }
                }
                else {
                    if ($rtString -eq "") {
                        $rtString = "($PropName%20eq%20'" + $Prop.DataType + "%20{" + $Prop.Guid + "}%20Id%20" + $Prop.Id + "')"
                    }
                    else {
                        $rtString += " or ($PropName%20eq%20'" + $Prop.DataType + "%20{" + $Prop.Guid + "}%20Id%20" + $Prop.Id + "')"
                    }
                }
            }
			
        }
        return $rtString
		
    }
}


