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

function Get-AutoReplySettings{
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

        [Parameter(Position =4,Mandatory = $true)]
		[string]
        $TargetUser
    )
    Process{
        $Certificate = Get-Item ("Cert:\CurrentUser\My\$CertificateThumbPrint")
        $token = Get-AccessTokenForGraphFromCertificate -TenantDomain $TenantDomain -ClientId $ClientId -Certificate $Certificate
        $AccessToken = $token.access_token
        $headers = @{
            'Authorization' = "Bearer $AccessToken"
        }
        $RequestURL = "https://graph.microsoft.com/v1.0/users/$TargetUser/mailboxSettings/automaticRepliesSetting"
        return (Invoke-RestMethod -Method Get -Uri $RequestURL -UserAgent "GraphBasicsPs101" -Headers $headers)  
        
    }
}

function Get-BatchAutoReplySettings{
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

        [Parameter(Position =4,Mandatory = $true)]
		[psobject]
        $TargetUsers
    )
    Process{
        $Certificate = Get-Item ("Cert:\CurrentUser\My\$CertificateThumbPrint")
        $token = Get-AccessTokenForGraphFromCertificate -TenantDomain $TenantDomain -ClientId $ClientId -Certificate $Certificate
        $AccessToken = $token.access_token
        $headers = @{
            'Authorization' = "Bearer $AccessToken"
        }
        $batchCount = 1
        $RequestIndex = @{}
        $BatchRequestContent = ""
        foreach($TargetUser in $TargetUsers){
            $RequestIndex.Add($batchCount,$TargetUser)
            $BatchRequestContent += "{`"id`": `"" + $batchCount + "`",`"method`": `"GET`"," 
            $BatchRequestContent += "`"url`": `"/users/$TargetUser/mailboxSettings/automaticRepliesSetting`"},`r`n" 
            $batchCount++
        }
        $RequestURL = "https://graph.microsoft.com/v1.0/`$batch"
        $RequestContent = "{`r`n`"requests`": ["
        $RequestContent += $BatchRequestContent
        $RequestContent += "]}"
        $BatchResponse = (Invoke-RestMethod -Method POST -Uri $RequestURL -UserAgent "GraphBasicsPs101" -Headers $headers -Body $RequestContent -ContentType "application/json" )   
        if($BatchResponse.responses){
            $repCount = 0;
            foreach($Response in $BatchResponse.responses){
                $ResponseUser = $RequestIndex[[Int32]$Response.id]
                Write-Verbose($Response.id)
                Write-Verbose($Response.status)
                if([Int32]$Response.status -eq 200){
                   $outputObj = "" | Select User,automaticRepliesSetting,Error
                   $outputObj.User = $ResponseUser
                   $outputObj.automaticRepliesSetting = $Response.body
                   Write-Output $outputObj
                }else{
                    if([Int32]$Response.status -eq 429){
                        $rptObject.ThrottleCount++
                        if(!$TimeOutServed){
                            Write-Verbose($Response.Headers.'Retry-After')
                            Write-Verbose("Serving Throttling Timeout " + $Response.Headers.'Retry-After')
                            Start-sleep -Seconds $Response.Headers.'Retry-After'
                            $TimeOutServed = $true
                        }
                    }else{
                        $outputObj = "" | Select User,automaticRepliesSetting,Error
                        $outputObj.User = $ResponseUser
                        $outputObj.Error = $Response.body;
                        Write-Output $outputObj
                    }

                }
                $repCount++
            }
        }
        
    }
}

function Get-AutoReplySettingsMailtips{
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

        [Parameter(Position =4,Mandatory = $true)]
		[string[]]
        $TargetUsers
    )
    Process{
        $Certificate = Get-Item ("Cert:\CurrentUser\My\$CertificateThumbPrint")
        $token = Get-AccessTokenForGraphFromCertificate -TenantDomain $TenantDomain -ClientId $ClientId -Certificate $Certificate
        $AccessToken = $token.access_token
        $headers = @{
            'Authorization' = "Bearer $AccessToken"
        }
        $GraphPost = Convertto-Json @{EmailAddresses=$TargetUsers;
            MailTipsOptions= "automaticReplies"
        }
        $user = $TargetUsers[0]
        $RequestURL = "https://graph.microsoft.com/v1.0/users/$user/getMailTips"
        return (Invoke-RestMethod -Method POST -Uri $RequestURL -UserAgent "GraphBasicsPs101" -Headers $headers -Body $GraphPost -ContentType "application/json").value  
        
    }
}