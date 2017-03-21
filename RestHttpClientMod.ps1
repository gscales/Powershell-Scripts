function Get-AppSettings(){
        param( 
        
        )  
 	Begin
		 {
            $configObj = "" |  select ResourceURL,ClientId,redirectUrl,ClientSecret,x5t,TenantId,ValidateForMinutes
            $configObj.ResourceURL = "outlook.office.com"
            $configObj.ClientId = "65bf30bb-8a4f-437e-9a72-8bb00d2edf6c"
            $configObj.redirectUrl = "http://192.168.0.1"
            $configObj.TenantId = "1c3a18bf-da31-4f6c-a404-2c06c9cf5ae4"
            $configObj.ClientSecret = ""
            $configObj.x5t = "z+bmTCHeNVR7TC4IG8dW/LBgXGk="
            $configObj.ValidateForMinutes = 60
            return $configObj            
         }    
}

function Get-HTTPClient{ 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName
    )  
 	Begin
		 {
            Add-Type -AssemblyName System.Net.Http
            $handler = New-Object  System.Net.Http.HttpClientHandler
            $handler.CookieContainer = New-Object System.Net.CookieContainer
            $handler.AllowAutoRedirect = $true;
            $HttpClient = New-Object System.Net.Http.HttpClient($handler);
            #$HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", "");
            $Header = New-Object System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json")
            $HttpClient.DefaultRequestHeaders.Accept.Add($Header);
            $HttpClient.Timeout = New-Object System.TimeSpan(0, 0, 90);
            $HttpClient.DefaultRequestHeaders.TransferEncodingChunked = $false
            if (!$HttpClient.DefaultRequestHeaders.Contains("X-AnchorMailbox")){
                $HttpClient.DefaultRequestHeaders.Add("X-AnchorMailbox", $MailboxName);
            }
            $Header = New-Object System.Net.Http.Headers.ProductInfoHeaderValue("RestClient", "1.1")
            $HttpClient.DefaultRequestHeaders.UserAgent.Add($Header);
            return $HttpClient
         }
}

function Convert-FromBase64StringWithNoPadding([string]$data)
{
    $data = $data.Replace('-', '+').Replace('_', '/')
    switch ($data.Length % 4)
    {
        0 { break }
        2 { $data += '==' }
        3 { $data += '=' }
        default { throw New-Object ArgumentException('data') }
    }
    return [System.Convert]::FromBase64String($data)
}

function Decode-Token { 
        param( 
        [Parameter(Position=1, Mandatory=$true)] [String]$Token
    )  
    ## Start Code Attribution
    ## Decode-Token function is based on work of the following Authors and should remain with the function if copied into other scripts
    ## https://gallery.technet.microsoft.com/JWT-Token-Decode-637cf001
    ## End Code Attribution
    Begin
    {
        $parts = $Token.Split('.');
        $headers = [System.Text.Encoding]::UTF8.GetString((Convert-FromBase64StringWithNoPadding $parts[0]))
        $claims = [System.Text.Encoding]::UTF8.GetString((Convert-FromBase64StringWithNoPadding $parts[1]))
        $signature = (Convert-FromBase64StringWithNoPadding $parts[2])

        $customObject = [PSCustomObject]@{
            headers = ($headers | ConvertFrom-Json)
            claims = ($claims | ConvertFrom-Json)
            signature = $signature
        }
        return $customObject
    }
}

function New-JWTToken{
        param( 
        [Parameter(Position=1, Mandatory=$true)] [string]$CertFileName,
        [Parameter(Mandatory=$True)][Security.SecureString]$password        
    )  
 	Begin
		 {
            $configObj = Get-AppSettings 
            $date1 = Get-Date -Date "01/01/1970"
            $date2 = (Get-Date).ToUniversalTime().AddMinutes($configObj.ValidateForMinutes)           
            $date3 = (Get-Date).ToUniversalTime().AddMinutes(-5)      
            $exp = [Math]::Round((New-TimeSpan -Start $date1 -End $date2).TotalSeconds,0) 
            $nbf = [Math]::Round((New-TimeSpan -Start $date1 -End $date3).TotalSeconds,0) 
            $exVal = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable
            $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $CertFileName,$password,$exVal
            $jti = [System.Guid]::NewGuid().ToString()
            $Headerassertaion =  "{" 
            $Headerassertaion += "     `"alg`": `"RS256`"," 
            $Headerassertaion += "     `"x5t`": `""+ $configObj.x5t + "`""
            $Headerassertaion += "}"
            $PayLoadassertaion += "{"
            $PayLoadassertaion += "    `"aud`": `"https://login.windows.net/" + $configObj.TenantId +"/oauth2/token`"," 
            $PayLoadassertaion += "    `"exp`": $exp,"            
            $PayLoadassertaion += "    `"iss`": `""+ $configObj.ClientId + "`"," 
            $PayLoadassertaion += "    `"jti`": `"" + $jti + "`","
            $PayLoadassertaion += "    `"nbf`": $nbf,"       
            $PayLoadassertaion += "    `"sub`": `"" + $configObj.ClientId + "`""              
            $PayLoadassertaion += "} " 
            $encodedHeader = [System.Convert]::ToBase64String([System.Text.UTF8Encoding]::UTF8.GetBytes($Headerassertaion)).Replace('=','').Replace('+', '-').Replace('/', '_')
            $encodedPayLoadassertaion = [System.Convert]::ToBase64String([System.Text.UTF8Encoding]::UTF8.GetBytes($PayLoadassertaion)).Replace('=','').Replace('+', '-').Replace('/', '_')
            $JWTOutput = $encodedHeader + "." + $encodedPayLoadassertaion
            $SigBytes = [System.Text.UTF8Encoding]::UTF8.GetBytes($JWTOutput)            
            $rsa = $cert.PrivateKey;
            $sha256 = [System.Security.Cryptography.SHA256]::Create()
            $hash = $sha256.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($encodedHeader + '.' + $encodedPayLoadassertaion));
            $sigform = New-Object System.Security.Cryptography.RSAPKCS1SignatureFormatter($rsa);
            $sigform.SetHashAlgorithm("SHA256");
            $sig = [System.Convert]::ToBase64String($sigform.CreateSignature($hash)).Replace('=','').Replace('+', '-').Replace('/', '_')
            $JWTOutput = $encodedHeader + '.' + $encodedPayLoadassertaion + '.' + $sig
            Write-Output ($JWTOutput)

         }
}

function Invoke-CreateSelfSignedCert{
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$CertName,
        [Parameter(Position=1, Mandatory=$true)] [string]$CertFileName,
        [Parameter(Position=2, Mandatory=$true)] [string]$KeyFileName
    )  
 	Begin
		 {
             $Cert = New-SelfSignedCertificate -certstorelocation cert:\currentuser\my -dnsname $CertName -Provider 'Microsoft Enhanced RSA and AES Cryptographic Provider' 
             $SecurePassword = Read-Host -Prompt "Enter password" -AsSecureString
             $CertPath = "cert:\currentuser\my\" + $Cert.Thumbprint.ToString()
             Export-PfxCertificate -cert $CertPath -FilePath $CertFileName -Password $SecurePassword 
             $bin = $cert.RawData
             $base64Value = [System.Convert]::ToBase64String($bin)
             $bin = $cert.GetCertHash()
             $base64Thumbprint = [System.Convert]::ToBase64String($bin)
             $keyid = [System.Guid]::NewGuid().ToString()
             $jsonObj = @{customKeyIdentifier=$base64Thumbprint;keyId=$keyid;type="AsymmetricX509Cert";usage="Verify";value=$base64Value}
             $keyCredentials=ConvertTo-Json @($jsonObj) | Out-File $KeyFileName
             Remove-Item $CertPath
             Write-Host ("Key written to " + $KeyFileName)
             
         }
    
}

Function Show-OAuthWindow
{
    param(
        [System.Uri]$Url
    )
    ## Start Code Attribution
    ## Show-AuthWindow function is the work of the following Authors and should remain with the function if copied into other scripts
    ## https://foxdeploy.com/2015/11/02/using-powershell-and-oauth/
    ## https://blogs.technet.microsoft.com/ronba/2016/05/09/using-powershell-and-the-office-365-rest-api-with-oauth/
    ## End Code Attribution
    Add-Type -AssemblyName System.Web
    Add-Type -AssemblyName System.Windows.Forms
 
    $form = New-Object -TypeName System.Windows.Forms.Form -Property @{Width=440;Height=640}
    $web  = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width=420;Height=600;Url=($url ) }
    $DocComp  = {
        $Global:uri = $web.Url.AbsoluteUri
        if ($Global:Uri -match "error=[^&]*|code=[^&]*") {$form.Close() }
    }
    $web.ScriptErrorsSuppressed = $true
    $web.Add_DocumentCompleted($DocComp)
    $form.Controls.Add($web)
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() | Out-Null
    $queryOutput = [System.Web.HttpUtility]::ParseQueryString($web.Url.Query)
    $output = @{}
    foreach($key in $queryOutput.Keys){
        $output["$key"] = $queryOutput[$key]
    }
    return $output 
}

function Get-AccessToken{ 
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName
    )  
 	Begin
		 {
            Add-Type -AssemblyName System.Web
            $HttpClient =  Get-HTTPClient($MailboxName)
            $AppSetting = Get-AppSettings 
            $ResourceURL = $AppSetting.ResourceURL
            $ClientId = $AppSetting.ClientId
            $ClientSecret = $AppSetting.ClientSecret
            $redirectUrl = [System.Web.HttpUtility]::UrlEncode($AppSetting.redirectUrl)
            $Phase1auth = Show-OAuthWindow -Url "https://login.microsoftonline.com/common/oauth2/authorize?resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&response_type=code&redirect_uri=$redirectUrl&prompt=login"
            $code = $Phase1auth["code"]
            $AuthorizationPostRequest = "resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&grant_type=authorization_code&code=$code&redirect_uri=$redirectUrl"
            if(![String]::IsNullOrEmpty($ClientSecret)){
                $AuthorizationPostRequest = "resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&client_secret=$ClientSecret&grant_type=authorization_code&code=$code&redirect_uri=$redirectUrl"
            }
            $content = New-Object System.Net.Http.StringContent($AuthorizationPostRequest, [System.Text.Encoding]::UTF8, "application/x-www-form-urlencoded")
            $ClientReesult = $HttpClient.PostAsync([Uri]("https://login.windows.net/common/oauth2/token"),$content)
            $JsonObject = ConvertFrom-Json -InputObject  $ClientReesult.Result.Content.ReadAsStringAsync().Result
            return $JsonObject
         }
}

function Get-AppOnlyToken{ 
    param( 
       
        [Parameter(Position=1, Mandatory=$true)] [string]$CertFileName,
        [Parameter(Mandatory=$True)][Security.SecureString]$password   
    )  
 	Begin
		 {
            $JWTToken = New-JWTToken -CertFileName $CertFileName -password $password
            Add-Type -AssemblyName System.Web
            $HttpClient =  Get-HTTPClient(" ")
            $AppSetting = Get-AppSettings 
            $ResourceURL = $AppSetting.ResourceURL
            $ClientId = $AppSetting.ClientId
            $AuthorizationPostRequest = "resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&client_assertion_type=urn%3Aietf%3Aparams%3Aoauth%3Aclient-assertion-type%3Ajwt-bearer&client_assertion=$JWTToken&grant_type=client_credentials&redirect_uri=$redirectUrl"
            $content = New-Object System.Net.Http.StringContent($AuthorizationPostRequest, [System.Text.Encoding]::UTF8, "application/x-www-form-urlencoded")
            $ClientReesult = $HttpClient.PostAsync([Uri]("https://login.windows.net/" + $AppSetting.TenantId + "/oauth2/v2.0/token"),$content)
            $JsonObject = ConvertFrom-Json -InputObject  $ClientReesult.Result.Content.ReadAsStringAsync().Result
            return $JsonObject
         }
}



function Refresh-AccessToken{ 
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$true)] [string]$RefreshToken
    )  
 	Begin
		 {
            Add-Type -AssemblyName System.Web
            $HttpClient =  Get-HTTPClient($MailboxName)
            $AppSetting = Get-AppSettings 
            $ResourceURL = $AppSetting.ResourceURL
            $ClientId = $AppSetting.ClientId
            $redirectUrl = [System.Web.HttpUtility]::UrlEncode($AppSetting.redirectUrl)
            $AuthorizationPostRequest = "client_id=$ClientId&refresh_token=$RefreshToken&grant_type=refresh_token&redirect_uri=$redirectUrl"
            $content = New-Object System.Net.Http.StringContent($AuthorizationPostRequest, [System.Text.Encoding]::UTF8, "application/x-www-form-urlencoded")
            $ClientResult = $HttpClient.PostAsync([Uri]("https://login.windows.net/common/oauth2/token"),$content)             
             if (!$ClientResult.Result.IsSuccessStatusCode)
             {                    
                     Write-Output ("Error making REST POST " + $ClientResult.Result.StatusCode + " : " + $ClientResult.Result.ReasonPhrase)
                     Write-Output $ClientResult.Result
                     if($ClientResult.Content -ne $null){
                         Write-Output ($ClientResult.Content.ReadAsStringAsync().Result);   
                     }                     
             }
            else
             {
               $JsonObject = ConvertFrom-Json -InputObject  $ClientResult.Result.Content.ReadAsStringAsync().Result
               return $JsonObject
             }

         }
}

function Invoke-RestGet
{
        param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$RequestURL,
        [Parameter(Position=1, Mandatory=$true)] [String]$MailboxName,
        [Parameter(Position=2, Mandatory=$true)] [System.Net.Http.HttpClient]$HttpClient,
        [Parameter(Position=3, Mandatory=$true)] [PSCustomObject]$AccessToken
    )  
 	Begin
		 {
             #Check for expired Token
             $minTime = new-object DateTime(1970, 1, 1, 0, 0, 0, 0,[System.DateTimeKind]::Utc);
             $expiry =  $minTime.AddSeconds($AccessToken.expires_on)
             if($expiry -le [DateTime]::Now.ToUniversalTime()){
                write-host "Refresh Token"
                $AccessToken = Refresh-AccessToken -MailboxName $MailboxName -RefreshToken $AccessToken.refresh_token               
                Set-Variable -Name "AccessToken" -Value $AccessToken -Scope Script -Visibility Public
             }
             $HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", $AccessToken.access_token);
             $ClientResult = $HttpClient.GetAsync($RequestURL)
             if (!$ClientResult.Result.IsSuccessStatusCode)
             {
                     Write-Output ("Error making REST Get " + $ClientResult.Result.StatusCode + " : " + $ClientResult.Result.ReasonPhrase)
                     Write-Output $ClientResult.Result
                     if($ClientResult.Content -ne $null){
                         Write-Output ($ClientResult.Content.ReadAsStringAsync().Result);   
                     }                     
             }
            else
             {
               $JsonObject = ConvertFrom-Json -InputObject  $ClientResult.Result.Content.ReadAsStringAsync().Result
               return $JsonObject
             }
  
         }    
}

function Get-MailboxSettings{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }   
        $HttpClient =  Get-HTTPClient($MailboxName)
        $RequestURL =  "https://outlook.office.com/api/v2.0/Users('$MailboxName')/MailboxSettings"
        return Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
    }
}

function Get-AutomaticRepliesSettings{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName  
                    
        }   
        $HttpClient =  Get-HTTPClient($MailboxName)
        $RequestURL =  "https://outlook.office.com/api/v2.0/Users('$MailboxName')/MailboxSettings/AutomaticRepliesSetting"
       return Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
    }
}



function Get-MailboxTimeZone{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }   
        $HttpClient =  Get-HTTPClient($MailboxName)
        $RequestURL =  "https://outlook.office.com/api/v2.0/Users('$MailboxName')/MailboxSettings/TimeZone"
        return Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
    }
}

function Get-Folders{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }   
        $HttpClient =  Get-HTTPClient($MailboxName)
        $RequestURL =  "https://outlook.office.com/api/v2.0/Users('$MailboxName')/MailFolders/msgfolderroot/childfolders"
        return Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
    }
}

function Get-Inbox{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }   
        $HttpClient =  Get-HTTPClient($MailboxName)
        $RequestURL =  "https://outlook.office.com/api/v2.0/Users('$MailboxName')/MailFolders/Inbox"
        return Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
    }
}
function Get-ArchiveFolder{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }        
        $HttpClient =  Get-HTTPClient($MailboxName)
        $RequestURL =  "https://outlook.office.com/api/v2.0/Users('$MailboxName')/MailboxSettings/ArchiveFolder"
        $JsonObject =  Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
        $folderId = $JsonObject.value.ToString()
        $HttpClient =  Get-HTTPClient($MailboxName)
        $RequestURL =  "https://outlook.office.com/api/v2.0/Users('$MailboxName')/MailFolders('$folderId')"
        return  Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
    }
}

function Get-MailboxSettingsReport{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [psobject]$Mailboxes,
        [Parameter(Position=1, Mandatory=$true)] [string]$CertFileName,
        [Parameter(Mandatory=$True)][Security.SecureString]$password   
    )
    Begin{
        $rptCollection = @()
        $AccessToken = Get-AppOnlyToken -CertFileName $CertFileName -password $password 
        $HttpClient =  Get-HTTPClient($Mailboxes[0])
        foreach ($MailboxName in $Mailboxes) {
            $rptObj = "" | Select MailboxName,Language,Locale,TimeZone,AutomaticReplyStatus
            $RequestURL =  "https://outlook.office.com/api/v2.0/Users('$MailboxName')/MailboxSettings"
            $Results = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
            $rptObj.MailboxName = $MailboxName
            $rptObj.Language = $Results.Language.DisplayName
            $rptObj.Locale = $Results.Language.Locale
            $rptObj.TimeZone = $Results.TimeZone
            $rptObj.AutomaticReplyStatus = $Results.AutomaticRepliesSetting.Status
            $rptCollection += $rptObj
        }
        Write-Output  $rptCollection
        
    }
}

function  Get-People {
    param(
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken   
    )
    Begin{
        
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }        
        $HttpClient =  Get-HTTPClient($MailboxName)
        $RequestURL =  "https://outlook.office.com/api/beta/me/people/?`$top=1000&`$Select=DisplayName"
        $JsonObject =  Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
        Write-Output $JsonObject 
    }
}


