function Get-AppSettings(){
        param( 
        
        )  
 	Begin
		 {
            $configObj = "" |  select ResourceURL,ClientId,redirectUrl
            $configObj.ResourceURL = "outlook.office.com"
            $configObj.ClientId = "5471030d-f311-4c5d-91ef-74ca885463a7"
            $configObj.redirectUrl = "urn:ietf:wg:oauth:2.0:oob"
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
            $redirectUrl = [System.Web.HttpUtility]::UrlEncode($AppSetting.redirectUrl)
            $Phase1auth = Show-OAuthWindow -Url "https://login.microsoftonline.com/common/oauth2/authorize?resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&response_type=code&redirect_uri=$redirectUrl&prompt=login"
            $code = $Phase1auth["code"]
            $AuthorizationPostRequest = "resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&grant_type=authorization_code&code=$code&redirect_uri=$redirectUrl"
            $content = New-Object System.Net.Http.StringContent($AuthorizationPostRequest, [System.Text.Encoding]::UTF8, "application/x-www-form-urlencoded")
            $ClientReesult = $HttpClient.PostAsync([Uri]("https://login.windows.net/common/oauth2/token"),$content)
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
            $ClientReesult = $HttpClient.PostAsync([Uri]("https://login.windows.net/common/oauth2/token"),$content)
            $JsonObject = ConvertFrom-Json -InputObject  $ClientReesult.Result.Content.ReadAsStringAsync().Result
            return $JsonObject
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
             $ClientReesult = $HttpClient.GetAsync($RequestURL)
             $JsonObject = ConvertFrom-Json -InputObject  $ClientReesult.Result.Content.ReadAsStringAsync().Result
             return $JsonObject     
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



