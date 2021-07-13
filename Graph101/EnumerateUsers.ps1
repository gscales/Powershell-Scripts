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

function Get-AzureUsersFromGraph{
    [CmdletBinding()]
    param (
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $MailboxName,
        [Parameter(Position = 3, Mandatory = $false)]
        [String]
        $ClientId,
        [Parameter(Position = 4, Mandatory = $false)]
        [String]
        $RedirectURI = "urn:ietf:wg:oauth:2.0:oob",
        [Parameter(Position = 5, Mandatory = $false)]
        [String]
        $scopes = "User.Read.All Mail.Read",
        [Parameter(Position = 6, Mandatory = $false)]
        [switch]
        $AutoPrompt,	
        [Parameter(Position = 7, Mandatory = $false)]
        [String]
        $AccessToken,
        [Parameter(Position = 8, Mandatory = $false)]
        [String]
        $filter,
        [Parameter(Position = 9, Mandatory = $false)]
        [String]
        $SelectList,
        [Parameter(Position = 10, Mandatory = $false)]
        [switch]
        $AdvancedQuery

    )

    process {
        $prompt = $true
        if($AutoPrompt.IsPresent){
            $prompt = $false
        }
        if([String]::IsNullOrEmpty($AccessToken)){
            $AccessToken = Get-AccessTokenForGraph -MailboxName $Mailboxname -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -Prompt:$prompt
        }     
        $top=999
        $EndPoint = "https://graph.microsoft.com/v1.0/users?`$top=$top"
        $RequestURL = $EndPoint 
        if(![String]::IsNullOrEmpty($SelectList)){
            $RequestURL = $EndPoint + "&`$Select=" + $SelectList
        }  
        if(![String]::IsNullOrEmpty($filter)){
            $RequestURL += "&`$filter=" + $filter
        } 
        if($AdvancedQuery.IsPresent){
            $RequestURL += "&`$Count=true"
            $headers = @{
                'Authorization' = "Bearer $AccessToken"
                'AnchorMailbox' = "$MailboxName"
                'ConsistencyLevel' = "eventual"
            }
        }else{
            $headers = @{
                'Authorization' = "Bearer $AccessToken"
                'AnchorMailbox' = "$MailboxName"
            } 
        } 
        do {
            $Results = (Invoke-RestMethod -Method Get -Uri $RequestURL -UserAgent "GraphBasicsPs101" -Headers $headers)
            $RequestURL  = $null
            if($Results){
                if ($Results.value) {
                    $QueryResults = $Results.value
                } else {
                    $QueryResults = $Results
                }
                foreach($Item in $QueryResults){        
                    Write-Output $Item                          
                }
                $QueryResults = $null         
                $RequestURL = $Results.'@odata.nextlink'      
            }
        } until (!($RequestURL))        
        
    }
}
