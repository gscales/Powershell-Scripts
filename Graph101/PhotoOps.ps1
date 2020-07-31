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

function Get-GraphUserPhoto {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [String]
        $Filename,
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
        $PhotoSize,
        [Parameter(Position = 7, Mandatory = $false)]
        [int32]
        $ReSizeDimension,
        [Parameter(Position = 8, Mandatory = $false)]
        [String]
        $ReSizeImageForamt="Png"			
    )

    process {
        
        $prompt = $true
        if($AutoPrompt.IsPresent){
            $prompt = $false
        }
        $EndPoint = "https://graph.microsoft.com/v1.0/users"
        $RequestURL = $EndPoint + "('$MailboxName')/photo/`$value" 
        if(![String]::IsNullOrEmpty($PhotoSize)){
            $RequestURL = $EndPoint + "('$MailboxName')/photos/$PhotoSize/`$value" 
        }
        $AccessToken = Get-AccessTokenForGraph -MailboxName $Mailboxname -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -Prompt:$prompt
        $headers = @{
            'Authorization' = "Bearer $AccessToken"
            'AnchorMailbox' = "$MailboxName"
        }
        Invoke-WebRequest -Method Get -Uri $RequestURL -UserAgent "GraphBasicsPs101" -Headers $headers -OutFile $Filename     
        If($ReSizeDimension -gt 0){
              Add-Type -AssemblyName System.Drawing
              $OriginalPhotoFile = [System.Drawing.Image]::FromFile((Get-Item $Filename))
              $ResizedPhoto = New-Object System.Drawing.Bitmap($ReSizeDimension,$ReSizeDimension)
              $NewGraphic = [System.Drawing.Graphics]::FromImage($ResizedPhoto)
              $NewGraphic.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy;
              $NewGraphic.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality;
              $NewGraphic.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic;
              $NewGraphic.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality;
              $NewGraphic.PixelOffsetMode =  [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality;
              $NewGraphic.DrawImage($OriginalPhotoFile, 0, 0, $ReSizeDimension,$ReSizeDimension)
              $OriginalPhotoFile.Dispose()              
              $ResizedPhoto.Save($Filename,[System.Drawing.Imaging.ImageFormat]::$ReSizeImageForamt);              
        }
    }
}




