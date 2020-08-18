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

function Get-AccessTokenForAzure {
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
        $scopes = "https://outlook.office.com/SMTP.Send",
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

function Get-SMTPTLSCert {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $ServerName,
        [Parameter(Mandatory = $true)]
        [String]
        $Sendingdomain,
        [Parameter(Mandatory = $true)]
        [String]
        $CertificateFilePath,
        [int]
        $Port = 587
    )
    Process {
        $socket = new-object System.Net.Sockets.TcpClient($ServerName, $Port)
        $stream = $socket.GetStream()
        $streamWriter = new-object System.IO.StreamWriter($stream)
        $streamReader = new-object System.IO.StreamReader($stream)
        $stream.ReadTimeout = 5000
        $stream.WriteTimeout = 5000   
        $streamWriter.AutoFlush = $true
        $sslStream = New-Object System.Net.Security.SslStream($stream)    
        $sslStream.ReadTimeout = 5000
        $sslStream.WriteTimeout = 5000        
        $ConnectResponse = $streamReader.ReadLine();
        Write-Host($ConnectResponse)
        if(!$ConnectResponse.StartsWith("220")){
            throw "Error connecting to the SMTP Server"
        }
        Write-Host(("helo " + $FromHeloString)) -ForegroundColor Green
        $streamWriter.WriteLine(("helo " + $Sendingdomain));
        $ehloResponse = $streamReader.ReadLine();
        Write-Host($ehloResponse)
        if (!$ehloResponse.StartsWith("250")){
            throw "Error in ehelo Response"
        }
        Write-Host("STARTTLS") -ForegroundColor Green
        $streamWriter.WriteLine("STARTTLS");
        $startTLSResponse = $streamReader.ReadLine();
        Write-Host($startTLSResponse)
        $ccCol = New-Object System.Security.Cryptography.X509Certificates.X509CertificateCollection
        $sslStream.AuthenticateAsClient($ServerName,$ccCol,[System.Security.Authentication.SslProtocols]::Tls12,$false);        
        $Cert = $sslStream.RemoteCertificate.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert);
        [System.IO.File]::WriteAllBytes($CertificateFilePath, $Cert);
        $stream.Dispose()
        $sslStream.Dispose()
        Write-Host("File written to " + $CertificateFilePath)
        Write-Host("Done")
    }
}

function Invoke-TestSMTPTLSwithOAuth {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $ServerName,
        [Parameter(Mandatory = $true)]
        [String]
        $SendingAddress,
        [Parameter(Mandatory = $true)]
        [String]
        $To,
        [int]
        $Port = 587,
        [Parameter(Mandatory = $false)]
        [String]
        $ClientId = "20773535-6b8f-4f3d-8f0e-4b7710d79afe",
        [Parameter(Mandatory = $false)]
        [String]
        $RedirectURI = "msal20773535-6b8f-4f3d-8f0e-4b7710d79afe://auth"
    )
    Process {

        $socket = new-object System.Net.Sockets.TcpClient($ServerName, $Port)
        $stream = $socket.GetStream()
        $streamWriter = new-object System.IO.StreamWriter($stream)
        $streamReader = new-object System.IO.StreamReader($stream)
        $streamWriter.AutoFlush = $true
        $sslStream = New-Object System.Net.Security.SslStream($stream)
        $sslStream.ReadTimeout = 25000
        $sslStream.WriteTimeout = 25000        
        $ConnectResponse = $streamReader.ReadLine();
        Write-Host($ConnectResponse)
        if(!$ConnectResponse.StartsWith("220")){
            throw "Error connecting to the SMTP Server"
        }
        $Domain = $SendingAddress.Split('@')[1]
        Write-Host(("helo " + $Domain)) -ForegroundColor Green
        $streamWriter.WriteLine(("helo " + $Domain));
        $ehloResponse = $streamReader.ReadLine();
        Write-Host($ehloResponse)
        if (!$ehloResponse.StartsWith("250")){
            throw "Error in ehelo Response"
        }
        Write-Host("STARTTLS") -ForegroundColor Green
        $streamWriter.WriteLine("STARTTLS");
        $startTLSResponse = $streamReader.ReadLine();
        Write-Host($startTLSResponse)
        $ccCol = New-Object System.Security.Cryptography.X509Certificates.X509CertificateCollection
        $sslStream.AuthenticateAsClient($ServerName,$ccCol,[System.Security.Authentication.SslProtocols]::Tls12,$false);        
        $SSLstreamReader = new-object System.IO.StreamReader($sslStream)
        $SSLstreamWriter = new-object System.IO.StreamWriter($sslStream)
        $SSLstreamWriter.AutoFlush = $true
        $SSLstreamWriter.WriteLine(("helo " + $Domain));
        $ehloResponse = $SSLstreamReader.ReadLine();
        Write-Host($ehloResponse)
        $command = "AUTH XOAUTH2" 
        write-host -foregroundcolor DarkGreen $command
        $SSLstreamWriter.WriteLine($command) 
        $AuthLoginResponse = $SSLstreamReader.ReadLine()
        write-host ($AuthLoginResponse)
        $token = Get-AccessTokenForAzure -MailboxName $SendingAddress -ClientId $ClientId -RedirectURI $RedirectURI
        $Bytes = [System.Text.Encoding]::ASCII.GetBytes(("user=" + $SendingAddress + [char]1 + "auth=Bearer " + $token + [char]1 + [char]1))
        $Base64AuthSALS = [Convert]::ToBase64String($Bytes)     
        write-host -foregroundcolor DarkGreen $Base64AuthSALS
        $SSLstreamWriter.WriteLine($Base64AuthSALS)        
        $AuthResponse = $SSLstreamReader.ReadLine()
        write-host $AuthResponse
        if($AuthResponse.StartsWith("235")){
            $command = "MAIL FROM: <" + $SendingAddress + ">" 
            write-host -foregroundcolor DarkGreen $command
            $SSLstreamWriter.WriteLine($command) 
            $FromResponse = $SSLstreamReader.ReadLine()
            write-host $FromResponse
            $command = "RCPT TO: <" + $To + ">" 
            write-host -foregroundcolor DarkGreen $command
            $SSLstreamWriter.WriteLine($command) 
            $ToResponse = $SSLstreamReader.ReadLine()
            write-host $ToResponse
            $command = "QUIT" 
            write-host -foregroundcolor DarkGreen $command
            $SSLstreamWriter.WriteLine($command) 
            # ## Close the streams 
            $stream.Close() 
            $sslStream.Close()
            Write-Host("Done")
        }  

    }
}

function readResponse() {
    while($stream.DataAvailable)
    {
        $buffer = new-object System.Byte[] 1024
        $read = $stream.Read($buffer, 0, 1024)
        $rstring = $encoding.GetString($buffer, 0, $read)
        Write-Host $rstring
    }
}



function Invoke-SendMessagewithOAuth{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $ServerName,
        [Parameter(Mandatory = $true)]
        [String]
        $SendingAddress,
        [Parameter(Mandatory = $true)]
        [String]
        $userName,
        [Parameter(Mandatory = $true)]
        [String]
        $To,
        [Parameter(Mandatory = $true)]
        [String]
        $Subject,
        [Parameter(Mandatory = $true)]
        [String]
        $Body,
        [Parameter(Mandatory = $false)]
        [String]
        $AttachmentFileName,
        [int]
        $Port = 587,
        [Parameter(Mandatory = $false)]
        [String]
        $ClientId = "20773535-6b8f-4f3d-8f0e-4b7710d79afe",
        [Parameter(Mandatory = $false)]
        [String]
        $RedirectURI = "msal20773535-6b8f-4f3d-8f0e-4b7710d79afe://auth"
    )
    Process {       

        $mailMessage = New-Object System.Net.Mail.MailMessage
        $mailMessage.From = New-Object System.Net.Mail.MailAddress($SendingAddress)
        $mailMessage.To.Add($To)
        $mailMessage.Subject = $Subject
        $mailMessage.Body = $Body     
        if(![String]::IsNullOrEmpty($AttachmentFileName)){
            $attachment = New-Object System.Net.Mail.Attachment -ArgumentList $AttachmentFileName
            $mailMessage.Attachments.Add($attachment);
        }        

        $binding = [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic
        $MessageType = $mailMessage.GetType()
        $smtpClient = New-Object System.Net.Mail.SmtpClient
        $scType = $smtpClient.GetType()
        $booleanType = [System.Type]::GetType("System.Boolean")
        $assembly = $scType.Assembly
        $mailWriterType = $assembly.GetType("System.Net.Mail.MailWriter")
        $MemoryStream = New-Object -TypeName "System.IO.MemoryStream"
        $typeArray = ([System.Type]::GetType("System.IO.Stream"))
        $mailWriterConstructor = $mailWriterType.GetConstructor($binding ,$null, $typeArray, $null)
        [System.Array]$paramArray = ($MemoryStream)
        $mailWriter = $mailWriterConstructor.Invoke($paramArray)
        $doubleBool = $true
        $typeArray = ($mailWriter.GetType(),$booleanType,$booleanType)
        $sendMethod = $MessageType.GetMethod("Send", $binding, $null, $typeArray, $null)
        if ($null -eq $sendMethod) {
            $doubleBool = $false
            [System.Array]$typeArray = ($mailWriter.GetType(),$booleanType)
            $sendMethod = $MessageType.GetMethod("Send", $binding, $null, $typeArray, $null)
         }
        [System.Array]$typeArray = @()
        $closeMethod = $mailWriterType.GetMethod("Close", $binding, $null, $typeArray, $null)
        [System.Array]$sendParams = ($mailWriter,$true)
        if ($doubleBool) {
            [System.Array]$sendParams = ($mailWriter,$true,$true)
        }
        $sendMethod.Invoke($mailMessage,$binding,$null,$sendParams,$null)
        [System.Array]$closeParams = @()
        $MessageString = [System.Text.Encoding]::UTF8.GetString($MemoryStream.ToArray());
        $closeMethod.Invoke($mailWriter,$binding,$null,$closeParams,$null)
        [Void]$MemoryStream.Dispose()
        [Void]$mailMessage.Dispose()
        $MessageString = $MessageString.SubString($MessageString.IndexOf("MIME-Version:"))
        $socket = new-object System.Net.Sockets.TcpClient($ServerName, $Port)
        $stream = $socket.GetStream()
        $streamWriter = new-object System.IO.StreamWriter($stream)
        $streamReader = new-object System.IO.StreamReader($stream)
        $streamWriter.AutoFlush = $true
        $sslStream = New-Object System.Net.Security.SslStream($stream)
        $sslStream.ReadTimeout = 30000
        $sslStream.WriteTimeout = 30000        
        $ConnectResponse = $streamReader.ReadLine();
        Write-Host($ConnectResponse)
        if(!$ConnectResponse.StartsWith("220")){
            throw "Error connecting to the SMTP Server"
        }
        $Domain = $SendingAddress.Split('@')[1]
        Write-Host(("helo " + $Domain)) -ForegroundColor Green
        $streamWriter.WriteLine(("helo " + $Domain));
        $ehloResponse = $streamReader.ReadLine();
        Write-Host($ehloResponse)
        if (!$ehloResponse.StartsWith("250")){
            throw "Error in ehelo Response"
        }
        Write-Host("STARTTLS") -ForegroundColor Green
        $streamWriter.WriteLine("STARTTLS");
        $startTLSResponse = $streamReader.ReadLine();
        Write-Host($startTLSResponse)
        $ccCol = New-Object System.Security.Cryptography.X509Certificates.X509CertificateCollection
        $sslStream.AuthenticateAsClient($ServerName,$ccCol,[System.Security.Authentication.SslProtocols]::Tls12,$false);        
        $SSLstreamReader = new-object System.IO.StreamReader($sslStream)
        $SSLstreamWriter = new-object System.IO.StreamWriter($sslStream)
        $SSLstreamWriter.AutoFlush = $true
        $SSLstreamWriter.WriteLine(("helo " + $Domain));
        $ehloResponse = $SSLstreamReader.ReadLine();
        Write-Host($ehloResponse)
        $command = "AUTH XOAUTH2" 
        write-host -foregroundcolor DarkGreen $command
        $SSLstreamWriter.WriteLine($command) 
        $AuthLoginResponse = $SSLstreamReader.ReadLine()
        write-host ($AuthLoginResponse)
        $token = Get-AccessTokenForAzure -MailboxName $userName -ClientId $ClientId -RedirectURI $RedirectURI
        $SALSHeaderBytes = [System.Text.Encoding]::ASCII.GetBytes(("user=" + $userName + [char]1 + "auth=Bearer " + $token + [char]1 + [char]1))
        $Base64AuthSALS = [Convert]::ToBase64String($SALSHeaderBytes)     
        write-host -foregroundcolor DarkGreen $Base64AuthSALS
        $SSLstreamWriter.WriteLine($Base64AuthSALS)        
        $AuthResponse = $SSLstreamReader.ReadLine()
        write-host $AuthResponse
        if($AuthResponse.StartsWith("235")){
            $command = "MAIL FROM: <" + $SendingAddress + ">" 
            write-host -foregroundcolor DarkGreen $command
            $SSLstreamWriter.WriteLine($command) 
            $FromResponse = $SSLstreamReader.ReadLine()
            write-host $FromResponse
            $command = "RCPT TO: <" + $To + ">" 
            write-host -foregroundcolor DarkGreen $command
            $SSLstreamWriter.WriteLine($command) 
            $ToResponse = $SSLstreamReader.ReadLine()
            write-host $ToResponse
            $command = "DATA"
            write-host -foregroundcolor DarkGreen $command
            $SSLstreamWriter.WriteLine($command) 
            $DataResponse = $SSLstreamReader.ReadLine()
            write-host $DataResponse
            write-host -foregroundcolor DarkGreen $MessageString
            $SSLstreamWriter.WriteLine($MessageString) 
            $SSLstreamWriter.WriteLine(".") 
            $DataResponse = $SSLstreamReader.ReadLine()
            write-host $DataResponse
            $command = "QUIT" 
            write-host -foregroundcolor DarkGreen $command
            $SSLstreamWriter.WriteLine($command) 
            # ## Close the streams 
            $stream.Close() 
            $sslStream.Close()
            Write-Host("Done")
        }  

    }


}