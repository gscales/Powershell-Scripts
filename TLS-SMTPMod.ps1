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

function Invoke-TestSMTPTLS {
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
        [Parameter(Mandatory = $false)] 
        [System.Management.Automation.PSCredential]$Credentials,
        [int]
        $Port = 587
    )
    Process {
        $socket = new-object System.Net.Sockets.TcpClient($ServerName, $Port)
        $stream = $socket.GetStream()
        $streamWriter = new-object System.IO.StreamWriter($stream)
        $streamReader = new-object System.IO.StreamReader($stream)
        $streamWriter.AutoFlush = $true
        $sslStream = New-Object System.Net.Security.SslStream($stream)
        $sslStream.ReadTimeout = 5000
        $sslStream.WriteTimeout = 5000        
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
        if($Credentials){
            $command = "AUTH LOGIN" 
            write-host -foregroundcolor DarkGreen $command
            $SSLstreamWriter.WriteLine($command) 
            $AuthLoginResponse = $SSLstreamReader.ReadLine()
            write-host ($AuthLoginResponse)
            $Bytes = [System.Text.Encoding]::ASCII.GetBytes($Credentials.UserName)
            $Base64UserName =  [Convert]::ToBase64String($Bytes)              
            $SSLstreamWriter.WriteLine($Base64UserName)
            $UserNameResponse = $SSLstreamReader.ReadLine()
            write-host ($UserNameResponse)
            $Bytes = [System.Text.Encoding]::ASCII.GetBytes($Credentials.GetNetworkCredential().password.ToString())
            $Base64Password = [Convert]::ToBase64String($Bytes)     
            $SSLstreamWriter.WriteLine($Base64Password)
            $PassWordResponse = $SSLstreamReader.ReadLine()
            write-host $PassWordResponse    
        }
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

function readResponse() {
    while($stream.DataAvailable)
    {
        $buffer = new-object System.Byte[] 1024
        $read = $stream.Read($buffer, 0, 1024)
        $rstring = $encoding.GetString($buffer, 0, $read)
        Write-Host $rstring
    }
}