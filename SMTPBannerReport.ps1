function Invoke-GetSMTPBannerVersion{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $Domain,
        [Parameter(Mandatory = $false)]
        [int]
        $Port = 25

    )
    Process {       
        Clear-DnsClientCache
        $rptCollection = @()
        $buffer = New-Object System.Byte[] 20240
        $encoding = New-Object System.Text.AsciiEncoding
        try{
            $mxRecords = Resolve-DnsName -Name $Domain -Type MX -ErrorAction Stop   
            Write-Verbose ("Number of MX Records for $Domain " + $mxRecords.Count)
            foreach($mxRecord in $mxRecords){ 
                if($mxRecord.Type = 15 -band (![String]::IsNullOrEmpty($mxRecord.NameExchange))){
                    $rptObj = "" | select Domain,Server,MXPreference,Banner
                    $rptObj.Domain = $Domain
                    $rptObj.Server = $mxRecord.NameExchange   
                    $rptObj.MXPreference = $mxRecord.Preference 
                    Write-Verbose("Testing $mxRecord.NameExchange")
                    $client = new-object System.Net.Sockets.TcpClient
                    $client.ConnectAsync($mxRecord.NameExchange, $Port).Wait(5000).Result | Out-Null
                    if($client.Connected){
                        Write-Verbose("Connection Sucess")
                        $stream =  $client.GetStream()
                        $streamWriter = new-object System.IO.StreamWriter($stream)
                        $streamReader = new-object System.IO.StreamReader($stream)
                        $stream.ReadTimeout = 5000
                        $stream.WriteTimeout = 5000   
                        $streamWriter.AutoFlush = $true
                        try{
                            $read = $streamReader.BaseStream.Read($buffer,0,20240)
                            $ConnectResponse = ($encoding.GetString($buffer, 0, $read))
                            $buffer = New-Object System.Byte[] 20240             
                            Write-Verbose($ConnectResponse)
                            $rptObj.Banner = $ConnectResponse
                            if(!$ConnectResponse.StartsWith("220")){                       
                                throw "Error connecting to the SMTP Server " + $ConnectResponse
                            }      
                        }catch{
                            $ExceptionMessage = Invoke-GetExceptionMessage $PSItem.Exception
                            Write-Verbose $ExceptionMessage                                       
                        }
                        $streamReader.Close()
                        $streamReader.Dispose()
                        $streamWriter.Close()
                        $streamWriter.Dispose()
                        $client.Close()
                        $rptCollection += $rptObj
                    }else{
                        Write-Verbose("Connection Failed")
                    }
                    Write-Verbose("Done")
                } 
            }                  
           
        }catch{
            Write-Output $PSItem.Exception.Message
        }
        return $rptCollection    
    }   
}

function Invoke-GetExceptionMessage{
    [CmdletBinding()]
    param (
        [Parameter()]
        [PSObject]
        $Exception
    )
    Process{
        if(!$Exception.Message){
            $exceptionMsg = $Exception.Message
            while ($Exception.InnerException) {
                $Exception = $Exception.InnerException
                if(!$Exception.Message){
                    if(!$exceptionMsg.Contains($Exception.Message)){
                        $exceptionMsg += " " + $Exception.Message 
                    }  
                }             
            }
        }else{
            $exceptionMsg = $Exception
        }
        return $exceptionMsg
    } 

}

