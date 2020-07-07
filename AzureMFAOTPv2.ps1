<#
    .SYNOPSIS
       Does a Multi-Factor authentication against Azure using UserName and Password and a One Time Password OTP as the second factor
 
    .DESCRIPTION
      Does a Multi-Factor authentication against Azure using UserName and Password and a One Time Password OTP as the second factor
 
    .INPUTS
       PSCrednetails which will contain the username and password for the Primary Auth
       OTP code can be generated with the Get-TimeBasedOneTimePassword function (requires that this be setup beforehand)
 
    .OUTPUTS
        AccessToken
 
    .EXAMPLE
        PS C:\> Get-AccessTokenMFA -OTP 123456
        
        Example use a SharedSecret stored in the Windows Credential Store

        PC C:\> Get-AccessTokenMFA -OTP (Get-TimeBasedOneTimePassword -SharedSecret (Get-StoredCredential -Target StoredAuth -AsCredentialObject).Password)
    .NOTES
        Author : Glen Scales
 
    .LINK
        

#>
function Get-AccessTokenMFA{
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [PSCredential]
        $Credential,
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $OTP,
        [Parameter(Position = 2, Mandatory = $false)]
        [String]
        $ClientId = "95a1b05c-60f1-420d-a5b2-0cca170dfadc",
        [Parameter(Position = 3, Mandatory = $false)]
        [String]
        $RedirectURI = "https://login.microsoftonline.com/common/oauth2/nativeclient",
        [Parameter(Position = 4, Mandatory = $false)]
        [String]
        $scopes = "https://outlook.office.com/EWS.AccessAsUser.All"
    )
    process {        
        $domain = $Credential.UserName.Split('@')[1]
        $openidURL = "https://login.windows.net/$domain/v2.0/.well-known/openid-configuration"
        $TenantId = (Invoke-WebRequest -Uri $openidURL | ConvertFrom-Json).token_endpoint.Split('/')[3] 
        $AuthURL = "https://login.microsoftonline.com/common/oauth2/authorize?client_id=$ClientId&response_mode=form_post&response_type=code&redirect_uri=" + [System.Web.HttpUtility]::UrlEncode($RedirectURI)
        $StartLogon = Invoke-WebRequest -uri $AuthURL  -SessionVariable 'AuthSession'
        $Context = [regex]::Match($StartLogon.RawContent,"`"sCtx`":`"(.*?)`"").Groups[1].Value
        $Flow = [regex]::Match($StartLogon.RawContent,"`"sFT`":`"(.*?)`"").Groups[1].Value
        $Canary = [regex]::Match($StartLogon.RawContent,"`"canary`":`"(.*?)`"").Groups[1].Value
        $FBAAuthBody=@{
            "login" = $Credential.UserName
            "loginFmt" = $Credential.UserName
            "i13"="0"
            "type"="11"
            "LoginOptions"="3"
            "passwd"= $Credential.GetNetworkCredential().password.ToString()
            "ps"="2"
            "flowToken"=$Flow
            "canary"=$Canary
            "ctx"=$Context
            "NewUser"="1"
            "fspost"="0"
            "i21"="0"
            "CookieDisclosure"="1"
            "IsFidoSupported"="1"
            "hpgrequestid"=(New-Guid).ToString()
        }
        $FBAResponse = Invoke-WebRequest -Uri "https://login.microsoftonline.com/common/login" -Method Post -ContentType "application/x-www-form-urlencoded" -Body $FBAAuthBody -WebSession $AuthSession
        $Context = [regex]::Match($FBAResponse.RawContent,"`"sCtx`":`"(.*?)`"").Groups[1].Value
        $Flow = [regex]::Match($FBAResponse.RawContent,"`"sFT`":`"(.*?)`"").Groups[1].Value
        $SasBegin=@{
            "AuthMethodId" = "PhoneAppOTP"
            "flowToken"=$Flow
            "ctx"=$Context
            "Method"="BeginAuth"
        }
        $SASBeginResponse = (Invoke-RestMethod -Uri "https://login.microsoftonline.com/common/SAS/BeginAuth" -Method Post -ContentType "application/json" -Body ($SasBegin | ConvertTo-Json) -WebSession $AuthSession)
        $SASEnd=@{
            "AdditionalAuthData"=$OTP
            "AuthMethodId"="PhoneAppOTP"
            "flowToken"=$SASBeginResponse.flowToken
            "ctx"=$SASBeginResponse.ctx
            "Method"="EndAuth"
            "PollCount"=1
            "SessionId"=$SASBeginResponse.SessionId
        }
        $SASEndResponse = (Invoke-RestMethod -Uri "https://login.microsoftonline.com/common/SAS/EndAuth" -Method Post -ContentType "application/json" -Body ($SASEnd | ConvertTo-Json) -WebSession $AuthSession)
        $SASProcess=@{
            "type"=19
            "GeneralVerify"="false"
            "otc"=$OTP
            "login"= $Credential.UserName
            "mfaAuthMethod"="PhoneAppOTP"
            "flowToken"=$SASEndResponse.flowToken
            "request"=$SASEndResponse.ctx
            "Method"="EndAuth"
            "PollCount"=1
            "SessionId"=$SASEndResponse.SessionId
            "canary"=$Canary
            "hpgrequestid"=(New-Guid).ToString()
        }
        $SASProcessResponse = (Invoke-WebRequest -Uri "https://login.microsoftonline.com/common/SAS/ProcessAuth" -Method Post -ContentType "application/x-www-form-urlencoded" -Body $SASProcess -WebSession $AuthSession)
        $formElements = ([XML]$SASProcessResponse.Content).GetElementsByTagName("input");  
        $authCode = ""
        foreach ($element in $formElements) {
            if ($element.Name -eq "code") {
                $authCode = $element.GetAttribute("value");
                Write-Verbose $authCode
            }
        }  
        $Body = @{"grant_type" = "authorization_code"; "scope" = $scopes; "client_id" = "$ClientId"; "code" = $authCode; "redirect_uri" = $RedirectURI }
        $tokenRequest = Invoke-RestMethod -Method Post -ContentType application/x-www-form-urlencoded -Uri "https://login.microsoftonline.com/$tenantid/oauth2/v2.0/token" -Body $Body 
        return $tokenRequest
    }
}

<#
    .SYNOPSIS
        Generate a Time-Base One-Time Password based on RFC 6238.
 
    .DESCRIPTION
        This command uses the reference implementation of RFC 6238 to calculate
        a Time-Base One-Time Password. It bases on the HMAC SHA-1 hash function
        to generate a shot living One-Time Password.
 
    .INPUTS
        None.
 
    .OUTPUTS
        System.String. The one time password.
 
    .EXAMPLE
        PS C:\> Get-TimeBasedOneTimePassword -SharedSecret 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        Get the Time-Based One-Time Password at the moment.
 
    .NOTES
        Author : Claudio Spizzi
        License : MIT License
 
    .LINK
        https://github.com/claudiospizzi/SecurityFever
        https://tools.ietf.org/html/rfc6238
#>
function Get-TimeBasedOneTimePassword
{
    [CmdletBinding()]
    [Alias('Get-TOTP')]
    param
    (
        # Base 32 formatted shared secret (RFC 4648).
        [Parameter(Mandatory = $true)]
        [System.String]
        $SharedSecret,

        # The date and time for the target calculation, default is now (UTC).
        [Parameter(Mandatory = $false)]
        [System.DateTime]
        $Timestamp = (Get-Date).ToUniversalTime(),

        # Token length of the one-time password, default is 6 characters.
        [Parameter(Mandatory = $false)]
        [System.Int32]
        $Length = 6,

        # The hash method to calculate the TOTP, default is HMAC SHA-1.
        [Parameter(Mandatory = $false)]
        [System.Security.Cryptography.KeyedHashAlgorithm]
        $KeyedHashAlgorithm = (New-Object -TypeName 'System.Security.Cryptography.HMACSHA1'),

        # Baseline time to start counting the steps (T0), default is Unix epoch.
        [Parameter(Mandatory = $false)]
        [System.DateTime]
        $Baseline = '1970-01-01 00:00:00',

        # Interval for the steps in seconds (TI), default is 30 seconds.
        [Parameter(Mandatory = $false)]
        [System.Int32]
        $Interval = 30
    )

    # Generate the number of intervals between T0 and the timestamp (now) and
    # convert it to a byte array with the help of Int64 and the bit converter.
    $numberOfSeconds   = ($Timestamp - $Baseline).TotalSeconds
    $numberOfIntervals = [Convert]::ToInt64([Math]::Floor($numberOfSeconds / $Interval))
    $byteArrayInterval = [System.BitConverter]::GetBytes($numberOfIntervals)
    [Array]::Reverse($byteArrayInterval)

    # Use the shared secret as a key to convert the number of intervals to a
    # hash value.
    $KeyedHashAlgorithm.Key = Convert-Base32ToByte -Base32 $SharedSecret
    $hash = $KeyedHashAlgorithm.ComputeHash($byteArrayInterval)

    # Calculate offset, binary and otp according to RFC 6238 page 13.
    $offset = $hash[($hash.Length-1)] -band 0xf
    $binary = (($hash[$offset + 0] -band '0x7f') -shl 24) -bor
              (($hash[$offset + 1] -band '0xff') -shl 16) -bor
              (($hash[$offset + 2] -band '0xff') -shl 8) -bor
              (($hash[$offset + 3] -band '0xff'))
    $otpInt = $binary % ([Math]::Pow(10, $Length))
    $otpStr = $otpInt.ToString().PadLeft($Length, '0')

    Write-Output $otpStr
}

function Convert-Base32ToByte
{
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Base32
    )

    # RFC 4648 Base32 alphabet
    $rfc4648 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ234567'

    $bits = ''

    # Convert each Base32 character to the binary value between starting at
    # 00000 for A and ending with 11111 for 7.
    foreach ($char in $Base32.ToUpper().ToCharArray())
    {
        $bits += [Convert]::ToString($rfc4648.IndexOf($char), 2).PadLeft(5, '0')
    }

    # Convert 8 bit chunks to bytes, ignore the last bits.
    for ($i = 0; $i -le ($bits.Length - 8); $i += 8)
    {
        [Byte] [Convert]::ToInt32($bits.Substring($i, 8), 2)
    }
}









