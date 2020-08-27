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

function Get-ModerationRequests {
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
        [Parameter(Position = 5, Mandatory = $false)]
        [Int32]
        $Top=1,
        [Parameter(Position = 6, Mandatory = $false)]
        [String]
        $AccessToken
    )

    process {
        
        $prompt = $true
        if($AutoPrompt.IsPresent){
            $prompt = $false
        }
        $EndPoint = "https://graph.microsoft.com/v1.0/users"
        $RequestURL = $EndPoint + "('$MailboxName')/MailFolders('Inbox')/messages?`$filter=singleValueExtendedProperties/any(ep:ep/id eq 'String 0x001a' and ep/value eq 'IPM.Note.Microsoft.Approval.Request')"
        if([String]::IsNullOrEmpty($AccessToken)){
            $AccessToken = Get-AccessTokenForGraph -MailboxName $Mailboxname -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -Prompt:$prompt
        }        
        $PropList = @()
        $ReportTag = Get-TaggedProperty -DataType "Binary" -Id "0x0031"
        $NormalizedSubject = Get-TaggedProperty -DataType "String" -Id "0x0E1D"
        $PropList += $ReportTag 
        $PropList += $NormalizedSubject
        $Props = Get-ExtendedPropList -PropertyList $PropList 
        $RequestURL += "&`$Top=" + $Top + "&`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
        $headers = @{
            'Authorization' = "Bearer $AccessToken"
            'AnchorMailbox' = "$MailboxName"
        }
        $ModerationRequests = (Invoke-RestMethod -Method Get -Uri $RequestURL -UserAgent "GraphBasicsPs101" -Headers $headers).value
        return  $ModerationRequests
    }
}

function Invoke-ApproveModerationRequest{
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
        [Parameter(Position = 5, Mandatory = $false)]
        [Int32]
        $Top=1,
        [Parameter(Position = 6, Mandatory = $false)]
        [String]
        $AccessToken,
        [Parameter(Position = 7, Mandatory = $false)]
        [psobject]
        $ApprovalMail

    )

    process {
        $prompt = $true
        if($AutoPrompt.IsPresent){
            $prompt = $false
        }
        $EndPoint = "https://graph.microsoft.com/v1.0/users"
        if($ApprovalMail -eq $null){            
            if([String]::IsNullOrEmpty($AccessToken)){
                $AccessToken = Get-AccessTokenForGraph -MailboxName $Mailboxname -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -Prompt:$prompt
            }
            $ApprovalMail =  Get-ModerationRequests -MailboxName $MailboxName -AccessToken $AccessToken
        }
        if($ApprovalMail){
            $RequestURL =  $EndPoint + "('$MailboxName')/SendMail"
            $SendRequest = [ordered]@{
                message = @{
                    subject = "Approve: " + $ApprovalMail.singleValueExtendedProperties[1].value
                    toRecipients = @(
                            @{
                                emailAddress = @{
                                        address = $ApprovalMail.sender.emailAddress.address                            
                                }
                            }
                   ) 
                    singleValueExtendedProperties = @(
                        @{
                            id = "Binary 0x31"
                            value =  $ApprovalMail.singleValueExtendedProperties[0].value
                        },
                        @{
                            id = "String 0x001A"
                            value = "IPM.Note.Microsoft.Approval.Reply.Approve"
                        },
                        @{
                            id = "String {00062008-0000-0000-C000-000000000046} Id 0x8524"
                            value = "Approve"
                        }
                    )
    
                }
            }
            $MessageToSend = ConvertTo-Json -InputObject $SendRequest -Depth 9 
            $headers = @{
                'Authorization' = "Bearer $AccessToken"
                'AnchorMailbox' = "$MailboxName"
            }
            return (Invoke-RestMethod -Method Post -Uri $RequestURL -UserAgent "GraphBasicsPs101" -Headers $headers -ContentType 'Application/json' -Body $MessageToSend)
        }else{
            return "No Moderation Messages Found"
        }


    }
}

function Invoke-RejectModerationRequest{
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
        [Parameter(Position = 5, Mandatory = $false)]
        [Int32]
        $Top=1,
        [Parameter(Position = 6, Mandatory = $false)]
        [String]
        $AccessToken,
        [Parameter(Position = 7, Mandatory = $false)]
        [psobject]
        $ApprovalMail

    )

    process {
        $prompt = $true
        if($AutoPrompt.IsPresent){
            $prompt = $false
        }
        $EndPoint = "https://graph.microsoft.com/v1.0/users"
        if($ApprovalMail -eq $null){            
            if([String]::IsNullOrEmpty($AccessToken)){
                $AccessToken = Get-AccessTokenForGraph -MailboxName $Mailboxname -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -Prompt:$prompt
            }
            $ApprovalMail =  Get-ModerationRequests -MailboxName $MailboxName -AccessToken $AccessToken
        }
        if($ApprovalMail){
            $RequestURL =  $EndPoint + "('$MailboxName')/SendMail"
            $SendRequest = [ordered]@{
                message = @{
                    subject = "Approve: " + $ApprovalMail.singleValueExtendedProperties[1].value
                    toRecipients = @(
                            @{
                                emailAddress = @{
                                        address = $ApprovalMail.sender.emailAddress.address                            
                                }
                            }
                   ) 
                    singleValueExtendedProperties = @(
                        @{
                            id = "Binary 0x31"
                            value =  $ApprovalMail.singleValueExtendedProperties[0].value
                        },
                        @{
                            id = "String 0x001A"
                            value = "IPM.Note.Microsoft.Approval.Reply.Reject"
                        },
                        @{
                            id = "String {00062008-0000-0000-C000-000000000046} Id 0x8524"
                            value = "Reject"
                        }
                    )
    
                }
            }
            $MessageToSend = ConvertTo-Json -InputObject $SendRequest -Depth 9 
            $headers = @{
                'Authorization' = "Bearer $AccessToken"
                'AnchorMailbox' = "$MailboxName"
            }
            return (Invoke-RestMethod -Method Post -Uri $RequestURL -UserAgent "GraphBasicsPs101" -Headers $headers -ContentType 'Application/json' -Body $MessageToSend)
        }else{
            return "No Moderation Messages Found"
        }


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


