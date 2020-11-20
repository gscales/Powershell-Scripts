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

function Get-FolderFromPath {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string]
        $FolderPath,
		
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
        $AutoPrompt	
    )

    process {
        
        $prompt = $true
        if($AutoPrompt.IsPresent){
            $prompt = $false
        }
        $EndPoint = "https://graph.microsoft.com/v1.0/users"
        $RequestURL = $EndPoint + "('$MailboxName')/MailFolders('MsgFolderRoot')/childfolders?"
        $AccessToken = Get-AccessTokenForGraph -MailboxName $Mailboxname -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -Prompt:$prompt
        $fldArray = $FolderPath.Split("\")
        $PropList = @()
        $FolderSizeProp = Get-TaggedProperty -DataType "Long" -Id "0x66b3"
        $EntryId = Get-TaggedProperty -DataType "Binary" -Id "0xfff"
        $PropList += $FolderSizeProp 
        $PropList += $EntryId
        $Props = Get-ExtendedPropList -PropertyList $PropList 
        $RequestURL += "`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
        #Loop through the Split Array and do a Search for each level of folder 
        for ($lint = 1; $lint -lt $fldArray.Length; $lint++) {
            #Perform search based on the displayname of each folder level
            $FolderName = $fldArray[$lint];
            $headers = @{
                'Authorization' = "Bearer $AccessToken"
                'AnchorMailbox' = "$MailboxName"
            }
            $RequestURL = $RequestURL += "`&`$filter=DisplayName eq '$FolderName'"
            $tfTargetFolder = (Invoke-RestMethod -Method Get -Uri $RequestURL -UserAgent "GraphBasicsPs101" -Headers $headers).value  
            if ($tfTargetFolder.displayname -match $FolderName) {
                $folderId = $tfTargetFolder.Id.ToString()
                $RequestURL = $EndPoint + "('$MailboxName')/MailFolders('$folderId')/childfolders?"
                $RequestURL += "`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
            }
            else {
                throw ("Folder Not found")
            }
        }
        if ($tfTargetFolder.singleValueExtendedProperties) {
            foreach ($Prop in $tfTargetFolder.singleValueExtendedProperties) {
                Write-Verbose $Prop.Id
                Switch ($Prop.Id) {
                    "Long 0x66b3" {      
                        $tfTargetFolder | Add-Member -NotePropertyName "FolderSize" -NotePropertyValue $Prop.value 
                    }
                    "Binary 0xfff" {
                        $tfTargetFolder | Add-Member -NotePropertyName "PR_ENTRYID" -NotePropertyValue ([System.BitConverter]::ToString([Convert]::FromBase64String($Prop.Value)).Replace("-", ""))
                        $tfTargetFolder | Add-Member -NotePropertyName "ComplianceSearchId" -NotePropertyValue ("folderid:" + $tfTargetFolder.PR_ENTRYID.SubString(($tfTargetFolder.PR_ENTRYID.length - 48)))
                    }
                    "Binary {00062004-0000-0000-c000-000000000046} Name FromFavoriteSendersFolderEntryId"  {
                        $tfTargetFolder | Add-Member -NotePropertyName "FromFavoriteSendersFolderEntryId" -NotePropertyValue ([System.BitConverter]::ToString([Convert]::FromBase64String($Prop.Value)).Replace("-", ""))
                    }
                }
            }
        }
        return $tfTargetFolder 
    }
}

function Get-WellKnownFolder {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string]
        $FolderName,		
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
        [psobject]
        $PropList,
        [Parameter(Position = 7, Mandatory = $false)]
        [String]
        $AccessToken		
    )

    process {
        
        $prompt = $true
        if($AutoPrompt.IsPresent){
            $prompt = $false
        }
        $EndPoint = "https://graph.microsoft.com/v1.0/users"
        $RequestURL = $EndPoint + "('$MailboxName')/MailFolders('$FolderName')?"
        if(!$AccessToken){
            $AccessToken = Get-AccessTokenForGraph -MailboxName $Mailboxname -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -Prompt:$prompt
        }       
        if(!$PropList){
            $PropList = @()
        }        
        $FolderSizeProp = Get-TaggedProperty -DataType "Long" -Id "0x66b3"
        $EntryId = Get-TaggedProperty -DataType "Binary" -Id "0xfff"
        $PropList += $FolderSizeProp 
        $PropList += $EntryId
        $Props = Get-ExtendedPropList -PropertyList $PropList 
            $headers = @{
                'Authorization' = "Bearer $AccessToken"
                'AnchorMailbox' = "$MailboxName"
            }
        $RequestURL += "`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
        $tfTargetFolder = (Invoke-RestMethod -Method Get -Uri $RequestURL -UserAgent "GraphBasicsPs101" -Headers $headers)  
        if ($tfTargetFolder.singleValueExtendedProperties) {
            foreach ($Prop in $tfTargetFolder.singleValueExtendedProperties) {
                Switch ($Prop.Id) {
                    "Long 0x66b3" {      
                        $tfTargetFolder | Add-Member -NotePropertyName "FolderSize" -NotePropertyValue $Prop.value 
                    }
                    "Binary 0xfff" {
                        $tfTargetFolder | Add-Member -NotePropertyName "PR_ENTRYID" -NotePropertyValue ([System.BitConverter]::ToString([Convert]::FromBase64String($Prop.Value)).Replace("-", ""))
                        $tfTargetFolder | Add-Member -NotePropertyName "ComplianceSearchId" -NotePropertyValue ("folderid:" + $tfTargetFolder.PR_ENTRYID.SubString(($tfTargetFolder.PR_ENTRYID.length - 48)))
                    }
                 
                }
            }
        }
        return $tfTargetFolder 
    }
}


function Get-FolderFromId {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string]
        $FolderId,		
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
        [psobject]
        $PropList,
        [Parameter(Position = 7, Mandatory = $false)]
        [String]
        $AccessToken		
    )

    process {
        
        $prompt = $true
        if($AutoPrompt.IsPresent){
            $prompt = $false
        }
        $EndPoint = "https://graph.microsoft.com/v1.0/users"
        $RequestURL = $EndPoint + "('$MailboxName')/MailFolders('$FolderName')?"
        if(!$AccessToken){
            $AccessToken = Get-AccessTokenForGraph -MailboxName $Mailboxname -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -Prompt:$prompt
        }       
        if(!$PropList){
            $PropList = @()
        }        
        $FolderSizeProp = Get-TaggedProperty -DataType "Long" -Id "0x66b3"
        $EntryId = Get-TaggedProperty -DataType "Binary" -Id "0xfff"
        $PropList += $FolderSizeProp 
        $PropList += $EntryId
        $Props = Get-ExtendedPropList -PropertyList $PropList 
            $headers = @{
                'Authorization' = "Bearer $AccessToken"
                'AnchorMailbox' = "$MailboxName"
            }
        $RequestURL += "`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
        $tfTargetFolder = (Invoke-RestMethod -Method Get -Uri $RequestURL -UserAgent "GraphBasicsPs101" -Headers $headers)  
        if ($tfTargetFolder.singleValueExtendedProperties) {
            foreach ($Prop in $tfTargetFolder.singleValueExtendedProperties) {
                Switch ($Prop.Id) {
                    "Long 0x66b3" {      
                        $tfTargetFolder | Add-Member -NotePropertyName "FolderSize" -NotePropertyValue $Prop.value 
                    }
                    "Binary 0xfff" {
                        $tfTargetFolder | Add-Member -NotePropertyName "PR_ENTRYID" -NotePropertyValue ([System.BitConverter]::ToString([Convert]::FromBase64String($Prop.Value)).Replace("-", ""))
                        $tfTargetFolder | Add-Member -NotePropertyName "ComplianceSearchId" -NotePropertyValue ("folderid:" + $tfTargetFolder.PR_ENTRYID.SubString(($tfTargetFolder.PR_ENTRYID.length - 48)))
                    }
                 
                }
            }
        }
        return $tfTargetFolder 
    }
}

function Get-ApplicationFolder{
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
            [Parameter(Position = 6, Mandatory = $false)]
            [psobject]
            $PropList		
        )
    
        process {
            
            $prompt = $true
            if($AutoPrompt.IsPresent){
                $prompt = $false
            }
            $appId = $Proplist[0].Id
            $AccessToken = Get-AccessTokenForGraph -MailboxName $Mailboxname -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -Prompt:$prompt
            $RootFolder = Get-WellKnownFolder -MailboxName $MailboxName -FolderName Inbox -PropList $PropList -AccessToken $AccessToken
            foreach ($Prop in $RootFolder.singleValueExtendedProperties) {
                if($Prop.Id -match $PropList[0].Id){
                    $RootFolder | Add-Member -NotePropertyName $PropList[0].Id -NotePropertyValue ([System.BitConverter]::ToString([Convert]::FromBase64String($Prop.Value)).Replace("-", ""))   
                }
            }
            $FolderId = Invoke-TranslateExchangeIds -SourceHexId $RootFolder."$appId" -SourceFormat entryid -TargetFormat restid -AccessToken $AccessToken -MailboxName $MailboxName
            $EndPoint = "https://graph.microsoft.com/v1.0/users"
            $RequestURL = $EndPoint + "('$MailboxName')/MailFolders/" + $FolderId
            $headers = @{
                'Authorization' = "Bearer $AccessToken"
                'AnchorMailbox' = "$MailboxName"
            }           
            $tfTargetFolder = (Invoke-RestMethod -Method Get -Uri $RequestURL -UserAgent "GraphBasicsPs101" -Headers $headers)  
            return $tfTargetFolder
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

function Get-NamedProperty
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[String]
		$DataType,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[String]
		$Id,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[String]
		$Guid,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[String]
		$Type,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[String]
		$Value
	)
	Begin
	{
		$Property = "" | Select-Object Id, DataType, PropertyType, Type, Guid, Value
		$Property.Id = $Id
		$Property.DataType = $DataType
		$Property.PropertyType = "Named"
		$Property.Guid = $Guid
		if ($Type -eq "String")
		{
			$Property.Type = "String"
		}
		else
		{
			$Property.Type = "Id"
		}
		if (![String]::IsNullOrEmpty($Value))
		{
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

function Invoke-TranslateExchangeIds {
    [CmdletBinding()]
    param (     
        [Parameter(Position = 1, Mandatory = $false)]
        [String]
        $MailboxName, 

        [Parameter(Position = 2, Mandatory = $false)]
        [String]
        $SourceId, 
        
        [Parameter(Position = 3, Mandatory = $false)]
        [String]
        $SourceHexId,
        
        [Parameter(Position = 4, Mandatory = $false)]
        [String]
        $SourceEMSId,

        [Parameter(Position = 5, Mandatory = $false)]
        [String]
        $SourceFormat,  

        [Parameter(Position = 6, Mandatory = $false)]
        [String]
        $TargetFormat,       
        
        [Parameter(Position = 7, Mandatory = $false)]
        [String]
        $AccessToken  
    )
    Begin {		

        $ConvertRequest = @{}
        $ConvertRequest.Add("inputIds", @())
        if ($SourceHexId) {
            $byteArray = @($SourceHexId -split '([a-f0-9]{2})' | foreach-object { if ($_) {[System.Convert]::ToByte($_, 16)}})
            $urlSafeString = [Convert]::ToBase64String($byteArray).replace("/", "_").replace("+", "-")
            if ($urlSafeString.contains("==")) {$urlSafeString = $urlSafeString.replace("==", "2")}
            $ConvertRequest["inputIds"] += $urlSafeString

        }
        else {
            if ($SourceEMSId) {
                $HexEntryId = [System.BitConverter]::ToString([Convert]::FromBase64String($SourceEMSId)).Replace("-", "").Substring(2)  
                $HexEntryId = $HexEntryId.SubString(0, ($HexEntryId.Length - 2))
                $byteArray = @($HexEntryId -split '([a-f0-9]{2})' | foreach-object { if ($_) {[System.Convert]::ToByte($_, 16)}})
                $urlSafeString = [Convert]::ToBase64String($byteArray).replace("/", "_").replace("+", "-")
                if ($urlSafeString.contains("==")) {$urlSafeString = $urlSafeString.replace("==", "2")}
                $ConvertRequest["inputIds"] += $urlSafeString
            }
            else {
                $ConvertRequest["inputIds"] += $SourceId
            }

        }        
        $ConvertRequest.targetIdType = $TargetFormat
        $ConvertRequest.sourceIdType = $SourceFormat
        $headers = @{
            'Authorization' = "Bearer $AccessToken"
            'AnchorMailbox' = "$MailboxName"
        }      
        $JsonResult = Invoke-RestMethod -Headers $headers -Uri ("https://graph.microsoft.com/v1.0/users('$MailboxName')/translateExchangeIds") -Method Post -ContentType "application/json" -Body (ConvertTo-Json $ConvertRequest -Depth 9)
        if($TargetFormat.ToLower() -eq "entryid"){
            $urldecodedstring = $JsonResult.value.targetId.replace("_", "/").replace("-", "+")
            $lastVal = $urldecodedstring.SubString($urldecodedstring.Length-1,1);
            if($lastVal -eq "2"){
                $urldecodedstring = $urldecodedstring.SubString(0,$urldecodedstring.Length-1) + "=="
            }
            return ([System.BitConverter]::ToString([Convert]::FromBase64String($urldecodedstring))).replace("-","")
        }else{
            return  $JsonResult.value.targetId
        }
       	
		
    }
}

