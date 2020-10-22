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

function Get-RecoverableItemsDeletionsFolder{
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
        [String]
        $AccessToken
    )

    process {
        
        $prompt = $true
        if($AutoPrompt.IsPresent){
            $prompt = $false
        }
        $EndPoint = "https://graph.microsoft.com/v1.0/users"
        $RequestURL = $EndPoint + "('$MailboxName')/MailFolders('RecoverableItemsDeletions')"
        if([String]::IsNullOrEmpty($AccessToken)){
            $AccessToken = Get-AccessTokenForGraph -MailboxName $Mailboxname -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -Prompt:$prompt
        }        
        $headers = @{
            'Authorization' = "Bearer $AccessToken"
            'AnchorMailbox' = "$MailboxName"
        }
        $tfTargetFolder = (Invoke-RestMethod -Method Get -Uri $RequestURL -UserAgent "GraphBasicsPs101" -Headers $headers)  
        return $tfTargetFolder 
    }
}

function Get-RecoverableItemsPurgesFolder{
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
        [String]
        $AccessToken
    )

    process {
        
        $prompt = $true
        if($AutoPrompt.IsPresent){
            $prompt = $false
        }
        $EndPoint = "https://graph.microsoft.com/v1.0/users"
        $RequestURL = $EndPoint + "('$MailboxName')/MailFolders('RecoverableItemsPurges')"
        if([String]::IsNullOrEmpty($AccessToken)){
            $AccessToken = Get-AccessTokenForGraph -MailboxName $Mailboxname -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -Prompt:$prompt
        }     
        $headers = @{
            'Authorization' = "Bearer $AccessToken"
            'AnchorMailbox' = "$MailboxName"
        }
        $tfTargetFolder = (Invoke-RestMethod -Method Get -Uri $RequestURL -UserAgent "GraphBasicsPs101" -Headers $headers)  
        return $tfTargetFolder 
    }
}

function Get-RecoverableItemsFolders{
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
        [String]
        $AccessToken
    )

    process {
        
        $prompt = $true
        if($AutoPrompt.IsPresent){
            $prompt = $false
        }
        $EndPoint = "https://graph.microsoft.com/v1.0/users"
        $RequestURL = $EndPoint + "('$MailboxName')/MailFolders('RecoverableItemsRoot')/ChildFolders?"
        $PropList = @()
        $FolderSizeProp = Get-EXRTaggedProperty -DataType "Long" -Id "0x66B3"
        $PropList += $FolderSizeProp 
        $Props = Get-ExtendedPropList -PropertyList $PropList 
        $RequestURL += "`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")" 
        if([String]::IsNullOrEmpty($AccessToken)){
            $AccessToken = Get-AccessTokenForGraph -MailboxName $Mailboxname -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -Prompt:$prompt
        }     
        $headers = @{
            'Authorization' = "Bearer $AccessToken"
            'AnchorMailbox' = "$MailboxName"
        }
        $tfTargetFolders = (Invoke-RestMethod -Method Get -Uri $RequestURL -UserAgent "GraphBasicsPs101" -Headers $headers).value 
        foreach($folder in $tfTargetFolders){
            Expand-ExtendedProperties $folder
        }
        return $tfTargetFolders 
    }
}
function Get-RecoverableItems{
    [CmdletBinding()]
    param (
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $MailboxName,
        [Parameter(Position = 2, Mandatory = $false)]
        [String]
        $FolderName,
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
        [Int32]
        $MessageCount=10,
        [Parameter(Position = 9, Mandatory = $false)]
        [String]
        $filter,
        [Parameter(Position = 10, Mandatory = $false)]
        [String]
        $SelectList = "sender,Subject,receivedDateTime,lastModifiedDateTime,internetmessageid,parentFolderId"
    )

    process {
        $prompt = $true
        if($AutoPrompt.IsPresent){
            $prompt = $false
        }
        if([String]::IsNullOrEmpty($AccessToken)){
            $AccessToken = Get-AccessTokenForGraph -MailboxName $Mailboxname -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -Prompt:$prompt
        }     
        if($MessageCount -lt 100){
            $top=$MessageCount
        }else{
            $top=100
        }

        $Folders = Get-AllMailFolders -AccessToken $AccessToken  -MailboxName $MailboxName   
        $fldIndex = @{};
        foreach($folder in $Folders){
            $laFid = $folder.PR_ENTRYID.substring(44,44)
            $fldIndex.Add($laFid,$folder);
        }  
        $EndPoint = "https://graph.microsoft.com/v1.0/users"
        $RequestURL = $EndPoint + "('$MailboxName')/MailFolders('$FolderName')/messages?`$Top=" + $top + "&$`select=" + $SelectList
        if(![String]::IsNullOrEmpty($filter)){
            $RequestURL = $EndPoint + "('$MailboxName')/MailFolders('$FolderName')/messages?`$Top=" + $top + "&`$select=" + $SelectList + "&`$filter=" + $filter
        }
        $LastActiveParentEntryId = Get-EXRTaggedProperty -DataType "Binary" -Id "0x348A"
        $PropList = @()
        $PropList += $LastActiveParentEntryId
        $Props = Get-ExtendedPropList -PropertyList $PropList 
        $RequestURL += "&`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")" 
   
        $headers = @{
            'Authorization' = "Bearer $AccessToken"
            'AnchorMailbox' = "$MailboxName"
        }
        $MessageEnumCount =0;
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
                    Expand-MessageProperties -Item $Item     
                    Expand-ExtendedProperties -Item $Item              
                    $MessageEnumCount++
                    Add-Member -InputObject $Item -NotePropertyName "MessageCount" -NotePropertyValue $MessageEnumCount
                    if($Item.LastActiveParentEntryId){
                        if($fldIndex.ContainsKey($Item.LastActiveParentEntryId)){
                            $Item | Add-Member -Name "LastActiveParentFolderPath" -Value $fldIndex[$Item.LastActiveParentEntryId].FolderPath -MemberType NoteProperty
                            $Item | Add-Member -Name "LastActiveParentFolder" -Value $fldIndex[$Item.LastActiveParentEntryId] -MemberType NoteProperty
                        }
                    }
                    Write-Output $Item
                    if($MessageEnumCount -gt $MessageCount){break}                
                }
                $QueryResults = $null
               
                if($MessageEnumCount -lt $MessageCount){
                    $RequestURL = $Results.'@odata.nextlink'
                }
            }
        } until (!($RequestURL))        
        
    }
}

function Get-RecoverableItemsPurges{
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
        [String]
        $AccessToken,
        [Parameter(Position = 7, Mandatory = $false)]
        [Int32]
        $MessageCount=10,
        [Parameter(Position = 8, Mandatory = $false)]
        [String]
        $filter,
        [Parameter(Position = 7, Mandatory = $false)]
        [String]
        $SelectList = "sender,Subject,receivedDateTime,lastModifiedDateTime,internetmessageid,parentFolderId"
    )

    process {
        return Get-RecoverableItems -MailboxName $MailboxName -FolderName "RecoverableItemsPurges" -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -AutoPrompt:$AutoPrompt.IsPresent -AccessToken $AccessToken -MessageCount $MessageCount -filter $filter -SelectList $SelectList     
        
    }
}

function Get-RecoverableItemsDeletions{
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
        [String]
        $AccessToken,
        [Parameter(Position = 7, Mandatory = $false)]
        [Int32]
        $MessageCount=10,
        [Parameter(Position = 8, Mandatory = $false)]
        [String]
        $filter,
        [Parameter(Position = 9, Mandatory = $false)]
        [String]
        $SelectList = "sender,Subject,receivedDateTime,lastModifiedDateTime,internetmessageid,parentFolderId"
    )

    process {
        return Get-RecoverableItems -MailboxName $MailboxName -FolderName "RecoverableItemsDeletions" -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -AutoPrompt:$AutoPrompt.IsPresent -AccessToken $AccessToken -MessageCount $MessageCount -filter $filter -SelectList $SelectList        	
	}
}

function Invoke-ExportItem{
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
        [String]
        $AccessToken,
        [Parameter(Position = 7, Mandatory = $false)]
        [psobject]
        $Item

    )

    process {
        $prompt = $true
        if($AutoPrompt.IsPresent){
            $prompt = $false
        }        
        $EndPoint = "https://graph.microsoft.com/v1.0/users"
        $RequestURL = $EndPoint + "('$MailboxName')/messages/" + $Item.id + "/`$value"
        if([String]::IsNullOrEmpty($AccessToken)){
            $AccessToken = Get-AccessTokenForGraph -MailboxName $Mailboxname -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -Prompt:$prompt
        }        
        $headers = @{
            'Authorization' = "Bearer $AccessToken"
            'AnchorMailbox' = "$MailboxName"
        }
        return (Invoke-RestMethod -Method Get -Uri $RequestURL -UserAgent "GraphBasicsPs101" -Headers $headers)     
        
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

function Get-AllMailFolders{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$MailboxName,		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
        $AccessToken,        
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
	Begin
	{
        
        if([String]::IsNullOrEmpty($AccessToken)){
            $AccessToken = Get-AccessTokenForGraph -MailboxName $Mailboxname -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -Prompt:$prompt
        }        
        $EndPoint = "https://graph.microsoft.com/v1.0/users"
        $RequestURL = $EndPoint + "('$MailboxName')/MailFolders/msgfolderroot/childfolders?`$Top=1000"
        $PREntryId = Get-EXRTaggedProperty -DataType "Binary" -Id "0xfff"
        $PropList = @()
        $PropList += $PREntryId
        $Props = Get-ExtendedPropList -PropertyList $PropList 
        $RequestURL += "&`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")" 
		$FldIndex = New-Object Collections.Hashtable ([StringComparer]::CurrentCulture) 
		$BatchItems = @()
		do
		{
            $headers = @{
                'Authorization' = "Bearer $AccessToken"
                'AnchorMailbox' = "$MailboxName"
            }
            $Folders = (Invoke-RestMethod -Method Get -Uri $RequestURL -UserAgent "GraphBasicsPs101" -Headers $headers).value  
			foreach ($Folder in $Folders)
			{				
				$Folder | Add-Member -NotePropertyName FolderPath -NotePropertyValue ("\\" + $Folder.DisplayName)
				$folderId = $Folder.Id.ToString()
				Add-Member -InputObject $Folder -NotePropertyName FolderRestURI -NotePropertyValue ($EndPoint + "('$MailboxName')/MailFolders('$folderId')")
				Expand-ExtendedProperties -Item $Folder
				$FldIndex.Add($Folder.Id,$Folder.FolderPath)
				Write-Output $Folder
				if ($Folder.ChildFolderCount -gt 0)
				{
					$BatchItems += $Folder
					if($BatchItems.Count -eq 4){
						Get-AllChildFoldersBatch -BatchItems $BatchItems -MailboxName $MailboxName -AccessToken $AccessToken -PropList $PropList -FldIndex $FldIndex
						$BatchItems = @()
					}
				}
			}
			$RequestURL = $JSONOutput.'@odata.nextLink'
		}
		while (![String]::IsNullOrEmpty($RequestURL))
		if($BatchItems.Count -gt 0){
			Get-AllChildFoldersBatch -BatchItems $BatchItems -MailboxName $MailboxName -AccessToken $AccessToken -PropList $PropList -FldIndex $FldIndex	
        }	
        	
	}
}


function Get-AllChildFoldersBatch
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[PSCustomObject]
		$BatchItems,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[PSCustomObject]
		$PropList,

		[Parameter(Position = 4, Mandatory = $false)]
		[PSCustomObject]
		$MailboxName,
		
		[Parameter(Position = 5, Mandatory = $false)]
		[PSCustomObject]
		$FldIndex

	)
	Begin
	{
		$ChildFolders = Get-BatchItems -Items $BatchItems -MailboxName $MailboxName -AccessToken $AccessToken -PropList $PropList -URLString ("/users" + "('" + $MailboxName + "')" + "/MailFolders") -ChildFolders		
		$BatchItems = @()
		for($intcnt=0;$intcnt -lt $ChildFolders.value.Count;$intcnt++){
				$Child = $ChildFolders.value[$intcnt]					
				if($Child -ne $null){
					$ParentId = $Child.parentFolderId
					$ChildFldPath = ($FldIndex[$ParentId] + "\" + $Child.displayName)
					$Child | Add-Member -NotePropertyName FolderPath -NotePropertyValue $ChildFldPath
					$FldIndex.Add($Child.Id,$ChildFldPath)
					$folderId = $Child.Id.ToString()
					Add-Member -InputObject $Child -NotePropertyName FolderRestURI -NotePropertyValue ($EndPoint + "('$MailboxName')/MailFolders('$folderId')")
					Expand-ExtendedProperties -Item $Child
					Write-Output $Child
					if ($Child.ChildFolderCount -gt 0)
					{
						$BatchItems += $Child
					}
					if($BatchItems.Count -gt 4){
						Get-AllChildFoldersBatch -BatchItems $BatchItems -MailboxName $MailboxName -AccessToken $AccessToken -PropList $PropList -FldIndex $FldIndex
						$BatchItems = @()
					}	
				}
		}
		if($BatchItems.Count -gt 0){
			Get-AllChildFoldersBatch -BatchItems $BatchItems -MailboxName $MailboxName -AccessToken $AccessToken -PropList $PropList -FldIndex $FldIndex	
			$BatchItems = @()
		}
		
	}
}
function Get-BatchItems {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $false)]
        [string]
        $MailboxName,
		
        [Parameter(Position = 1, Mandatory = $false)]
        [psobject]
        $AccessToken,

        [Parameter(Position = 2, Mandatory = $false)]
        [psobject]
        $Items,
        [Parameter(Position = 3, Mandatory = $false)]
        [psobject]
        $SelectProperties,
        [Parameter(Position = 4, Mandatory = $false)]
        [psobject]
        $PropList,		
		  
        [Parameter(Position = 5, Mandatory = $false)]
        [psobject]
        $URLString,

        [Parameter(Position = 6, Mandatory = $false)] 
        [switch]
        $RestrictProps,
		
        [Parameter(Position = 7, Mandatory = $false)] 
        [switch]
        $ChildFolders

		
    )
    Process {
        $ReturnCollection = @()
        $RequestURL = "https://graph.microsoft.com/v1.0/`$batch"
        $RequestContent = "{`r`n`"requests`": ["
        $itemCount = 1
        foreach ($Item in $Items) {
            $ItemURI = $URLString + "('" + $Item.Id + "')"
            if ($ChildFolders.IsPresent) {
                $ItemURI += "/childfolders/?`$Top=100"
            }
            $boolSelectProp = $false
            if ($RestrictProps.IsPresent) {
                if (![String]::IsNullOrEmpty($SelectProperties)) {
                    $ItemURI += "/?" + $SelectProperties
                    $boolSelectProp = $true
                }
            }
            if ($PropList -ne $null) {
                $Props = Get-ExtendedPropList -PropertyList $PropList 
                if ($boolSelectProp) {
                    $ItemURI += "`&"
                }
                else {
                    if (!$ItemURI.Contains("/?")) {
                        $ItemURI += "/?"
                    }                    
                }
                if ($ChildFolders.IsPresent) {
                    $ItemURI += "`&"
                }
                $ItemURI += "`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
            }
            if ($itemCount -ne 1) {$RequestContent += ",`r`n"}  
            #if($itemCount -eq 1){
            $RequestContent += "{`r`n`"id`": `"" + $itemCount + "`",`r`n`"method`": `"GET`","
            # }else{
            #   $RequestContent += "{`r`n`"id`": `"" + $itemCount + "`",`r`n`"dependsOn`": [ `"" + ($itemCount-1) +"`" ],`r`n`"method`": `"GET`","
            #}
            
            $RequestContent += "`"url`": `"" + $ItemURI + "`"`r`n }"
            $itemCount++
        }
        $RequestContent += "`r`n]`r`n}"
        $headers = @{
            'Authorization' = "Bearer $AccessToken"
            'AnchorMailbox' = "$MailboxName"
        }
        $JSONOutput = (Invoke-RestMethod -Method POST -Uri $RequestURL -UserAgent "GraphBasicsPs101" -Headers $headers -Body $RequestContent -ContentType "application/json" )   
        foreach ($BatchItem in $JSONOutput.responses) {
            Expand-MessageProperties -Item $BatchItem.Body
            $ReturnCollection += $BatchItem.Body
        }
        return $ReturnCollection
	
    }
}

function Expand-MessageProperties
{
	[CmdletBinding()] 
    param (
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
        $Item
	)
	
 	process
	{
        try{
            if ([bool]($Item.PSobject.Properties.name -match "sender"))
            {
                $SenderProp = $Item.sender
                if ([bool]($SenderProp.PSobject.Properties.name -match "emailaddress"))
                {
                    Add-Member -InputObject $Item -NotePropertyName "SenderEmailAddress" -NotePropertyValue $SenderProp.emailaddress.address
                    Add-Member -InputObject $Item -NotePropertyName "SenderName" -NotePropertyValue $SenderProp.emailaddress.name
                }
                
            }
            if ([bool]($Item.PSobject.Properties.name -match "InternetMessageHeaders"))
            {
                $IndexedHeaders = New-Object 'system.collections.generic.dictionary[string,string]'
                foreach($header in $Item.InternetMessageHeaders){
                    if(!$IndexedHeaders.ContainsKey($header.name)){
                        $IndexedHeaders.Add($header.name,$header.value)
                    }
                }
                 Add-Member -InputObject $Item -NotePropertyName "IndexedInternetMessageHeaders" -NotePropertyValue $IndexedHeaders
            }
            
        }
        catch{

        }
    }
}

function Expand-ExtendedProperties
{
	[CmdletBinding()] 
    param (
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$Item
	)
	
 	process
	{
		if ($Item.singleValueExtendedProperties -ne $null)
		{
			foreach ($Prop in $Item.singleValueExtendedProperties)
			{
				Switch ($Prop.Id)
				{
                    "Binary 0x3019" {
                        Add-Member -InputObject $Item -NotePropertyName "PR_POLICY_TAG" -NotePropertyValue ([System.GUID]([Convert]::FromBase64String($Prop.Value)))
                    }
                    "Binary 0x1013"{
                        Add-Member -InputObject $Item -NotePropertyName "PR_BODY_HTML" -NotePropertyValue ([System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($Prop.Value)))
                    }
                    "Binary 0xfff" {
                        Add-Member -InputObject $Item -NotePropertyName "PR_ENTRYID" -NotePropertyValue ([System.BitConverter]::ToString([Convert]::FromBase64String($Prop.Value)).Replace("-",""))
                    }
                    "Binary 0x301B" {
                        $fileTime = [BitConverter]::ToInt64([Convert]::FromBase64String($Prop.Value), 4);
                        $StartTime = [DateTime]::FromFileTime($fileTime)
                        Add-Member -InputObject $Item -NotePropertyName "PR_START_DATE_ETC" -NotePropertyValue $StartTime
                    }
                    "Binary 0x348A"{                            
                        Add-Member  -InputObject $Item -NotePropertyName "LastActiveParentEntryId" -NotePropertyValue ([System.BitConverter]::ToString([Convert]::FromBase64String($Prop.Value)).Replace("-",""))
                    }
                    "Integer 0x301D" {
                        Add-Member -InputObject $Item -NotePropertyName "PR_RETENTION_FLAGS" -NotePropertyValue $Prop.Value
                    }
                    "Integer 0x301A" {
                        Add-Member -InputObject $Item -NotePropertyName "PR_RETENTION_PERIOD" -NotePropertyValue $Prop.Value
                    }
                    "SystemTime 0x301C" {
                        Add-Member -InputObject $Item -NotePropertyName "PR_RETENTION_DATE" -NotePropertyValue ([DateTime]::Parse($Prop.Value))
                    }
		             "String {403fc56b-cd30-47c5-86f8-ede9e35a022b} Name ComplianceTag" {
                        Add-Member -InputObject $Item -NotePropertyName "ComplianceTag" -NotePropertyValue $Prop.Value
                    }
                    "Integer {23239608-685D-4732-9C55-4C95CB4E8E33} Name InferenceClassificationResult" {
                        Add-Member -InputObject $Item -NotePropertyName "InferenceClassificationResult" -NotePropertyValue $Prop.Value
                    }
                    "Binary {e49d64da-9f3b-41ac-9684-c6e01f30cdfa} Name TeamChatFolderEntryId" {
                        Add-Member -InputObject $Item -NotePropertyName "TeamChatFolderEntryId" -NotePropertyValue $Prop.Value
                    }
                    "Integer 0xe08" {
                        Add-Member -InputObject $Item -NotePropertyName "Size" -NotePropertyValue $Prop.Value
                    }
                    "Long 0x66B3" {
                        Add-Member -InputObject $Item -NotePropertyName "FolderSize" -NotePropertyValue $Prop.Value
                    }
		            "String 0x7d" {
                        Add-Member -InputObject $Item -NotePropertyName "PR_TRANSPORT_MESSAGE_HEADERS" -NotePropertyValue $Prop.Value
                    }
                    "SystemTime 0xF02"{
                        Add-Member -InputObject $Item -NotePropertyName "PR_RENEWTIME" -NotePropertyValue ([DateTime]::Parse($Prop.Value))
                    }
                    "SystemTime 0xF01"{
                        Add-Member -InputObject $Item -NotePropertyName "PR_RENEWTIME2" -NotePropertyValue ([DateTime]::Parse($Prop.Value))
                    }
                    "String 0x66b5"{
                          Add-Member -InputObject $Item -NotePropertyName "PR_Folder_Path" -NotePropertyValue $Prop.Value.Replace("￾","\") -Force
                    }
                    "Short 0x3a4d"{
                          Add-Member -InputObject $Item -NotePropertyName "PR_Gender" -NotePropertyValue $Prop.Value -Force
                    }
                    "String 0x001a"{
                          Add-Member -InputObject $Item -NotePropertyName "PR_MESSAGE_CLASS" -NotePropertyValue $Prop.Value -Force
                    }
                    "Integer 0x6638"{
                          Add-Member -InputObject $Item -NotePropertyName "PR_FOLDER_CHILD_COUNT" -NotePropertyValue $Prop.Value -Force
                    }
                    "Integer 0x1081"{
                        Add-Member -InputObject $Item -NotePropertyName "PR_LAST_VERB_EXECUTED" -NotePropertyValue $Prop.Value -Force
                        $verbHash = Get-LASTVERBEXECUTEDHash;
                        if($verbHash.ContainsKey($Prop.Value)){
                            Add-Member -InputObject $Item -NotePropertyName "PR_LAST_VERB_EXECUTED_DisplayName" -NotePropertyValue $verbHash[$Prop.Value]
                        } 
                    }   
                    "SystemTime 0x1082"{
                        Add-Member -InputObject $Item -NotePropertyName "PR_LAST_VERB_EXECUTION_TIME" -NotePropertyValue ([DateTime]::Parse($Prop.Value))
                    }    
                                 
                    "String {00062008-0000-0000-C000-000000000046} Name EntityExtraction/Sentiment1.0" {
                          Invoke-EXRProcessSentiment -Item $Item -JSONData $Prop.Value
                    }
                    "Integer {00062002-0000-0000-c000-000000000046} Id 0x8213" {
                        Add-Member -InputObject $Item -NotePropertyName "AppointmentDuration" -NotePropertyValue $Prop.Value -Force
                    }
                    default {Write-Host $Prop.Id}
                }
            }
        }
    }
}

function Invoke-RestoreItemToLastActiveParentFolderId{
    [CmdletBinding()] 
    param (
		[Parameter(Position = 1, Mandatory = $true)]
		[psobject]
        $Item,
        [Parameter(Position = 2, Mandatory = $false)]
		[psobject]
        $AccessToken,
        [Parameter(Position = 3, Mandatory = $true)]
        [String]
        $MailboxName,
        [Parameter(Position = 4, Mandatory = $false)]
        [String]
        $ClientId,
        [Parameter(Position = 5, Mandatory = $false)]
        [String]
        $RedirectURI = "urn:ietf:wg:oauth:2.0:oob",
        [Parameter(Position = 6, Mandatory = $false)]
        [String]
        $scopes = "User.Read.All Mail.Read",
        [Parameter(Position = 7, Mandatory = $false)]
        [switch]
        $AutoPrompt	
	)	
 	process
	{
        $prompt = $true
        if($AutoPrompt.IsPresent){
            $prompt = $false
        }
        if([String]::IsNullOrEmpty($AccessToken)){
            $AccessToken = Get-AccessTokenForGraph -MailboxName $Mailboxname -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -Prompt:$prompt
        }   
        $EndPoint = "https://graph.microsoft.com/v1.0/users"
        $RequestURL = $EndPoint + "('$MailboxName')/MailFolders('" + $Item.parentFolderId + "')/messages/" + $Item.id + "/move"
        $headers = @{
            'Authorization' = "Bearer $AccessToken"
            'AnchorMailbox' = "$MailboxName"
        }
        $PostConent = "{`"destinationId`": `"" + $Item.LastActiveParentFolder.id + "`"}"
        return (Invoke-RestMethod -Method POST -Uri $RequestURL -UserAgent "GraphBasicsPs101" -Headers $headers -Body $PostConent -ContentType "application/json" )   
    }
}

