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

function Get-SharedFolderFromPath {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string]
        $FolderPath,		
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $TargetMailbox,
        [Parameter(Position = 2, Mandatory = $true)]
        [String]
        $LogonMailbox,
        [Parameter(Position = 3, Mandatory = $false)]
        [String]
        $ClientId,
        [Parameter(Position = 4, Mandatory = $false)]
        [String]
        $RedirectURI = "urn:ietf:wg:oauth:2.0:oob",
        [Parameter(Position = 5, Mandatory = $false)]
        [String]
        $scopes = "User.Read.All Mail.Read Mail.Read.Shared",
        [Parameter(Position = 6, Mandatory = $false)]
        [switch]
        $AutoPrompt	
    )

    process {
        
        $prompt = $true
        if($AutoPrompt.IsPresent){
            $prompt = $false
        }
        $EndPoint = "https://graph.microsoft.com/v1.0/users"
        $RequestURL = $EndPoint + "('$TargetMailbox')/MailFolders('MsgFolderRoot')/childfolders?"
        $AccessToken = Get-AccessTokenForGraph -MailboxName $LogonMailbox -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -Prompt:$prompt
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
                'AnchorMailbox' = "$TargetMailbox"
            }
            $RequestURL = $RequestURL += "`&`$filter=DisplayName eq '$FolderName'"
            $tfTargetFolder = (Invoke-RestMethod -Method Get -Uri $RequestURL -UserAgent "GraphBasicsPs101" -Headers $headers).value  
            if ($tfTargetFolder.displayname -match $FolderName) {
                $folderId = $tfTargetFolder.Id.ToString()
                $RequestURL = $EndPoint + "('$TargetMailbox')/MailFolders('$folderId')/childfolders?"
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

function Get-SharedWellKnownFolder {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string]
        $FolderName,		
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $TargetMailbox,
        [Parameter(Position = 2, Mandatory = $true)]
        [String]
        $LogonMailbox,
        [Parameter(Position = 3, Mandatory = $false)]
        [String]
        $ClientId,
        [Parameter(Position = 4, Mandatory = $false)]
        [String]
        $RedirectURI = "urn:ietf:wg:oauth:2.0:oob",
        [Parameter(Position = 5, Mandatory = $false)]
        [String]
        $scopes = "User.Read.All Mail.Read Mail.Read.Shared",
        [Parameter(Position = 6, Mandatory = $false)]
        [switch]
        $AutoPrompt,
        [Parameter(Position = 7, Mandatory = $false)]
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
        $RequestURL = $EndPoint + "('$TargetMailbox')/MailFolders('$FolderName')?"
        if(!$AccessToken){
            $AccessToken = Get-AccessTokenForGraph -MailboxName $LogonMailbox -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -Prompt:$prompt
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
                'AnchorMailbox' = "$TargetMailbox"
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


function Get-SharedFolderFromId {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string]
        $FolderId,		
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $TargetMailbox,
        [Parameter(Position = 2, Mandatory = $true)]
        [String]
        $LogonMailbox,
        [Parameter(Position = 3, Mandatory = $false)]
        [String]
        $ClientId,
        [Parameter(Position = 4, Mandatory = $false)]
        [String]
        $RedirectURI = "urn:ietf:wg:oauth:2.0:oob",
        [Parameter(Position = 5, Mandatory = $false)]
        [String]
        $scopes = "User.Read.All Mail.Read Mail.Read.Shared",
        [Parameter(Position = 6, Mandatory = $false)]
        [switch]
        $AutoPrompt,
        [Parameter(Position = 7, Mandatory = $false)]
        [psobject]
        $PropList,
        [Parameter(Position = 8, Mandatory = $false)]
        [String]
        $AccessToken		
    )

    process {
        
        $prompt = $true
        if($AutoPrompt.IsPresent){
            $prompt = $false
        }
        $EndPoint = "https://graph.microsoft.com/v1.0/users"
        $RequestURL = $EndPoint + "('$TargetMailbox')/MailFolders('$FolderId')?"
        if(!$AccessToken){
            $AccessToken = Get-AccessTokenForGraph -MailboxName $LogonMailbox -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -Prompt:$prompt
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
                'AnchorMailbox' = "$TargetMailbox"
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

function Get-SharedLastEmail{
    [CmdletBinding()]
    param (
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $TargetMailbox,
        [Parameter(Position = 2, Mandatory = $true)]
        [String]
        $LogonMailbox,
        [Parameter(Position = 3, Mandatory = $false)]
        [String]
        $FolderName = "Inbox",
        [Parameter(Position = 4, Mandatory = $false)]
        [String]
        $ClientId,
        [Parameter(Position = 5, Mandatory = $false)]
        [String]
        $RedirectURI = "urn:ietf:wg:oauth:2.0:oob",
        [Parameter(Position = 6, Mandatory = $false)]
        [String]
        $scopes = "User.Read.All Mail.Read Mail.Read.Shared",
        [Parameter(Position = 7, Mandatory = $false)]
        [switch]
        $AutoPrompt,	
        [Parameter(Position = 8, Mandatory = $false)]
        [String]
        $AccessToken,
        [Parameter(Position = 9, Mandatory = $false)]
        [Int32]
        $MessageCount=1,
        [Parameter(Position = 10, Mandatory = $false)]
        [String]
        $filter,
        [Parameter(Position = 11, Mandatory = $false)]
        [String]
        $SelectList = "sender,Subject,receivedDateTime,lastModifiedDateTime,internetmessageid,parentFolderId",
        [Parameter(Position = 12, Mandatory = $false)]
        [switch]
        $FocusedInbox,
        [Parameter(Position = 13, Mandatory = $false)]
        [switch]
        $Other
    )

    process {
        $prompt = $true
        if($AutoPrompt.IsPresent){
            $prompt = $false
        }
        if([String]::IsNullOrEmpty($AccessToken)){
            $AccessToken = Get-AccessTokenForGraph -MailboxName $LogonMailbox -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -Prompt:$prompt
        }     
        if($MessageCount -lt 100){
            $top=$MessageCount
        }else{
            $top=100
        }
        $EndPoint = "https://graph.microsoft.com/v1.0/users"
        $RequestURL = $EndPoint + "('$TargetMailbox')/MailFolders('$FolderName')/messages?`$Top=" + $top + "&$`select=" + $SelectList
        if(![String]::IsNullOrEmpty($filter)){
            $RequestURL = $EndPoint + "('$TargetMailbox')/MailFolders('$FolderName')/messages?`$Top=" + $top + "&`$select=" + $SelectList + "&`$filter=" + $filter
        }  
        $headers = @{
            'Authorization' = "Bearer $AccessToken"
            'AnchorMailbox' = "$TargetMailbox"
        }
        if($FocusedInbox.IsPresent){
            $RequestURL +="&`$orderby=InferenceClassification, createdDateTime DESC&`$filter=InferenceClassification eq 'Focused'"
        }
        if($Other.IsPresent){
            $RequestURL +="&`$orderby=InferenceClassification, createdDateTime DESC&`$filter=InferenceClassification eq 'Other'"
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
                    $MessageEnumCount++
                    Expand-MessageProperties -Item $Item     
                    Expand-ExtendedProperties -Item $Item              
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

function Send-SharedMessage
{
	[CmdletBinding()]
	param (
        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $TargetMailbox,
        [Parameter(Position = 2, Mandatory = $true)]
        [String]
        $LogonMailbox,
        [Parameter(Position = 3, Mandatory = $false)]
        [String]
        $ClientId,
        [Parameter(Position = 5, Mandatory = $false)]
        [String]
        $RedirectURI = "urn:ietf:wg:oauth:2.0:oob",
        [Parameter(Position = 6, Mandatory = $false)]
        [String]
        $scopes = "User.Read.All Mail.Send Mail.Send.Shared",
        [Parameter(Position = 7, Mandatory = $false)]
        [switch]
        $AutoPrompt,	
        [Parameter(Position = 8, Mandatory = $false)]
        [String]
        $AccessToken,	
		
		[Parameter(Position = 9, Mandatory = $true)]
		[String]
		$Subject,
		
		[Parameter(Position = 10, Mandatory = $false)]
		[String]
		$Body,
		
		[Parameter(Position = 11, Mandatory = $false)]
		[psobject]
        $From,       
		
		[Parameter(Position = 12, Mandatory = $false)]
		[psobject]
		$Attachment,
		
		[Parameter(Position = 13, Mandatory = $false)]
		[psobject]
		$To,

		[Parameter(Position = 14, Mandatory = $false)]
		[psobject]
		$CC,
		
		[Parameter(Position = 15, Mandatory = $false)]
		[psobject]
		$BCC,
		
		[Parameter(Position = 16, Mandatory = $false)]
		[switch]
		$DontSaveToSentItems,
		
		[Parameter(Position = 17, Mandatory = $false)]
		[switch]
		$ShowRequest,
		
		[Parameter(Position = 18, Mandatory = $false)]
		[switch]
		$RequestReadRecipient,
		
		[Parameter(Position = 19, Mandatory = $false)]
		[switch]
        $RequestDeliveryRecipient,        
        		
		[Parameter(Position = 20, Mandatory = $false)]
		[switch]
		$SendOnBehalf,
		
		[Parameter(Position = 21, Mandatory = $false)]
		[psobject]
        $ReplyTo,
        
        [Parameter(Position = 22, Mandatory = $false)]
		[switch]
		$SaveToLogonMailbox
	)
	Begin
	{
        if([String]::IsNullOrEmpty($From)){
            if($SendOnBehalf){
                $From = $LogonMailbox
            }else{
                $From = $TargetMailbox
            }
           
        }
        if([String]::IsNullOrEmpty($SenderEmailAddress)){
            if($SendOnBehalf){
                $SenderEmailAddress = $TargetMailbox
            }
        }
        $prompt = $true
        if($AutoPrompt.IsPresent){
            $prompt = $false
        }
        if([String]::IsNullOrEmpty($AccessToken)){
            $AccessToken = Get-AccessTokenForGraph -MailboxName $LogonMailbox -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -Prompt:$prompt
        }     
		$SaveToSentFolder = "true"
		if ($DontSaveToSentItems.IsPresent)
		{
			$SaveToSentFolder = "false"
		}
		$Attachments = @()
		if(![String]::IsNullOrEmpty($Attachment)){
			$Attachments += (Resolve-Path $Attachment).Path
		}
		$ToRecipients = @()
		if(![String]::IsNullOrEmpty($To)){
			$ToRecipients += (New-RESTEmailAddress -Address $To)
		}
		$CCRecipients = @()
		if(![String]::IsNullOrEmpty($CC)){
		   $CCRecipients += (New-RESTEmailAddress -Address $CC)
		}
		$BCCRecipients = @()
		if(![String]::IsNullOrEmpty($BCC)){
		   $BCCRecipients += (New-RESTEmailAddress -Address $BCC)
		}
		if(![String]::IsNullOrEmpty($From)){
			$From = (New-RESTEmailAddress -Address $From)
        }
        if(![String]::IsNullOrEmpty($SenderEmailAddress)){
			$SenderEmailAddress = (New-RESTEmailAddress -Address $SenderEmailAddress)
		}
		$NewMessage = Get-MessageJSONFormat -Subject $Subject -Body $Body.Replace("`"","\`"") -SenderEmailAddress $SenderEmailAddress -From $From -Attachments $Attachments -ReferanceAttachments $ReferanceAttachments -ToRecipients $ToRecipients -SentDate $SentDate -ExPropList $ExPropList -CcRecipients $CCRecipients -bccRecipients $BCCRecipients -StandardPropList $StandardPropList -SaveToSentItems $SaveToSentFolder -SendMail -ReplyTo $ReplyTo -RequestReadRecipient $RequestReadRecipient.IsPresent -RequestDeliveryRecipient $RequestDeliveryRecipient.IsPresent
		if ($ShowRequest.IsPresent)
		{
			write-host $NewMessage
		}
        $EndPoint = "https://graph.microsoft.com/v1.0/users"
        $RequestURL = $EndPoint + "('$TargetMailbox')/sendMail"
        if($SendOnBehalf){
            $RequestURL = "https://graph.microsoft.com/v1.0/me/sendMail"
        }
        if($SaveToLogonMailbox){
            $RequestURL = "https://graph.microsoft.com/v1.0/me/sendMail"
        }
        $headers = @{
            'Authorization' = "Bearer $AccessToken"
            'AnchorMailbox' = "$TargetMailbox"
        }
        $Results = (Invoke-RestMethod -Method POST -Uri $RequestURL -UserAgent "GraphBasicsPs101" -Headers $headers -Body $NewMessage -ContentType "application/JSON") 
        return $Results
	}
}

function New-RESTEmailAddress
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$Name,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$Address
	)
	Begin
	{
		$EmailAddress = "" | Select-Object Name, Address
		if ([String]::IsNullOrEmpty($Name))
		{
			$EmailAddress.Name = $Address
		}
		else
		{
			$EmailAddress.Name = $Name
		}
		$EmailAddress.Address = $Address
		return, $EmailAddress
	}
}

function Get-MessageJSONFormat {
    [CmdletBinding()]
    param (
        [Parameter(Position = 1, Mandatory = $false)]
        [String]
        $Subject,
		
        [Parameter(Position = 2, Mandatory = $false)]
        [String]
        $Body,
		
        [Parameter(Position = 3, Mandatory = $false)]
        [psobject]
        $SenderEmailAddress,

        [Parameter(Position = 4, Mandatory = $false)]
        [psobject]
        $From,
		
        [Parameter(Position = 5, Mandatory = $false)]
        [psobject]
        $Attachments,
		
        [Parameter(Position = 5, Mandatory = $false)]
        [psobject]
        $ReferanceAttachments,
		
        [Parameter(Position = 6, Mandatory = $false)]
        [psobject]
        $ToRecipients,
		
        [Parameter(Position = 7, Mandatory = $false)]
        [psobject]
        $CcRecipients,
		
        [Parameter(Position = 7, Mandatory = $false)]
        [psobject]
        $bccRecipients,
		
        [Parameter(Position = 8, Mandatory = $false)]
        [psobject]
        $SentDate,
		
        [Parameter(Position = 9, Mandatory = $false)]
        [psobject]
        $StandardPropList,
		
        [Parameter(Position = 10, Mandatory = $false)]
        [psobject]
        $ExPropList,
		
        [Parameter(Position = 11, Mandatory = $false)]
        [switch]
        $ShowRequest,
		
        [Parameter(Position = 12, Mandatory = $false)]
        [String]
        $SaveToSentItems,
		
        [Parameter(Position = 13, Mandatory = $false)]
        [switch]
        $SendMail,
		
        [Parameter(Position = 14, Mandatory = $false)]
        [psobject]
        $ReplyTo,
		
        [Parameter(Position = 17, Mandatory = $false)]
        [bool]
        $RequestReadRecipient,
		
        [Parameter(Position = 18, Mandatory = $false)]
        [bool]
        $RequestDeliveryRecipient
    )
    Begin {
        $NewMessage = "{" + "`r`n"
        if ($SendMail.IsPresent) {
            $NewMessage += "  `"Message`" : {" + "`r`n"
        }
        if (![String]::IsNullOrEmpty($Subject)) {
            $NewMessage += "`"Subject`": `"" + $Subject + "`"" + "`r`n"
        }
        if ($SenderEmailAddress -ne $null) {
            if ($NewMessage.Length -gt 5) { $NewMessage += "," }
            $NewMessage += "`"Sender`":{" + "`r`n"
            $NewMessage += " `"EmailAddress`":{" + "`r`n"
            $NewMessage += "  `"Name`":`"" + $SenderEmailAddress.Name + "`"," + "`r`n"
            $NewMessage += "  `"Address`":`"" + $SenderEmailAddress.Address + "`"" + "`r`n"
            $NewMessage += "}}" + "`r`n"
        }
        if ($From -ne $null) {
            if ($NewMessage.Length -gt 5) { $NewMessage += "," }
            $NewMessage += "`"From`":{" + "`r`n"
            $NewMessage += " `"EmailAddress`":{" + "`r`n"
            $NewMessage += "  `"Name`":`"" + $From.Name + "`"," + "`r`n"
            $NewMessage += "  `"Address`":`"" + $From.Address + "`"" + "`r`n"
            $NewMessage += "}}" + "`r`n"
        }
        if (![String]::IsNullOrEmpty($Body)) {
            if ($NewMessage.Length -gt 5) { $NewMessage += "," }
            $NewMessage += "`"Body`": {" + "`r`n"
            $NewMessage += "`"ContentType`": `"HTML`"," + "`r`n"
            $NewMessage += "`"Content`": `"" + $Body + "`"" + "`r`n"
            $NewMessage += "}" + "`r`n"
        }
		
        $toRcpcnt = 0;
        if ($ToRecipients -ne $null) {
            if ($NewMessage.Length -gt 5) { $NewMessage += "," }
            $NewMessage += "`"ToRecipients`": [ " + "`r`n"
            foreach ($EmailAddress in $ToRecipients) {
                if ($toRcpcnt -gt 0) {
                    $NewMessage += "      ,{ " + "`r`n"
                }
                else {
                    $NewMessage += "      { " + "`r`n"
                }
                $NewMessage += " `"EmailAddress`":{" + "`r`n"
                $NewMessage += "  `"Name`":`"" + $EmailAddress.Name + "`"," + "`r`n"
                $NewMessage += "  `"Address`":`"" + $EmailAddress.Address + "`"" + "`r`n"
                $NewMessage += "}}" + "`r`n"
                $toRcpcnt++
            }
            $NewMessage += "  ]" + "`r`n"
        }
        $ccRcpcnt = 0
        if ($CcRecipients -ne $null) {
            if ($NewMessage.Length -gt 5) { $NewMessage += "," }
            $NewMessage += "`"CcRecipients`": [ " + "`r`n"
            foreach ($EmailAddress in $CcRecipients) {
                if ($ccRcpcnt -gt 0) {
                    $NewMessage += "      ,{ " + "`r`n"
                }
                else {
                    $NewMessage += "      { " + "`r`n"
                }
                $NewMessage += " `"EmailAddress`":{" + "`r`n"
                $NewMessage += "  `"Name`":`"" + $EmailAddress.Name + "`"," + "`r`n"
                $NewMessage += "  `"Address`":`"" + $EmailAddress.Address + "`"" + "`r`n"
                $NewMessage += "}}" + "`r`n"
                $ccRcpcnt++
            }
            $NewMessage += "  ]" + "`r`n"
        }
        $bccRcpcnt = 0
        if ($bccRecipients -ne $null) {
            if ($NewMessage.Length -gt 5) { $NewMessage += "," }
            $NewMessage += "`"BccRecipients`": [ " + "`r`n"
            foreach ($EmailAddress in $bccRecipients) {
                if ($bccRcpcnt -gt 0) {
                    $NewMessage += "      ,{ " + "`r`n"
                }
                else {
                    $NewMessage += "      { " + "`r`n"
                }
                $NewMessage += " `"EmailAddress`":{" + "`r`n"
                $NewMessage += "  `"Name`":`"" + $EmailAddress.Name + "`"," + "`r`n"
                $NewMessage += "  `"Address`":`"" + $EmailAddress.Address + "`"" + "`r`n"
                $NewMessage += "}}" + "`r`n"
                $bccRcpcnt++
            }
            $NewMessage += "  ]" + "`r`n"
        }
        $ReplyTocnt = 0
        if ($ReplyTo -ne $null) {
            if ($NewMessage.Length -gt 5) { $NewMessage += "," }
            $NewMessage += "`"ReplyTo`": [ " + "`r`n"
            foreach ($EmailAddress in $ReplyTo) {
                if ($ReplyTocnt -gt 0) {
                    $NewMessage += "      ,{ " + "`r`n"
                }
                else {
                    $NewMessage += "      { " + "`r`n"
                }
                $NewMessage += " `"EmailAddress`":{" + "`r`n"
                $NewMessage += "  `"Name`":`"" + $EmailAddress.Name + "`"," + "`r`n"
                $NewMessage += "  `"Address`":`"" + $EmailAddress.Address + "`"" + "`r`n"
                $NewMessage += "}}" + "`r`n"
                $ReplyTocnt++
            }
            $NewMessage += "  ]" + "`r`n"
        }
        if ($RequestDeliveryRecipient) {
            $NewMessage += ",`"IsDeliveryReceiptRequested`": true`r`n"
        }
        if ($RequestReadRecipient) {
            $NewMessage += ",`"IsReadReceiptRequested`": true `r`n"
        }
        if ($StandardPropList -ne $null) {
            foreach ($StandardProp in $StandardPropList) {
                if ($NewMessage.Length -gt 5) { $NewMessage += "," }
                switch ($StandardProp.PropertyType) {
                    "Single" {
                        if ($StandardProp.QuoteValue) {
                            $NewMessage += "`"" + $StandardProp.Name + "`": `"" + $StandardProp.Value + "`"" + "`r`n"
                        }
                        else {
                            $NewMessage += "`"" + $StandardProp.Name + "`": " + $StandardProp.Value + "`r`n"
                        }
						
						
                    }
                    "Object" {
                        if ($StandardProp.isArray) {
                            $NewMessage += "`"" + $StandardProp.PropertyName + "`": [ {" + "`r`n"
                        }
                        else {
                            $NewMessage += "`"" + $StandardProp.PropertyName + "`": {" + "`r`n"
                        }
                        $acCount = 0
                        foreach ($PropKeyValue in $StandardProp.PropertyList) {
                            if ($acCount -gt 0) {
                                $NewMessage += ","
                            }
                            $NewMessage += "`"" + $PropKeyValue.Name + "`": `"" + $PropKeyValue.Value + "`"" + "`r`n"
                            $acCount++
                        }
                        if ($StandardProp.isArray) {
                            $NewMessage += "}]" + "`r`n"
                        }
                        else {
                            $NewMessage += "}" + "`r`n"
                        }
						
                    }
                    "ObjectCollection" {
                        if ($StandardProp.isArray) {
                            $NewMessage += "`"" + $StandardProp.PropertyName + "`": [" + "`r`n"
                        }
                        else {
                            $NewMessage += "`"" + $StandardProp.PropertyName + "`": {" + "`r`n"
                        }
                        foreach ($EnclosedStandardProp in $StandardProp.PropertyList) {
                            $NewMessage += "`"" + $EnclosedStandardProp.PropertyName + "`": {" + "`r`n"
                            foreach ($PropKeyValue in $EnclosedStandardProp.PropertyList) {
                                $NewMessage += "`"" + $PropKeyValue.Name + "`": `"" + $PropKeyValue.Value + "`"," + "`r`n"
                            }
                            $NewMessage += "}" + "`r`n"
                        }
                        if ($StandardProp.isArray) {
                            $NewMessage += "]" + "`r`n"
                        }
                        else {
                            $NewMessage += "}" + "`r`n"
                        }
                    }
					
                }
            }
        }
        $atcnt = 0
        $processAttachments = $false
        if ($Attachments -ne $null) { $processAttachments = $true }
        if ($ReferanceAttachments -ne $null) { $processAttachments = $true }
        if ($processAttachments) {
            if ($NewMessage.Length -gt 5) { $NewMessage += "," }
            $NewMessage += "  `"Attachments`": [ " + "`r`n"
            if ($Attachments -ne $null) {
                foreach ($Attachment in $Attachments) {
                    if ($atcnt -gt 0) {
                        $NewMessage += "   ,{" + "`r`n"
                    }
                    else {
                        $NewMessage += "    {" + "`r`n"
                    }
                    if ($Attachment.name) {
                        $NewMessage += "     `"@odata.type`": `"#Microsoft.OutlookServices.FileAttachment`"," + "`r`n"
                        $NewMessage += "     `"Name`": `"" + $Attachment.name + "`"," + "`r`n"
                        $NewMessage += "     `"ContentBytes`": `" " + $Attachment.contentBytes + "`"" + "`r`n"
                    }
                    else {
                        $Item = Get-Item $Attachment

                        $NewMessage += "     `"@odata.type`": `"#Microsoft.OutlookServices.FileAttachment`"," + "`r`n"
                        $NewMessage += "     `"Name`": `"" + $Item.Name + "`"," + "`r`n"
                        $NewMessage += "     `"ContentBytes`": `" " + [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($Attachment)) + "`"" + "`r`n"

                    }
                    $NewMessage += "    } " + "`r`n"
                    $atcnt++
					
                }
            }
            $atcnt = 0
            if ($ReferanceAttachments -ne $null) {
                foreach ($Attachment in $ReferanceAttachments) {
                    if ($atcnt -gt 0) {
                        $NewMessage += "   ,{" + "`r`n"
                    }
                    else {
                        $NewMessage += "    {" + "`r`n"
                    }
                    $NewMessage += "     `"@odata.type`": `"#Microsoft.OutlookServices.ReferenceAttachment`"," + "`r`n"
                    $NewMessage += "     `"Name`": `"" + $Attachment.Name + "`"," + "`r`n"
                    $NewMessage += "     `"SourceUrl`": `"" + $Attachment.SourceUrl + "`"," + "`r`n"
                    $NewMessage += "     `"ProviderType`": `"" + $Attachment.ProviderType + "`"," + "`r`n"
                    $NewMessage += "     `"Permission`": `"" + $Attachment.Permission + "`"," + "`r`n"
                    $NewMessage += "     `"IsFolder`": `"" + $Attachment.IsFolder + "`"" + "`r`n"
                    $NewMessage += "    } " + "`r`n"
                    $atcnt++
                }
            }
            $NewMessage += "  ]" + "`r`n"
        }
		
        if ($ExPropList -ne $null) {
            if ($NewMessage.Length -gt 5) { $NewMessage += "," }
            $NewMessage += "`"SingleValueExtendedProperties`": [" + "`r`n"
            $propCount = 0
            foreach ($Property in $ExPropList) {
                if ($propCount -eq 0) {
                    $NewMessage += "{" + "`r`n"
                }
                else {
                    $NewMessage += ",{" + "`r`n"
                }
                if ($Property.PropertyType -eq "Tagged") {
                    $NewMessage += "`"PropertyId`":`"" + $Property.DataType + " " + $Property.Id + "`", " + "`r`n"
                }
                else {
                    if ($Property.Type -eq "String") {
                        $NewMessage += "`"PropertyId`":`"" + $Property.DataType + " " + $Property.Guid + " Name " + $Property.Id + "`", " + "`r`n"
                    }
                    else {
                        $NewMessage += "`"PropertyId`":`"" + $Property.DataType + " " + $Property.Guid + " Id " + $Property.Id + "`", " + "`r`n"
                    }
                }
                if ($Property.Value -eq "null") {
                    $NewMessage += "`"Value`":null" + "`r`n"
                }
                else {
                    $NewMessage += "`"Value`":`"" + $Property.Value + "`"" + "`r`n"
                }				
                $NewMessage += " } " + "`r`n"
                $propCount++
            }
            $NewMessage += "]" + "`r`n"
        }
        if (![String]::IsNullOrEmpty($SaveToSentItems)) {
            $NewMessage += "}   ,`"SaveToSentItems`": `"" + $SaveToSentItems.ToLower() + "`"" + "`r`n"
        }
        $NewMessage += "}"
        if ($ShowRequest.IsPresent) {
            Write-Host $NewMessage
        }
        return, $NewMessage
    }
}
