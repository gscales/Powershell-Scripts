function Get-AccessTokenForGraph {
    ## Start Code Attribution
    ## Get-AccessTokenForGraph function contains work of the following Authors and should remain with the function if copied into other scripts
	## https://www.lee-ford.co.uk/getting-started-with-microsoft-graph-with-powershell/
	## End Code Attribution
    ## 
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
        $TenantId = (Invoke-WebRequest ("https://login.windows.net/" + $Domain  + "/v2.0/.well-known/openid-configuration") | ConvertFrom-Json).token_endpoint.Split('/')[3]
        Add-Type -AssemblyName System.Web, PresentationFramework, PresentationCore
        $state = Get-Random
        $authURI = "https://login.microsoftonline.com/$TenantId"
        $authURI += "/oauth2/v2.0/authorize?client_id=$ClientId"
        $authURI += "&response_type=code&redirect_uri= " + [System.Web.HttpUtility]::UrlEncode($RedirectURI)
        $authURI += "&response_mode=query&scope=" + [System.Web.HttpUtility]::UrlEncode($scopes) + "&state=$state"
        if($Prompt.IsPresent){
            $authURI += "&prompt=select_account"
        }
        # Create Window for User Sign-In
        $windowProperty = @{
            Width  = 500
            Height = 700
        }
        $signInWindow = New-Object System.Windows.Window -Property $windowProperty
        # Create WebBrowser for Window
        $browserProperty = @{
            Width  = 480
            Height = 680
        }
        $signInBrowser = New-Object System.Windows.Controls.WebBrowser -Property $browserProperty
        [void]$signInBrowser.navigate($authURI)
        
        # Create a condition to check after each page load
        $pageLoaded = {

            # Once a URL contains "code=*", close the Window
            if ($signInBrowser.Source -match "code=[^&]*") {

                # With the form closed and complete with the code, parse the query string

                $urlQueryString = [System.Uri]($signInBrowser.Source).Query
                $script:urlQueryValues = [System.Web.HttpUtility]::ParseQueryString($urlQueryString)

                [void]$signInWindow.Close()

            }
        }

        # Add condition to document completed
        [void]$signInBrowser.Add_LoadCompleted($pageLoaded)

        # Show Window
        [void]$signInWindow.AddChild($signInBrowser)
        [void]$signInWindow.ShowDialog()

        # Extract code from query string
        $authCode = $script:urlQueryValues.GetValues(($script:urlQueryValues.keys | Where-Object { $_ -eq "code" }))
        $Body =  @{"grant_type" = "authorization_code"; "scope" = $scopes; "client_id" = "$ClientId"; "code" =$authCode[0]; "redirect_uri" = $RedirectURI }
        $tokenRequest = Invoke-RestMethod -Method Post -ContentType application/x-www-form-urlencoded -Uri "https://login.microsoftonline.com/$tenantid/oauth2/token" -Body $Body 
        $AccessToken = $tokenRequest.access_token
        return $AccessToken
		
    }
    
}

function Get-FolderFromPath{
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
        $Prompt	
	)

    process{
        
        $EndPoint = "https://graph.microsoft.com/v1.0/users"
        $RequestURL = $EndPoint + "('$MailboxName')/MailFolders('MsgFolderRoot')/childfolders?"
        $AccessToken = Get-AccessTokenForGraph -MailboxName $Mailboxname -ClientId $ClientId -RedirectURI $RedirectURI -scopes $scopes -Prompt
        $fldArray = $FolderPath.Split("\")
        $PropList = @()
        $FolderSizeProp = Get-TaggedProperty -DataType "Long" -Id "0x66b3"
        $EntryId = Get-TaggedProperty -DataType "Binary" -Id "0xfff"
        $PropList += $FolderSizeProp 
        $PropList += $EntryId
        $Props = Get-ExtendedPropList -PropertyList $PropList 
        $RequestURL += "`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
        #Loop through the Split Array and do a Search for each level of folder 
        for ($lint = 1; $lint -lt $fldArray.Length; $lint++)
        {
            #Perform search based on the displayname of each folder level
            $FolderName = $fldArray[$lint];
            $headers = @{
                'Authorization' = "Bearer $AccessToken"
                'AnchorMailbox' = "$MailboxName"
            }
            $RequestURL = $RequestURL += "`&`$filter=DisplayName eq '$FolderName'"
            $tfTargetFolder = (Invoke-RestMethod -Method Get -Uri $RequestURL -UserAgent "GraphBasicsPs" -Headers $headers).value    

            if ($tfTargetFolder.displayname -match $FolderName)
            {
                $folderId = $tfTargetFolder.Id.ToString()
                $RequestURL = $EndPoint + "('$MailboxName')/MailFolders('$folderId')/childfolders?"
                $RequestURL += "`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
            }
            else
            {
                throw ("Folder Not found")
            }
        }
        if ($tfTargetFolder.singleValueExtendedProperties)
		{
			foreach ($Prop in $tfTargetFolder.singleValueExtendedProperties)
			{
				Switch ($Prop.Id)
				{
                    "Long 0x66b3" {      
                        $tfTargetFolder | Add-Member -NotePropertyName "FolderSize" -NotePropertyValue $Prop.value 
                    }
                    "Binary 0xfff" {
                        $tfTargetFolder | Add-Member -NotePropertyName "PR_ENTRYID" -NotePropertyValue ([System.BitConverter]::ToString([Convert]::FromBase64String($Prop.Value)).Replace("-",""))
                        $tfTargetFolder | Add-Member -NotePropertyName "ComplianceSearchId" -NotePropertyValue ("folderid:" + $tfTargetFolder.PR_ENTRYID.SubString(($tfTargetFolder.PR_ENTRYID.length-48)))
                    }
                }
            }
        }
        return $tfTargetFolder 
    }
}

function Get-TaggedProperty
{
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
	Begin
	{
		$Property = "" | Select-Object Id, DataType, PropertyType, Value
		$Property.Id = $Id
		$Property.DataType = $DataType
		$Property.PropertyType = "Tagged"
		if (![String]::IsNullOrEmpty($Value))
		{
			$Property.Value = $Value
		}
		return, $Property
	}
}


function Get-ExtendedPropList
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 1, Mandatory = $false)]
		[PSCustomObject]
		$PropertyList
	)
	Begin
	{
		$rtString = "";
		$PropName = "Id"
		foreach ($Prop in $PropertyList)
		{
			if ($Prop.PropertyType -eq "Tagged")
			{
				if ($rtString -eq "")
				{
					$rtString = "($PropName%20eq%20'" + $Prop.DataType + "%20" + $Prop.Id + "')"
				}
				else
				{
					$rtString += " or ($PropName%20eq%20'" + $Prop.DataType + "%20" + $Prop.Id + "')"
				}
			}
			else
			{
				if ($Prop.Type -eq "String")
				{
					if ($rtString -eq "")
					{
						$rtString = "($PropName%20eq%20'" + $Prop.DataType + "%20{" + $Prop.Guid + "}%20Name%20" + $Prop.Id + "')"
					}
					else
					{
						$rtString += " or ($PropName%20eq%20'" + $Prop.DataType + "%20{" + $Prop.Guid + "}%20Name%20" + $Prop.Id + "')"
					}
				}
				else
				{
					if ($rtString -eq "")
					{
						$rtString = "($PropName%20eq%20'" + $Prop.DataType + "%20{" + $Prop.Guid + "}%20Id%20" + $Prop.Id + "')"
					}
					else
					{
						$rtString += " or ($PropName%20eq%20'" + $Prop.DataType + "%20{" + $Prop.Guid + "}%20Id%20" + $Prop.Id + "')"
					}
				}
			}
			
		}
		return $rtString
		
	}
}


