function Export-EXCContactFolder {
    <#
		.SYNOPSIS
			Exports contacts from a mailbox/public folder to file.
		
		.DESCRIPTION
			Exports contacts from a mailbox/public folder to file.
	
			Currently, only export to CSV is being supported.
		
		.PARAMETER MailboxName
			The name of the mailbox to access.
		
		.PARAMETER Credentials
			Credentials that are authorized to access that mailbox.
		
		.PARAMETER Folder
			Default: Contacts
			The relative path to a folder with contacts within the mailbox.
			Example: Contacts\Private
		
		.PARAMETER PublicFolderPath
			The relative path to a public folder with contacts to export.
		
		.PARAMETER FileName
			Name of the file to export to.
			- File will be overwritten if it exists.
			- File needs not exist
			- Parent Folder must exist
			- User must have write access to the target path.
		
		.PARAMETER OutputType
			Default: CSV
			Currently, only CSV is supported as output type.
		
		.EXAMPLE
			PS C:\> Export-EXCContactFolder -MailboxName 'ben@example.com' -Credentials $Credentials -FileName 'C:\temp\contacts.csv'
	
			Exports the content of ben@example.com's default contacts folder to file.
	#>
    [CmdletBinding(DefaultParameterSetName = "Default")]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string]
        $MailboxName,
		
        [Parameter(Position = 1, Mandatory = $false)]
        [System.Management.Automation.PSCredential]
        $Credentials,
		
        [Parameter(Position = 2, Mandatory = $false, ParameterSetName = "Default")]
        [string]
        $Folder,
		
        [Parameter(Position = 2, Mandatory = $true, ParameterSetName = "PublicFolder")]
        [string]
        $PublicFolderPath,
		
        [Parameter(Position = 3, Mandatory = $true)]
        [string]
        $FileName,

        [Parameter(Position = 4, Mandatory = $false)]
        [switch]
        $Recurse,

        [Parameter(Position = 5, Mandatory = $false)]
        [switch]
        $RecurseMailbox,

        [Parameter(Position = 6, Mandatory = $False)]
		[switch]
		$ModernAuth,
		
		[Parameter(Position = 7, Mandatory = $False)]
		[String]
		$ClientId,

        [Parameter(Position = 8, Mandatory = $false)]
        [switch]
        $SkypeForBusinessContacts,
     
       
        [ValidateSet('CSV')]
        [string]
        $OutputType = "CSV"
    )
    begin
    {
        #region Utility functions
        function Get-Contacts {
            [CmdletBinding()]
            param (
                [Parameter(Position = 1, Mandatory = $true)]
                [Microsoft.Exchange.WebServices.Data.Folder]
                $ContactFolder
            )
            process {
                $FolderCollection = @()
                $FolderCollection += $ContactFolder
                $PR_Folder_Path = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
                if ($Recurse.IsPresent) {
                    Write-Host Getting Children
                    $FolderCollection = Get-ChildFolders -ContactFolder $ContactFolder -FolderCollection $FolderCollection 
                }
                if ($RecurseMailbox.IsPresent) {
                    $FolderCollection = Get-AllContactFolders -SMTPAddress $MailboxName -ContactFolder $ContactFolder 
                }
                foreach ($Folder in $FolderCollection) {
                    Write-Host ("Processing " + $Folder.DisplayName)
                    if ($Folder.TotalCount -gt 0) {
                        $psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
                        $PR_Gender = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(14925, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Short)
                        $psPropset.Add($PR_Gender)
                        #Define ItemView to retrive just 1000 Items      
                        $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
                        $fiItems = $null
                        do {
                            $fiItems = $service.FindItems($Folder.Id, $ivItemView)
                            [Void]$Folder.Service.LoadPropertiesForItems($fiItems, $psPropset)
                            foreach ($Item in $fiItems.Items) {
                                if ($Item -is [Microsoft.Exchange.WebServices.Data.Contact]) {
                                    if ($Recurse.IsPresent -bor $RecurseMailbox.IsPresent) {
                                        $expObj = "" | Select-Object FolderName, DisplayName, GivenName, Surname, Gender, Email1DisplayName, Email1Type, Email1EmailAddress, BusinessPhone, MobilePhone, HomePhone, BusinessStreet, BusinessCity, BusinessState, HomeStreet, HomeCity, HomeState
                                        $expObj.FolderName = $Folder.DisplayName
                                        $foldpathval = $null
                                        if ($Folder.TryGetProperty($PR_Folder_Path, [ref] $foldpathval)) {  

                                            $binarry = [Text.Encoding]::UTF8.GetBytes($foldpathval)  
                                            $hexArr = $binarry | ForEach-Object { $_.ToString("X2") }  
                                            $hexString = $hexArr -join ''  
                                            $hexString = $hexString.Replace("FEFF", "5C00")  
                                            $fpath = ConvertToString($hexString)
                                            $expObj.FolderName = $fpath
                                        }  
                                    }
                                    else {
                                        $expObj = "" | Select-Object DisplayName, GivenName, Surname, Gender, Email1DisplayName, Email1Type, Email1EmailAddress, BusinessPhone, MobilePhone, HomePhone, BusinessStreet, BusinessCity, BusinessState, HomeStreet, HomeCity, HomeState
                                    }
                                    $expObj.DisplayName = $Item.DisplayName
                                    $expObj.GivenName = $Item.GivenName
                                    $expObj.Surname = $Item.Surname
                                    $expObj.Gender = ""
                                    $Gender = $null
                                    if ($item.TryGetProperty($PR_Gender, [ref]$Gender)) {
                                        if ($Gender -eq 2) {
                                            $expObj.Gender = "Male"
                                        }
                                        if ($Gender -eq 1) {
                                            $expObj.Gender = "Female"
                                        }
                                    }
                                    $BusinessPhone = $null
                                    $MobilePhone = $null
                                    $HomePhone = $null
                                    if ($Item.PhoneNumbers -ne $null) {
                                        if ($Item.PhoneNumbers.TryGetValue([Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone, [ref]$BusinessPhone)) {
                                            $expObj.BusinessPhone = $BusinessPhone
                                        }
                                        if ($Item.PhoneNumbers.TryGetValue([Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone, [ref]$MobilePhone)) {
                                            $expObj.MobilePhone = $MobilePhone
                                        }
                                        if ($Item.PhoneNumbers.TryGetValue([Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::HomePhone, [ref]$HomePhone)) {
                                            $expObj.HomePhone = $HomePhone
                                        }
                                    }
                                    if ($Item.EmailAddresses.Contains([Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1)) {
                                        $expObj.Email1DisplayName = $Item.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Name
                                        $expObj.Email1Type = $Item.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].RoutingType
                                        $expObj.Email1EmailAddress = $Item.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address
                                    }
                                    $HomeAddress = $null
                                    $BusinessAddress = $null
                                    if ($item.PhysicalAddresses -ne $null) {
                                        if ($item.PhysicalAddresses.TryGetValue([Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Home, [ref]$HomeAddress)) {
                                            $expObj.HomeStreet = $HomeAddress.Street
                                            $expObj.HomeCity = $HomeAddress.City
                                            $expObj.HomeState = $HomeAddress.State
                                        }
                                        if ($item.PhysicalAddresses.TryGetValue([Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business, [ref]$BusinessAddress)) {
                                            $expObj.BusinessStreet = $BusinessAddress.Street
                                            $expObj.BusinessCity = $BusinessAddress.City
                                            $expObj.BusinessState = $BusinessAddress.State
                                        }
                                    }
									
                                    $expObj
                                }
                            }
                            $ivItemView.Offset += $fiItems.Items.Count
                        }
                        while ($fiItems.MoreAvailable -eq $true)
                    }
                }
            }
        }
        #endregion Utility functions
        function ConvertToString($ipInputString) {  
            $Val1Text = ""  
            for ($clInt = 0; $clInt -lt $ipInputString.length; $clInt++) {  
                $Val1Text = $Val1Text + [Convert]::ToString([Convert]::ToChar([Convert]::ToInt32($ipInputString.Substring($clInt, 2), 16)))  
                $clInt++  
            }  
            return $Val1Text  
        } 
        function Get-ChildFolders {
            [CmdletBinding()]
            param (
                [Parameter(Position = 1, Mandatory = $true)]
                [Microsoft.Exchange.WebServices.Data.Folder]
                $ContactFolder,
                [Parameter(Position = 2, Mandatory = $true)]
                [psObject]
                $FolderCollection
            )
            Process {
                $fvFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)  
                #Deep Transval will ensure all folders in the search path are returned  
                $fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow;  
                $psPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
                $PR_Folder_Path = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
                #Add Properties to the  Property Set  
                $psPropertySet.Add($PR_Folder_Path);  
                $fvFolderView.PropertySet = $psPropertySet;  
                $fiResult = $null  
                #The Do loop will handle any paging that is required if there are more the 1000 folders in a mailbox  
                do {  
                    $fiResult = $ContactFolder.FindFolders($fvFolderView)  
                    foreach ($ffFolder in $fiResult.Folders) {  
                        $foldpathval = $null  
                        #Try to get the FolderPath Value and then covert it to a usable String   
                        if ($ffFolder.TryGetProperty($PR_Folder_Path, [ref] $foldpathval)) {  
                            $binarry = [Text.Encoding]::UTF8.GetBytes($foldpathval)  
                            $hexArr = $binarry | ForEach-Object { $_.ToString("X2") }  
                            $hexString = $hexArr -join ''  
                            $hexString = $hexString.Replace("FEFF", "5C00")  
                            $fpath = ConvertToString($hexString)  
                        }  
                        write-host ("FolderPath : " + $fpath + " : " + $ffFolder.TotalCount)				
                        $FolderCollection += $ffFolder
                        if ($ffFolder.ChildFolderCount -gt 0) {
                            $FolderCollection = Get-ChildFolders -ContactFolder $ffFolder -FolderCollection $FolderCollection
                        }
                    } 
                    $fvFolderView.Offset += $fiResult.Folders.Count
                }while ($fiResult.MoreAvailable -eq $true)  
                return , $FolderCollection
            }
        }
        function Get-AllContactFolders {
            [CmdletBinding()]
            param (		
                [Parameter(Position = 1, Mandatory = $true)]
                [String]
                $SMTPAddress,	
                [Parameter(Position = 2, Mandatory = $true)]
                [Microsoft.Exchange.WebServices.Data.Folder]
                $ContactFolder
            )
            Process {
                $FolderCollection = @()
                $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $SMTPAddress)
                $RootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ContactFolder.Service, $folderid)
                $fvFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)  
                #Deep Transval will ensure all folders in the search path are returned  
                $fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;  
                $psPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
                $PR_Folder_Path = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
                $PR_ATTR_HIDDEN = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x10F4, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Boolean);
                #Add Properties to the  Property Set  
                $psPropertySet.Add($PR_Folder_Path);  
                $psPropertySet.Add($PR_ATTR_HIDDEN)
                $fvFolderView.PropertySet = $psPropertySet;  
                $fiResult = $null  
                #The Do loop will handle any paging that is required if there are more the 1000 folders in a mailbox  
                do {  
                    $fiResult = $RootFolder.FindFolders($fvFolderView)  
                    foreach ($ffFolder in $fiResult.Folders) {  
                        if ($ffFolder.FolderClass -contains "IPF.Contact") {
                            $BoolIsHidden = $false
                            [Void]$ffFolder.TryGetProperty($PR_ATTR_HIDDEN, [ref]$BoolIsHidden)
                            if (!$BoolIsHidden) {
                                $foldpathval = $null  
                                #Try to get the FolderPath Value and then covert it to a usable String   
                                if ($ffFolder.TryGetProperty($PR_Folder_Path, [ref] $foldpathval)) {  
                                    $binarry = [Text.Encoding]::UTF8.GetBytes($foldpathval)  
                                    $hexArr = $binarry | ForEach-Object { $_.ToString("X2") }  
                                    $hexString = $hexArr -join ''  
                                    $hexString = $hexString.Replace("FEFF", "5C00")  
                                    $fpath = ConvertToString($hexString)  
                                }  			
                                $FolderCollection += $ffFolder
                            }
                        }
                    } 
                    $fvFolderView.Offset += $fiResult.Folders.Count
                }while ($fiResult.MoreAvailable -eq $true)  
                return , $FolderCollection
            }
        }
        # Connect
        $service = Connect-EXCExchange -MailboxName $MailboxName -Credentials $Credentials -ModernAuth:$ModernAuth.IsPresent -ClientId $ClientId
    }
    process
    {
        if ($PublicFolderPath) {
            $service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName)
            $fldId = Get-PublicFolderIdFromPath -FolderPath $PublicFolderPath -SmtpAddress $MailboxName -service $service
            $contactsId = New-Object Microsoft.Exchange.WebServices.Data.FolderId($fldId)
            $contactFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $contactsId)
            Set-PublicFolderRoutingHeader -Service $service -Credentials $Credentials -MailboxName $MailboxName -Header "X-AnchorMailbox"
        }
        else {
            if ($Folder) {
                $contactFolder = Get-EXCContactFolder -Service $service -FolderPath $Folder -SmptAddress $MailboxName
            }
            else {
                if ($SkypeForBusinessContacts.IsPresent) {
                    $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts, $MailboxName)
                }
                else {
                    $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts, $MailboxName)
                }				
                $contactFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderid)
            }
        }
		
        switch ($OutputType) {
            "CSV" { Get-Contacts -ContactFolder $contactFolder | Export-Csv -NoTypeInformation -Path $FileName -encoding "UTF8"}
            default { throw "Invalid output type: $OutputType" }
        }
    }
}

