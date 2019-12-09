function Search-EXCAllContactFolders
{

	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[System.Management.Automation.PSCredential]
		$Credentials,		
		
		[Parameter(Position = 2, Mandatory = $false)]
		[switch]
		$useImpersonation,

		[Parameter(Position = 3, Mandatory = $false)]
		[String]
		$EmailAddress,

		[Parameter(Position = 4, Mandatory = $false)]
		[String]
		$DisplayName,

		[Parameter(Position = 5, Mandatory = $False)]
		[switch]
		$ModernAuth,
		
		[Parameter(Position = 6, Mandatory = $False)]
		[String]
		$ClientId
		
	

	)
	Begin
	{
		#Connect
		$service = Connect-EXCExchange -MailboxName $MailboxName -Credential $Credentials -ModernAuth:$ModernAuth.IsPresent -ClientId $ClientId
		if ($useImpersonation.IsPresent)
		{
			$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}
		$SearchType = ""
		$Folders = Get-EXCAllContactFolders -MailboxName $MailboxName -Credential $Credentials -useImpersonation:$useImpersonation.IsPresent
		if(![String]::IsNullOrEmpty($EmailAddress)){
			$KQL =  $EmailAddress
			$SearchType = "EmailAddress"
		}
		if(![String]::IsNullOrEmpty($DisplayName)){
			$KQL = $DisplayName
			$SearchType = "DisplayName"
		}
		foreach($Folder in $Folders){
			write-host ("Searching Folder : " + $Folder.FolderPath)
			$ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
			$fiItems = $null
			do
			{
				$fiItems = $service.FindItems($Folder.Id, $KQL, $ivItemView)
				if ($fiItems.Items.Count -gt 0)
				{
					$psPropset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
					$DisplayNameFirstLast = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Address,"DisplayNameFirstLast", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);
					$DisplayNameLastFirst = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Address,"DisplayNameLastFirst", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);
					$psPropset.Add($DisplayNameFirstLast)
					$psPropset.Add($DisplayNameLastFirst)
					[Void]$service.LoadPropertiesForItems($fiItems, $psPropset)
					foreach ($Contact in $fiItems.Items)
					{
						$output = $fase;
						$EmailAddress1 = $null
        				$EmailAddress2 = $null
       				    $EmailAddress3 = $null
						switch($SearchType){
							"EmailAddress" {
								if ($Contact.EmailAddresses -ne $null){
									if ($Contact.EmailAddresses.TryGetValue([Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1,[ref]$EmailAddress1)) {
										if($EmailAddress1.Address.ToLower() -eq $EmailAddress){
											$Contact | Add-Member -Name "Matched" -Value EmailAddress1 -MemberType NoteProperty -Force
											$output = $true;
										}
									}
									if ($Contact.EmailAddresses.TryGetValue([Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress2,[ref]$EmailAddress2)) {
										if($EmailAddress2.Address.ToLower() -eq $EmailAddress){
											$Contact | Add-Member -Name "Matched" -Value EmailAddress2 -MemberType NoteProperty -Force
											$output = $true;
										}
									}
									if ($Contact.EmailAddresses.TryGetValue([Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress3,[ref]$EmailAddress3)) {
										if($EmailAddress3.Address.ToLower() -eq $EmailAddress){
											$Contact | Add-Member -Name "Matched" -Value EmailAddress3 -MemberType NoteProperty -Force
											$output = $true;
										}
									}
								}
							}
							"DisplayName" {
								if([String]::IsNullOrEmpty(!$Contact.DisplayName)){
									$Sanitize = $Contact.DisplayName.ToLower().Replace(",")
									if($Sanitize -eq $DisplayName.ToLower()){
										$Contact | Add-Member -Name "Matched" -Value DisplayName -MemberType NoteProperty -Force
										$output = $true;
									}
								}
								$EmailAddress1 = $null
								if ($Contact.EmailAddresses -ne $null){
									if ($Contact.EmailAddresses.TryGetValue([Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1,[ref]$EmailAddress1)) {
										$Sanitize = $EmailAddress1.Name.ToLower()
										if($Sanitize -eq $DisplayName.ToLower()){
											$Contact | Add-Member -Name "Matched" -Value EmailAddress1 -MemberType NoteProperty -Force
											$output = $true;
										}
									}
								}
								$DisplayNameFirstLastValue = $null
								if($Contact.TryGetProperty($DisplayNameFirstLast,[ref]$DisplayNameFirstLastValue)){
										if(![String]::IsNullOrEmpty($DisplayNameFirstLastValue)){
											$Sanitize = $DisplayNameFirstLastValue.ToLower()
											if($Sanitize -eq $DisplayName.ToLower()){
												$Contact | Add-Member -Name "Matched" -Value DisplayNameFirstLastValue -MemberType NoteProperty -Force
												$output = $true;
											}
										}

								}
								$DisplayNameLastFirstValue = $null
								if($Contact.TryGetProperty($DisplayNameLastFirst,[ref]$DisplayNameLastFirstValue)){
									if(![String]::IsNullOrEmpty($DisplayNameLastFirstValue)){
										$Sanitize = $DisplayNameLastFirstValue.ToLower()
										if($Sanitize -eq $DisplayName.ToLower()){
											$Contact | Add-Member -Name "Matched" -Value DisplayNameLastFirstValue -MemberType NoteProperty -Force
											$output = $true;
										}
									}

								}
							}

						}
						if($output){
							$Contact | Add-Member -Name "FolderPath" -Value $Folder.FolderPath -MemberType NoteProperty -Force
							Write-Output $Contact
						}
					}
				}
				$ivItemView.Offset += $fiItems.Items.Count
			}
			while ($fiItems.MoreAvailable -eq $true)
		}
	}
}
function ConvertToString($ipInputString) {  
    $Val1Text = ""  
    for ($clInt = 0; $clInt -lt $ipInputString.length; $clInt++) {  
        $Val1Text = $Val1Text + [Convert]::ToString([Convert]::ToChar([Convert]::ToInt32($ipInputString.Substring($clInt, 2), 16)))  
        $clInt++  
    }  
    return $Val1Text  
} 