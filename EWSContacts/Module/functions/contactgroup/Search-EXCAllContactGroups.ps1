function Search-EXCAllContactGroups
{

	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $true)]
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
		if(![String]::IsNullOrEmpty($EmailAddress)){			
			$SearchType = "EmailAddress"
		}
		if(![String]::IsNullOrEmpty($DisplayName)){			
			$SearchType = "DisplayName"
		}
		$Folders = Get-EXCAllContactFolders -MailboxName $MailboxName -Credential $Credentials -useImpersonation:$useImpersonation.IsPresent
		$cnpsPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
		$cnpsPropset.Add([Microsoft.Exchange.WebServices.Data.ContactGroupSchema]::Members)
		foreach($Folder in $Folders){
			write-host ("Searching Folder : " + $Folder.FolderPath)
			$SfSearchFilter2 = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass, "IPM.DistList")

			#Define ItemView to retrive just 1000 Items    
			$ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
			#$ivItemView.PropertySet = $cnpsPropset
			$fiItems = $null
			do
			{
				
				$fiItems = $service.FindItems($Folder.Id, $SfSearchFilter2, $ivItemView)
				if ($fiItems.Item.Count -gt 0)
				{
					[Void]$service.LoadPropertiesForItems($fiItems,$cnpsPropset)  
				}				
				foreach ($Item in $fiItems.Items)
				{
					$output = $false
					foreach($Member in $Item.Members){
						switch($SearchType){
							
							"EmailAddress" {if(![String]::IsNullOrEmpty($member.AddressInformation.Address)){ 
								if($member.AddressInformation.Address.ToLower() -eq $EmailAddress.ToLower()){
										$Item | Add-Member -Name "MatchedMember" -Value $Member -MemberType NoteProperty -Force
										$output = $true
								  }
								}
							}
							"DisplayName" { 								
								if(![String]::IsNullOrEmpty($member.AddressInformation.Name)){ 
								if($member.AddressInformation.Name.ToLower() -eq $DisplayName.ToLower()){
										$Item | Add-Member -Name "MatchedMember" -Value $Member -MemberType NoteProperty -Force
										$output = $true
								  }
								}
							}
						}
					}
					if($output){
						$Item | Add-Member -Name "FolderPath" -Value $Folder.FolderPath -MemberType NoteProperty -Force
						Write-Output $Item
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