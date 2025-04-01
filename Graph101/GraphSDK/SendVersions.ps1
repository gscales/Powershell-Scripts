function Invoke-GetUserMailSendVersions{
    [CmdletBinding()] 
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [String] $MailboxName,
        [Parameter(Position = 1, Mandatory = $false)]
        [String] $FolderId="SentItems",
        [Parameter(Position = 1, Mandatory = $false)]
        [DateTime] $StartTime
    )
Process{
        $filter = ""
        if($StartTime){
           $filter =  "receivedDateTime gt " + $StartTime.ToString("yyyy-MM-dd") + "T00:00:00Z"
        }
        $MailClients = @{}
        Get-MgUserMailFolderMessage -MailFolderId $FolderId -UserId $MailboxName -All -Select "Subject,receivedDateTime,singleValueExtendedProperties" -ExpandProperty "singleValueExtendedProperties(`$filter=id eq 'String {41F28F13-83F4-4114-A584-EEDB5A6B0BFF} name ClientInfo')" -Filter $filter | ForEach-Object{
            $Item = $_
            Expand-ExtendedProperties -Item $Item
            if($Item.ClientInfo){
                if($MailClients.ContainsKey($Item.ClientInfo)){
                    $MailClients[$Item.ClientInfo].NumberOfItems++
                    if($MailClients[$Item.ClientInfo].LastEmail -lt $Item.receivedDateTime){
                        $MailClients[$Item.ClientInfo].LastEmail = $Item.receivedDateTime
                        $MailClients[$Item.ClientInfo].LastEmailSubject = $Item.Subject
                    }
                }else{
                    $rptobj = "" | Select ClientInfo,NumberOfItems,AppId,ServicePrincipalDisplayName,LastEmail,LastEmailSubject
                    $rptobj.ClientInfo = $Item.ClientInfo
                    $rptobj.NumberOfItems = 1
                    $rptobj.LastEmailSubject = $Item.Subject
                    $rptobj.LastEmail = $Item.receivedDateTime 
                    $rptobj.AppId = $Item.AppId
                    $rptobj.ServicePrincipalDisplayName = $Item.ServicePrincipalDisplayName
                    $MailClients.Add($Item.ClientInfo,$rptobj)
                }
            }            
        }
        return $MailClients.Values
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
                    "String {41F28F13-83F4-4114-A584-EEDB5A6B0BFF} name ClientInfo"{
                        Add-Member -InputObject $Item -NotePropertyName "ClientInfo" -NotePropertyValue $Prop.Value -Force
                        $appId = [Regex]::Match($Item.ClientInfo ,"\[AppId=([^\]]+)\]")                        
                        if($appId.Success){
                            Add-Member -InputObject $Item -NotePropertyName "AppId" -NotePropertyValue $appId.captures[0].Value -Force
                            if($Item.AppId){
                                $guid = $Item.AppId.Replace("[AppId=","").Replace("]","")
                                Write-Verbose ("Processing Guid " + $guid)
                                if(!$Script:SpCache.ContainsKey($guid)){
                                    $sp = Get-MgServicePrincipal -Filter "AppId eq '$guid'"
                                    $Script:SpCache.Add($guid,$sp)
                                }else{
                                    $sp = $Script:SpCache[$guid]
                                }                                
                                Add-Member -InputObject $Item -NotePropertyName "ServicePrincipalDisplayName" -NotePropertyValue $sp.DisplayName -Force
                            }
                        }
                    }
                }
            }
        }
    }
}

$Script:SpCache = @{}