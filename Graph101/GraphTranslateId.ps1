function Invoke-TranslateId {
  [CmdletBinding()] 
  param( 
    [Parameter(Position = 1, Mandatory = $false)]
    [String]
    $Base64FolderIdToTranslate = "",
    [Parameter(Position = 1, Mandatory = $false)]
    [String]
    $HexEntryId = ""

  )  
  Process {
    if(![String]::IsNullOrEmpty($Base64FolderIdToTranslate)){
      $HexEntryId = [System.BitConverter]::ToString([Convert]::FromBase64String($Base64FolderIdToTranslate)).Replace("-", "").Substring(2)  
      $HexEntryId = $HexEntryId.SubString(0, ($HexEntryId.Length - 2))
    }
    $FolderIdBytes = [byte[]]::new($HexEntryId.Length / 2)
    For ($i = 0; $i -lt $HexEntryId.Length; $i += 2) {      $FolderIdBytes[$i / 2] = [convert]::ToByte($HexEntryId.Substring($i, 2), 16)
    }

    $FolderIdToConvert = [System.Web.HttpServerUtility]::UrlTokenEncode($FolderIdBytes)

    $ConvertRequest = @"
{
    "inputIds" : [
      "$FolderIdToConvert"
    ],
    "sourceIdType": "entryId",
    "targetIdType": "restId"
  }
"@

    $ConvertResult = Invoke-MgGraphRequest -Method POST -Uri https://graph.microsoft.com/v1.0/me/translateExchangeIds -Body $ConvertRequest


    return $ConvertResult.Value.targetId

  }

}
