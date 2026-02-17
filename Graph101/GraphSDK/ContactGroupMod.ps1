function ConvertTo-MapiGuid {
    param([byte[]]$bytes)
    
    if ($bytes.Length -ne 16) {
        throw "GUID must be exactly 16 bytes"
    }
    
    # MAPI GUID format: first 3 fields are little-endian, last 2 are big-endian
    # Reverse bytes for Data1 (first 4 bytes)
    $data1 = [BitConverter]::ToString($bytes[3..0]).Replace("-", "")
    
    # Reverse bytes for Data2 (next 2 bytes)
    $data2 = [BitConverter]::ToString($bytes[5..4]).Replace("-", "")
    
    # Reverse bytes for Data3 (next 2 bytes)  
    $data3 = [BitConverter]::ToString($bytes[7..6]).Replace("-", "")
    
    # Data4 stays in order (last 8 bytes)
    $data4 = [BitConverter]::ToString($bytes[8..15]).Replace("-", "")
    
    # Format as GUID
    return "$data1-$data2-$data3-$($data4.Substring(0,4))-$($data4.Substring(4))".ToUpper()
}

function ConvertTo-RawGuidString {
    param([byte[]]$bytes)
    
    if ($bytes.Length -ne 16) {
        throw "GUID must be exactly 16 bytes"
    }
    
    # Just convert bytes directly to hex string in GUID format (no endian conversion)
    $data1 = [BitConverter]::ToString($bytes[0..3]).Replace("-", "")
    $data2 = [BitConverter]::ToString($bytes[4..5]).Replace("-", "")
    $data3 = [BitConverter]::ToString($bytes[6..7]).Replace("-", "")
    $data4 = [BitConverter]::ToString($bytes[8..15]).Replace("-", "")
    
    return "$data1-$data2-$data3-$($data4.Substring(0,4))-$($data4.Substring(4))".ToUpper()
}

function Invoke-ParseOneOffEntryID {
    param (
        [Parameter(Mandatory = $true)]
        [string]$entryID
    )

    # 1. Convert hex to byte array
    $hex = $entryID -replace '\s',''
    $bytes = for ($i = 0; $i -lt $hex.Length; $i += 2) {
        [Convert]::ToByte($hex.Substring($i, 2), 16)
    }

    # 2. Extract Bitmask (Bytes 20-23)
    # We look at Byte 23 directly (the most significant byte in Little Endian)
    # 0x80000000 means the high bit of Byte 23 is set.
    $isUnicode = ($bytes[23] -band 0x80) -eq 0x80
    $encoding = if ($isUnicode) { [System.Text.Encoding]::Unicode } else { [System.Text.Encoding]::Default }
    $nullSize = if ($isUnicode) { 2 } else { 1 }

    # 3. String extraction logic
    $cursor = 24
    $results = @()

    for ($s = 0; $s -lt 3; $s++) {
        $start = $cursor
        while ($cursor -lt $bytes.Length) {
            if ($isUnicode) {
                # Look for double-null (00 00)
                if ($bytes[$cursor] -eq 0 -and $bytes[$cursor+1] -eq 0) { break }
                $cursor += 2
            } else {
                # Look for single-null (00)
                if ($bytes[$cursor] -eq 0) { break }
                $cursor += 1
            }
        }
        
        # Extract string and move cursor past null
        $len = $cursor - $start
        if ($len -gt 0) {
            $results += $encoding.GetString($bytes, $start, $len)
        } else {
            $results += ""
        }
        $cursor += $nullSize
    }

    # 4. Correct for GUID Endianness (MFCMAPI style)
    $g = $bytes[4..19]
    $guidStr = "{0:X2}{1:X2}{2:X2}{3:X2}-{4:X2}{5:X2}-{6:X2}{7:X2}-{8:X2}{9:X2}-{10:X2}{11:X2}{12:X2}{13:X2}{14:X2}{15:X2}" -f `
        $g[3],$g[2],$g[1],$g[0], $g[5],$g[4], $g[7],$g[6], $g[8],$g[9], $g[10],$g[11],$g[12],$g[13],$g[14],$g[15]

    return [PSCustomObject]@{
        DisplayName  = $results[0]
        AddressType  = $results[1]
        EmailAddress = $results[2]
        IsUnicode    = $isUnicode
        Bitmask      = "0x{0:X2}{1:X2}{2:X2}{3:X2}" -f $bytes[23], $bytes[22], $bytes[21], $bytes[20]
        ProviderGUID = $guidStr
    }
}

function Invoke-ParseWrappedEntryID {
    param([string]$Hex)

    $cleanHex = $Hex -replace '[^0-9A-Fa-f]', ''
    $b = [byte[]]::new($cleanHex.Length / 2)
    for($i=0; $i -lt $cleanHex.Length; $i+=2) { $b[$i/2] = [Convert]::ToByte($cleanHex.Substring($i, 2), 16) }

    # DisplayType Enums (MS-OXOABK)
    $DisplayTypes = @{
        0x00 = "DT_MAILUSER (Local User)"
        0x01 = "DT_DISTLIST (Distribution List)"
        0x02 = "DT_FORUM (Bulletin Board)"
        0x03 = "DT_AGENT (Automated Mailbox)"
        0x04 = "DT_ORGANIZATION"
        0x05 = "DT_PRIVATE_DISTLIST (Personal DL)"
        0x06 = "DT_REMOTE_MAILUSER (Exchange User)"
    }

    $results = [Ordered]@{
        OuterFlags   = "0x" + [BitConverter]::ToString($b, 0, 4).Replace("-", "")
        ProviderGUID = ConvertTo-MapiGuid -bytes ([byte[]]$b[4..19])
    }

    $cursor = 20
    if ($results.ProviderGUID -eq "D3AD91C0-9D51-11CF-A4A9-00AA0047FAA4") {
        $wrappedByte = $b[$cursor]
        $results.WrappedTypeHex = "0x{0:X2}" -f $wrappedByte
        
        # Bitmask Interpretation (MS-OXOCNTC)
        $innerTypeCode = $wrappedByte -band 0x0F
        $results.InnerType = switch($innerTypeCode) {
            3 { "Contact Object" }
            4 { "Personal Distribution List (Outlook Internal)" }
            5 { "GAL Mail User" }
            6 { "GAL Distribution List" }
            default { "Unknown ($innerTypeCode)" }
        }
        $cursor++

        # Process Inner Data
        $results.InnerFlags = "0x" + [BitConverter]::ToString($b, $cursor, 4).Replace("-", "")
        $cursor += 4
        $innerGuid = ConvertTo-MapiGuid -bytes ([byte[]]$b[$cursor..($cursor + 15)])
        $results.InnerProviderGUID = $innerGuid
        $cursor += 16

        # CASE: Exchange/GAL (0xB5 or 0xB6)
        if ($innerGuid -eq "DC07C840-C042-101A-B4B9-08002B2FE182") {
            if ($b.Length -ge ($cursor + 8)) {
                $results.Version = [BitConverter]::ToUInt32($b, $cursor)
                $cursor += 4
                $dType = [BitConverter]::ToUInt32($b, $cursor)
                if ($DisplayTypes.ContainsKey($dType)) {
                    $results.DisplayType = $DisplayTypes[$dType]
                } else {
                    $results.DisplayType = "0x{0:X8}" -f $dType
                }
                $cursor += 4
                $nullPos = [Array]::IndexOf($b, [byte]0, $cursor)
                $len = if ($nullPos -eq -1) { $b.Length - $cursor } else { $nullPos - $cursor }
                if ($len -gt 0) { $results.X500DN = [System.Text.Encoding]::ASCII.GetString($b, $cursor, $len) }
            }
        }
        # CASE: Personal Distribution List / Contact (0xB4)
        else {
            $results.Note = "This is a Message EntryID (Stored in a PST or Mailbox Store)."
            # Message EntryIDs often contain the Store UID and the Folder/Message ID
            $results.StoreUID = $innerGuid
            $results.RemainingData = [BitConverter]::ToString($b, $cursor).Replace("-", "")
        }
    }
    return $results
}

function Invoke-ProcessEntryID {
    param([string]$Hex)
    
    $cleanHex = $Hex -replace '[^0-9A-Fa-f]', ''
    if ($cleanHex.Length -lt 40) { return "Unknown" }
    
    $b = [byte[]]::new($cleanHex.Length / 2)
    for($i=0; $i -lt $cleanHex.Length; $i+=2) { 
        $b[$i/2] = [Convert]::ToByte($cleanHex.Substring($i, 2), 16) 
    }
    
    # Extract Provider GUID (bytes 4-19)
    $guidBytes = [byte[]]$b[4..19]
    
    # Try as raw bytes first (for One-Off)
    $providerGuidRaw = ConvertTo-RawGuidString -bytes $guidBytes
    
    # Try as MAPI format (for Wrapped)
    $providerGuidMapi = ConvertTo-MapiGuid -bytes $guidBytes
    
    # One-Off EntryID GUID (raw byte order)
    if ($providerGuidMapi -eq "A41F2B81-A3BE-1910-9D6E-00DD010F5402") {
        Invoke-ParseOneOffEntryID -entryID $Hex
    }
    # Wrapped EntryID GUID (MAPI format)
    elseif ($providerGuidMapi -eq "D3AD91C0-9D51-11CF-A4A9-00AA0047FAA4") {
        Invoke-ParseWrappedEntryID -Hex $Hex
    }
    else {
        return "Other - Raw: $providerGuidRaw | MAPI: $providerGuidMapi"
    }
}

function Invoke-EnumerateContactGroups{
    param(
        [Parameter(Mandatory = $true)]
        [string]$MailboxName
    )
    $MailboxId = (Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/beta/users/$MailboxName/settings/exchange").primaryMailboxId
    # Full path to the mailbox Contacts items
    $baseUri = "https://graph.microsoft.com/beta/admin/exchange/mailboxes/$MailboxId/folders/Contacts/items"

    # Filter to only Contact Groups
    $filter = "singleValueExtendedProperties/any(p:p/id eq 'String 0x001A' and p/value eq 'IPM.DistList')"

    # Expand all extended properties
    $expand ="singleValueExtendedProperties(`$filter=" +
            "id eq 'String 0x001A' or id eq 'String 0x0037'" +
            ")"

    # Properly build the URI
    $uri = "$baseUri`?`$top=999&`$filter=$filter&`$expand=$expand"

    $response = Invoke-MgGraphRequest -Method Get -Uri $uri

    $groups = @()
    foreach($item in $response.value) {
        $group = @{
            Id = $item.id           
        }
        $idVal = $item.id
        # Extract the Contact Group Name (0x0037)
        $groupNameProp = $item.singleValueExtendedProperties | Where-Object { $_.id -eq 'String 0x37' }
        if ($groupNameProp) {
            $group.DisplayName = $groupNameProp.value
        }
        $expand = "singleValueExtendedProperties(`$filter=" +
              "id eq 'String 0x0037' or id eq 'Binary {00062004-0000-0000-C000-000000000046} Id 0x8064'" +
              "),multiValueExtendedProperties(`$filter=id eq 'BinaryArray {00062004-0000-0000-C000-000000000046} Id 0x8055' or id eq 'BinaryArray {00062004-0000-0000-C000-000000000046} Id 0x8054')"
        $uri = "$baseUri\$idval`?`$expand=$expand"
        $ContactGroup = Invoke-MgGraphRequest -Method Get -Uri $uri 
        Expand-ExtendedProperties -Item $ContactGroup       
        Write-Output $ContactGroup
    }    
}

function Invoke-FindContactGroup{
    param(
        [Parameter(Mandatory = $true)]
        [string]$MailboxName,
        [Parameter(Mandatory = $true)]
        [string]$GroupName
    )
    $MailboxId = (Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/beta/users/$MailboxName/settings/exchange").primaryMailboxId
    # Full path to the mailbox Contacts items
    $baseUri = "https://graph.microsoft.com/beta/admin/exchange/mailboxes/$MailboxId/folders/Contacts/items"

    # Filter to only Contact Groups
    $filter = "singleValueExtendedProperties/any(p:p/id eq 'String 0x001A' and p/value eq 'IPM.DistList') and singleValueExtendedProperties/any(p:p/id eq 'String 0x0037' and p/value eq '$GroupName')"

    # Expand all extended properties
    $expand ="singleValueExtendedProperties(`$filter=" +
            "id eq 'String 0x001A' or id eq 'String 0x0037'" +
            ")"

    # Properly build the URI
    $uri = "$baseUri`?`$top=999&`$filter=$filter&`$expand=$expand"

    $response = Invoke-MgGraphRequest -Method Get -Uri $uri

    $groups = @()
    foreach($item in $response.value) {
        $group = @{
            Id = $item.id           
        }
        $idVal = $item.id
        # Extract the Contact Group Name (0x0037)
        $groupNameProp = $item.singleValueExtendedProperties | Where-Object { $_.id -eq 'String 0x37' }
        if ($groupNameProp) {
            $group.DisplayName = $groupNameProp.value
        }
        $expand = "singleValueExtendedProperties(`$filter=" +
              "id eq 'String 0x0037' or id eq 'Binary {00062004-0000-0000-C000-000000000046} Id 0x8064'" +
              "),multiValueExtendedProperties(`$filter=id eq 'BinaryArray {00062004-0000-0000-C000-000000000046} Id 0x8055' or id eq 'BinaryArray {00062004-0000-0000-C000-000000000046} Id 0x8054')"
        $uri = "$baseUri\$idval`?`$expand=$expand"
        $ContactGroup = Invoke-MgGraphRequest -Method Get -Uri $uri 
        Expand-ExtendedProperties -Item $ContactGroup       
        Write-Output $ContactGroup
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
                    "String 0x37"{
                         $Item.Add("DisplayName", $Prop.Value) | Out-Null
                    } 
                    "Binary {00062004-0000-0000-C000-000000000046} Id 0x8064"{
                        $Item.Add("MemberStream", $Prop.Value) | Out-Null
                    }                
                }
            }
        }
        if ($Item.multiValueExtendedProperties -ne $null)
        {
            foreach ($Prop in $Item.multiValueExtendedProperties)
            {
                Switch ($Prop.Id)
                {
                    "BinaryArray {00062004-0000-0000-C000-000000000046} Id 0x8055"{
                        $Item.add("Members", $Prop.Value) | Out-Null      
                    }
                    "BinaryArray {00062004-0000-0000-C000-000000000046} Id 0x8054"{
                        $Item.Add("MemberOneOffEntryIDs", $Prop.Value) | Out-Null
                    }
                }
            }
        }
    }
}
function Expand-GroupMembership {
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [psobject]$ContactGroup
    )
    process {
        if($ContactGroup.MemberStream){
            $streamBytes = [System.Convert]::FromBase64String($ContactGroup.MemberStream)
            $memberInfo = Read-DistListMemberInfoArray -Bytes $streamBytes
            foreach($member in $memberInfo.Members){
                $entryIdHex = [BitConverter]::ToString($member.OneOffEntryIdBytes).Replace("-","")
                $parsed = Invoke-ParseOneOffEntryID -entryID $entryIdHex
                Write-Output $parsed
            }
        }else{
            foreach($oneOffMember in $ContactGroup.MemberOneOffEntryIDs){
                $oneOffEntryID = [System.BitConverter]::ToString([System.Convert]::FromBase64String($oneOffMember)).Replace("-","")
                $parsed = Invoke-ParseOneOffEntryID -entryID $oneOffEntryID
                Write-Output $parsed
            }
        }
    }
}

function Read-DistributionListStreamHeader {
    param($Ctx)

    [pscustomobject]@{
        StreamVersion        = Read-UInt16LE $Ctx
        Reserved             = Read-UInt16LE $Ctx
        BuildVersion         = Read-UInt32LE $Ctx
        DistListStreamFlags  = Read-UInt32LE $Ctx
        CountOfEntries       = Read-UInt32LE $Ctx
        TotalEntryIdSize     = Read-UInt32LE $Ctx
        TotalOneOffSize      = Read-UInt32LE $Ctx
        TotalExtraSize       = Read-UInt32LE $Ctx
    }
}

function Read-DistListMemberInfoArray {
    param(
        [byte[]]$Bytes
    )

    $ctx = New-BinaryContext $Bytes

    $header = Read-DistributionListStreamHeader $ctx
    $members = @()

    for ($i = 0; $i -lt $header.CountOfEntries; $i++) {

        $memberStart = $ctx.Offset

        # EntryId
        $entryIdSize = Read-UInt32LE $ctx
        $entryIdData = Read-Bytes   $ctx $entryIdSize

        # Optional OneOffEntryId
        $oneOffSize = Read-UInt32LE $ctx
        $oneOffData = $null
        if ($oneOffSize -gt 0) {
            $oneOffData = Read-Bytes $ctx $oneOffSize
        }

        # ExtraMemberInfo (MUST be zero)
        $extraSize = Read-UInt32LE $ctx
        if ($extraSize -ne 0) {
            throw "ExtraMemberInfoSize != 0 at member index $i (offset $memberStart)"
        }

        $members += [pscustomobject]@{
            Index              = $i
            Offset             = $memberStart
            EntryIdSize        = $entryIdSize
            EntryIdRawBytes    = $entryIdData
            OneOffEntryIdSize  = $oneOffSize
            OneOffEntryIdBytes = $oneOffData
        }
    }

    # Terminators (DWORD 0, DWORD 0)
    $terminator1 = Read-UInt32LE $ctx
    $terminator2 = Read-UInt32LE $ctx

    if ($terminator1 -ne 0 -or $terminator2 -ne 0) {
        throw "Invalid stream terminator"
    }

    return [pscustomobject]@{
        Header  = $header
        Members = $members
    }
}

function New-BinaryContext {
    param([byte[]]$Bytes)

    [pscustomobject]@{
        Bytes  = $Bytes
        Offset = 0
        Length = $Bytes.Length
    }
}

function Read-UInt32LE {
    param($Ctx)

    if ($Ctx.Offset + 4 -gt $Ctx.Length) {
        throw "UInt32 read overflow at offset $($Ctx.Offset)"
    }

    $v = [BitConverter]::ToUInt32($Ctx.Bytes, $Ctx.Offset)
    $Ctx.Offset += 4
    return $v
}

function Read-UInt16LE {
    param($Ctx)

    if ($Ctx.Offset + 2 -gt $Ctx.Length) {
        throw "UInt16 read overflow at offset $($Ctx.Offset)"
    }

    $v = [BitConverter]::ToUInt16($Ctx.Bytes, $Ctx.Offset)
    $Ctx.Offset += 2
    return $v
}

function Read-Bytes {
    param($Ctx, [int]$Count)

    if ($Ctx.Offset + $Count -gt $Ctx.Length) {
        throw "Byte read overflow ($Count bytes) at offset $($Ctx.Offset)"
    }

    $data = $Ctx.Bytes[$Ctx.Offset..($Ctx.Offset + $Count - 1)]
    $Ctx.Offset += $Count
    return ,$data
}
