function Invoke-GetReactionDetailsForMessages {
    [CmdletBinding()] 
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [String] $MailboxName,
        [Parameter(Position = 1, Mandatory = $false)]
        [String] $FolderId="Inbox",
        [Parameter(Position = 2, Mandatory = $false)]
        [DateTime] $StartTime,
        [Parameter(Position = 3, Mandatory = $false)]
        [switch] $OwnerOnly

    )
Process{
        $ExpandProperty = "singleValueExtendedProperties(`$filter=(id eq 'SystemTime {41F28F13-83F4-4114-A584-EEDB5A6B0BFF} name OwnerReactionTime') or (id eq 'String {41F28F13-83F4-4114-A584-EEDB5A6B0BFF} name OwnerReactionType')" 
        $ExpandProperty += "or (id eq 'Binary {41F28F13-83F4-4114-A584-EEDB5A6B0BFF} name ReactionsSummary') or (id eq 'Integer {41F28F13-83F4-4114-A584-EEDB5A6B0BFF} name ReactionsCount') or (id eq 'Binary {41F28F13-83F4-4114-A584-EEDB5A6B0BFF} name ReactionsHistory'))"
        if($OwnerOnly){
            $filter = "singleValueExtendedProperties/any(ep: ep/id eq 'String {41F28F13-83F4-4114-A584-EEDB5A6B0BFF} name OwnerReactionType' and ep/value ne null)"
        }else{
            $filter = "singleValueExtendedProperties/any(ep: ep/id eq 'Integer {41F28F13-83F4-4114-A584-EEDB5A6B0BFF} name ReactionsCount' and cast(ep/value, Edm.Int32) gt 0)"
        }       
        if($StartTime){
            $filter =  "($filter) And (receivedDateTime gt " + $StartTime.ToString("yyyy-MM-dd") + "T00:00:00Z)"
         }
        $Messages = Get-MgUserMailFolderMessage -MailFolderId $FolderId -UserId $MailboxName -All -PageSize 999 -Select "Subject,receivedDateTime,singleValueExtendedProperties,InternetMessageId"  -ExpandProperty $ExpandProperty -Filter $filter 
        foreach($Message in $Messages){
            Expand-ExtendedProperties -Item $Message
            Write-Output $Message
        }    
    }
}

function Invoke-GetReactionsDetailsOnMessage {
    [CmdletBinding()] 
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [String] $MailboxName,
        [Parameter(Position = 1, Mandatory = $false)]
        [String] $MessageId
    )
Process{
        $ExpandProperty = "singleValueExtendedProperties(`$filter=(id eq 'SystemTime {41F28F13-83F4-4114-A584-EEDB5A6B0BFF} name OwnerReactionTime') or (id eq 'String {41F28F13-83F4-4114-A584-EEDB5A6B0BFF} name OwnerReactionType')" 
        $ExpandProperty += "or (id eq 'Binary {41F28F13-83F4-4114-A584-EEDB5A6B0BFF} name ReactionsSummary') or (id eq 'Integer {41F28F13-83F4-4114-A584-EEDB5A6B0BFF} name ReactionsCount') or (id eq 'Binary {41F28F13-83F4-4114-A584-EEDB5A6B0BFF} name ReactionsHistory'))"
        $filter = "internetMessageId eq '$MessageId'"
        $Messages = Get-MgUserMessage -UserId $MailboxName -All -PageSize 999 -Select "Subject,receivedDateTime,singleValueExtendedProperties,InternetMessageId"  -ExpandProperty $ExpandProperty -Filter $filter 
        foreach($Message in $Messages){
            Expand-ExtendedProperties -Item $Message
            Write-Output $Message
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
                    "SystemTime {41F28F13-83F4-4114-A584-EEDB5A6B0BFF} name OwnerReactionTime"{
                        Add-Member -InputObject $Item -NotePropertyName "OwnerReactionTime" -NotePropertyValue $Prop.Value -Force
                    }
                    "String {41F28F13-83F4-4114-A584-EEDB5A6B0BFF} name OwnerReactionType"{
                        Add-Member -InputObject $Item -NotePropertyName "OwnerReactionType" -NotePropertyValue $Prop.Value -Force
                    } 
                    "Binary {41F28F13-83F4-4114-A584-EEDB5A6B0BFF} name ReactionsSummary"{
                        $ByteArrayValue =  [Convert]::FromBase64String($Prop.Value)
                        $ReactionSummary = Get-ReactionsFromSummary -SummaryBytes $ByteArrayValue
                        Add-Member -InputObject $Item -NotePropertyName "ReactionsSummary" -NotePropertyValue $ReactionSummary -Force
                    } 
                    "Binary {41F28F13-83F4-4114-A584-EEDB5A6B0BFF} name ReactionsHistory"{
                        $plainText = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($Prop.Value))
                        $jsonObject = $plainText | ConvertFrom-Json
                        Add-Member -InputObject $Item -NotePropertyName "ReactionsHistory" -NotePropertyValue $jsonObject -Force
                    } 
                    "Integer {41F28F13-83F4-4114-A584-EEDB5A6B0BFF} name ReactionsCount"{
                        Add-Member -InputObject $Item -NotePropertyName "ReactionsCount" -NotePropertyValue $Prop.Value -Force
                    }                    
                }
            }
        }
    }
}


function New-ReactionObject {
    param (
        [bool]$IsBcc,
        [string]$Name,
        [string]$Email,
        [string]$Type,
        [datetime]$DateTime
    )

    return [PSCustomObject]@{
        IsBcc    = $IsBcc 
        Name     = $Name
        Email    = $Email
        Type     = $Type
        DateTime = $DateTime
    }
}

function Convert-FileTimeToDateTime {
    param (
        [long]$FileTime
    )
    
    $Date1601 = Get-Date -Date "1601-01-01T00:00:00Z"
    if ($FileTime -ge 0) {
        return $Date1601.AddMinutes($FileTime).ToUniversalTime()
    }

    return (Get-Date).ToUniversalTime()
}

<#
.SYNOPSIS
    Get-ReactionsFromSummary takes the ReactionSummary property value, which is a binary array, 
    and parses the reaction for each user from this property. This PowerShell port is based on 
    the code from https://github.com/Sicos1977/MSGReader/blob/master/MsgReaderCore/Outlook/Reaction.cs

.AUTHOR
    Glen Scales <gscales@msgdevelop.com>

.ORIGINAL AUTHOR
    Sicos1977 (Kees van Spelde)
    Source: https://github.com/Sicos1977/MSGReader
    Notes: Original code in C#; this script adapts it to PowerShell with additional modifications.

.LICENSE
    This script is based on code originally written by Sicos1977 (Kees van Spelde).
    The original code is licensed under the MIT License:
    -----------------------------------------------------------------------
    MIT License

    Copyright (c) 2013-2025 Magic-Sessions

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in
    all copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
    THE SOFTWARE.
    -----------------------------------------------------------------------
#>
function Get-ReactionsFromSummary {
    param (
        [byte[]]$SummaryBytes
    )

    if (-not $SummaryBytes) {
        return @()
    }

    $BlobStream = [System.IO.MemoryStream]::new($SummaryBytes)
    $BlobReader = [System.IO.BinaryReader]::new($BlobStream)
    $Reactions = @()
    try {
        # Read version information
        $VersionPrefix = [char]$BlobReader.ReadByte()
        $VersionNumber = [System.UInt16]($BlobReader.ReadByte())

        if ($VersionPrefix -ne 'v' -or ($VersionNumber -lt 1 -or $VersionNumber -gt 255)) {
            throw "ReactionsSummary blob is not of new format: $VersionPrefix$VersionNumber"
        }        
        #Read the reaction counts
        $FoundEndOfLine = $false   
        while (-not $FoundEndOfLine) {
            $Ch = $BlobReader.ReadByte()            
            if([char]$Ch -eq '='){$BlobReader.ReadUInt16() | Out-Null};
            if([char]$Ch -eq [char]::MinValue){$FoundEndOfLine = $true}
        }        
        # Parse reactions
        while ($BlobReader.BaseStream.Position -lt $BlobReader.BaseStream.Length) {
            # Read IsBcc and SkinTone
            $bccVal = $BlobReader.ReadByte() 
            if (($bccVal -band 0x1) -ne 0) {
                $IsBcc = $true
            } else {
                $IsBcc = $false
            }
            
            $DateValue = $BlobReader.ReadInt32()           
            # Read the reaction timestamp
            $Date = Convert-FileTimeToDateTime -FileTime $DateValue

            # Define the Unicode to ReactionType mapping
            $UnicodeReverseLookup = @{
                "240,159,145,141" = "like"       # 👍
                "226,157,164"      = "heart"      # ❤️
                "240,159,142,137" = "celebrate"  # 🎉
                "240,159,152,134" = "laugh"      # 😆
                "240,159,152,178" = "surprised"  # 😲
                "240,159,152,162" = "sad"        # 😢
            }
                
            # Read ReactionTypeBytes from the stream
            $ReactionTypeBytes = @()
            do {
                $Ch = $BlobReader.ReadByte()
                if ([char]$Ch -ne ',') {
                    $ReactionTypeBytes += $Ch
                }
            } while ([char]$Ch -ne ',')
            
            # Convert the PowerShell array to a proper byte array
            $ReactionTypeBytes = [byte[]]$ReactionTypeBytes
            
            # Convert the byte array to a string for comparison
            $ReactionTypeKey = ($ReactionTypeBytes -join ',')
            
            
            # Lookup the reaction type in the UnicodeReverseLookup dictionary
            if ($UnicodeReverseLookup.ContainsKey($ReactionTypeKey)) {
                $ReactionType = $UnicodeReverseLookup[$ReactionTypeKey]
            } else {
                $ReactionType = "unknown"
            }
            
            # Read reactor name
            $ReactorNameBytes = @()
            do {
                $Ch = $BlobReader.ReadByte()
                if ($Ch -ne [char]','[0]) {
                    $ReactorNameBytes += $Ch
                }
            } while ($Ch -ne [char]','[0])
            $ReactorName = [System.Text.Encoding]::UTF8.GetString($ReactorNameBytes)

            # Read reactor email
            $ReactorEmailBytes = @()
            do {
                $Ch = $BlobReader.ReadByte()
                if ($Ch -ne [char]','[0]) {
                    $ReactorEmailBytes += $Ch
                }
            } while ($Ch -ne [char]','[0])
            $ReactorEmail = [System.Text.Encoding]::UTF8.GetString($ReactorEmailBytes)

            # Skip till the record separator
            do {
                $Ch = $BlobReader.ReadByte()
            } while ($Ch -ne [byte][char]0)

            # Add the reaction object
            $Reactions += (New-ReactionObject -IsBcc $IsBcc -Name $ReactorName -Email $ReactorEmail -Type $ReactionType -DateTime $Date)
        }

        return $Reactions
    } finally {
        $BlobReader.Close()
    }
}