<#
.SYNOPSIS
    GraphContact — flexible Microsoft Graph contact management module.

.DESCRIPTION
    A PowerShell module that mirrors the SetProp / CreateContact pattern from
    Glen Scales' EWS contact creation script, adapted for the Microsoft Graph API.

    Exported cmdlets
    ────────────────
      New-GCSContactProperty     Add a typed property to a contact property bag
      New-GCSContact             Create a contact from a property bag
      Set-GCSContact             Update an existing contact from a property bag
      Get-GCSContactProperty     Retrieve a contact with extended properties
      Get-GCSContactIndex        Build a hashtable of contacts keyed on email address
      New-GCSExtendedPropertyId  Build a Graph singleValueExtendedProperty id string
      New-GCSExtendedPropertyLid Build a Graph id from a MAPI LID (hex)
      New-GCSExtendedPropertyTag Build a Graph id from a raw MAPI proptag
      New-GCSContactPropertyBag  Create an empty property bag hashtable
#>

# ── Module-level constants ────────────────────────────────────────────────────

# PSETID_Address — same GUID used in the EWS script as $AddressGuid
$script:PsetidAddress = '00062004-0000-0000-C000-000000000046'

# ── Lookup tables ─────────────────────────────────────────────────────────────

$script:NormalFieldMap = @{
    GivenName          = 'givenName'
    Surname            = 'surname'
    DisplayName        = 'displayName'
    FileAs             = 'fileAs'
    Subject            = 'Subject'     
    JobTitle           = 'jobTitle'
    CompanyName        = 'companyName'
    Department         = 'department'
    OfficeLocation     = 'officeLocation'
    BusinessHomePage   = 'businessHomePage'
    NickName           = 'nickName'
    Initials           = 'initials'
    MiddleName         = 'middleName'
    Generation         = 'generation'
    SpouseName         = 'spouseName'
    Manager            = 'manager'
    AssistantName      = 'assistantName'
    Profession         = 'profession'
    IMAddress          = 'imAddresses'
    BirthDay           = 'birthday'
    WeddingAnniversary = 'anniversary'
    Notes              = 'personalNotes'
}

$script:AddressKeyMap = @{
    Home     = 'homeAddress'
    Business = 'businessAddress'
    Other    = 'otherAddress'
}

$script:PhoneMap = @{
    MobilePhone      = @{ Field = 'mobilePhone';  Kind = 'direct' }
    BusinessPhone    = @{ Field = 'businessPhones'; Kind = 'array'  }
    BusinessPhone2   = @{ Field = 'businessPhones'; Kind = 'array'  }
    HomePhone        = @{ Field = 'homePhones';     Kind = 'array'  }
    HomePhone2       = @{ Field = 'homePhones';     Kind = 'array'  }
    AssistantPhone   = @{ Field = 'String 0x3A2E';  Kind = 'extended' }
    BusinessFax      = @{ Field = 'String 0x3A24';  Kind = 'extended' }
    HomeFax          = @{ Field = 'String 0x3A25';  Kind = 'extended' }
    OtherFax         = @{ Field = 'String 0x3A23';  Kind = 'extended' }
    Pager            = @{ Field = 'String 0x3A21';  Kind = 'extended' }
    PrimaryPhone     = @{ Field = 'String 0x3A1A';  Kind = 'extended' }
    RadioPhone       = @{ Field = 'String 0x3A1D';  Kind = 'extended' }
    CarPhone         = @{ Field = 'String 0x3A1E';  Kind = 'extended' }
    Isdn             = @{ Field = 'String 0x3A2D';  Kind = 'extended' }
    OtherTelephone   = @{ Field = 'String 0x3A1F';  Kind = 'extended' }
    Telex            = @{ Field = 'String 0x3A2C';  Kind = 'extended' }
    Callback         = @{ Field = 'String 0x3A02';  Kind = 'extended' }
    CompanyMainPhone = @{ Field = 'String 0x3A57';  Kind = 'extended' }
    TtyTddPhone      = @{ Field = 'String 0x3A4B';  Kind = 'extended' }
}

# ─────────────────────────────────────────────────────────────────────────────
# Extended property id helpers
# ─────────────────────────────────────────────────────────────────────────────

function New-GCSExtendedPropertyId {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$DataType,

        [Parameter(Mandatory, Position = 1)]
        [string]$Name,

        [Parameter(Position = 2)]
        [string]$Guid = $script:PsetidAddress
    )
    "$DataType {$Guid} Name $Name"
}

function New-GCSExtendedPropertyLid {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$DataType,

        [Parameter(Mandatory, Position = 1)]
        [int]$Lid,

        [Parameter(Position = 2)]
        [string]$Guid = $script:PsetidAddress
    )
    "$DataType {$Guid} Id 0x{0:X4}" -f $Lid
}

function New-GCSExtendedPropertyTag {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$DataType,

        [Parameter(Mandatory, Position = 1)]
        [int]$Tag
    )
    "$DataType 0x{0:X4}" -f $Tag
}

# ─────────────────────────────────────────────────────────────────────────────
# Property bag
# ─────────────────────────────────────────────────────────────────────────────

function New-GCSContactPropertyBag {
    [CmdletBinding()]
    [OutputType([System.Collections.IDictionary])]
    param()
    $h = @{}
    $h
}

function New-GCSContactProperty {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [System.Collections.IDictionary]$Bag,

        [Parameter(Mandatory, Position = 1)]
        [ValidateSet('Normal','Email','Phone','Address','Extended')]
        [string]$Type,

        [Parameter(Mandatory, Position = 2)]
        [object]$Name,

        [Parameter(Mandatory, Position = 3)]
        [object]$Value
    )

    if (-not $Bag) {
        if (-not (Get-Variable -Name ContactProps -Scope Script -ErrorAction SilentlyContinue)) {
            $script:ContactProps = [ordered]@{}
        }
        $Bag = $script:ContactProps
    }

    # Early Key Validation (Fail Fast Pattern)
    switch ($Type) {
        'Normal' {
            if (-not $script:NormalFieldMap.ContainsKey($Name)) {
                throw "Invalid Normal property: '$Name'. Allowed values: $($script:NormalFieldMap.Keys -join ', ')"
            }
        }
        'Phone' {
            if (-not $script:PhoneMap.ContainsKey($Name)) {
                throw "Invalid Phone property: '$Name'. Allowed values: $($script:PhoneMap.Keys -join ', ')"
            }
        }
        'Address' {
            $addressParts = $Name -split '\.'
            if ($addressParts.Count -ne 2 -or -not $script:AddressKeyMap.ContainsKey($addressParts[0])) {
                throw "Invalid Address property: '$Name'. Must be formatted as 'Type.Field' (e.g., Home.City). Allowed Types: $($script:AddressKeyMap.Keys -join ', ')"
            }
        }
        'Email' {
            if ($Name -notmatch '^Email(Address)?\d+\.(Address|Name)$') {
                throw "Invalid Email property: '$Name'. Must be formatted as 'Email1.Address' or 'Email1.Name'."
            }
        }
        'Extended' {
            if ($Name -notmatch '^(String|Integer|Boolean|Double|SystemTime|Binary|StringArray) ') {
                throw "Invalid Extended property ID: '$Name'. Must begin with a valid MAPI type (e.g., 'String {guid}...')."
            }
        }
    }

    $Bag[$Name] = [PSCustomObject]@{
        PropType = $Type
        Name     = $Name
        Value    = $Value
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Body builder (internal)
# ─────────────────────────────────────────────────────────────────────────────

function ConvertTo-GCSContactBody {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$PropertyBag
    )

    $body           = [ordered]@{}
    $emailSlots     = [ordered]@{}
    $businessPhones = [System.Collections.Generic.List[string]]::new()
    $homePhones     = [System.Collections.Generic.List[string]]::new()
    $addresses      = @{}
    $extProps       = [System.Collections.Generic.List[hashtable]]::new()

    foreach ($entry in $PropertyBag.Values) {

        switch ($entry.PropType) {

            'Normal' {
                $graphField = $script:NormalFieldMap[$entry.Name]
                if (-not $graphField) { continue }
                
                $val = $entry.Value
                if ($entry.Name -in @('BirthDay','WeddingAnniversary') -and $val -is [datetime]) {
                    $val = $val.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
                }
                $body[$graphField] = $val
            }

            'Email' {
                $cleanName = $entry.Name -replace '^EmailAddress','Email'
                $parts  = $cleanName.Split('.')
                $slot   = $parts[0] -replace '^Email',''  
                $field  = $parts[1]                       

                if (-not $emailSlots.Contains($slot)) {
                    $emailSlots[$slot] = @{ address = $null; name = $null }
                }
                switch ($field) {
                    'Address' { $emailSlots[$slot].address = $entry.Value }
                    'Name'    { $emailSlots[$slot].name    = $entry.Value }
                }
            }

            'Phone' {
                $mapEntry = $script:PhoneMap[$entry.Name]
                if (-not $mapEntry) { continue }
                
                switch ($mapEntry.Kind) {
                    'direct'   { $body[$mapEntry.Field] = $entry.Value }
                    'array'    {
                        if ($mapEntry.Field -eq 'businessPhones') { $businessPhones.Add($entry.Value) }
                        else                                      { $homePhones.Add($entry.Value)     }
                    }
                    'extended' { $extProps.Add(@{ id = $mapEntry.Field; value = $entry.Value.ToString() }) }
                }
            }

            'Address' {
                $parts      = $entry.Name.Split('.')
                $addressKey = $script:AddressKeyMap[$parts[0]]
                if (-not $addressKey) { continue }
                
                $subField = switch ($parts[1]) {
                    'CountryOrRegion' { 'countryOrRegion' }
                    'PostalCode'      { 'postalCode'      }
                    'Street'          { 'street'          }
                    'City'            { 'city'            }
                    'State'           { 'state'           }
                    default           { $parts[1].Substring(0,1).ToLower() + $parts[1].Substring(1) }
                }
                if (-not $addresses.ContainsKey($addressKey)) { $addresses[$addressKey] = @{} }
                $addresses[$addressKey][$subField] = $entry.Value
            }

            'Extended' {
                $extId = $entry.Name.ToString()
                $rawValue = $entry.Value
                $typedValue = $null

                # Enforce JSON-compatible strict primitive casting for Graph singleValueExtendedProperties
                if ($extId -match '^Integer ') {
                    $typedValue = [int]$rawValue
                } elseif ($extId -match '^Boolean ') {
                    if ($rawValue -is [bool]) { $typedValue = $rawValue }
                    else { $typedValue = [convert]::ToBoolean($rawValue.ToString()) }
                } elseif ($extId -match '^Double ') {
                    $typedValue = [double]$rawValue
                } elseif ($extId -match '^SystemTime ') {
                    if ($rawValue -is [datetime]) {
                        $typedValue = $rawValue.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
                    } else {
                        $typedValue = [datetime]::Parse($rawValue.ToString()).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
                    }
                } else {
                    $typedValue = $rawValue.ToString()
                }

                $extProps.Add(@{ id = $extId; value = $typedValue })
            }
        }
    }

    # Collapse distinct email slots into Graph payload objects
    if ($emailSlots.Count -gt 0) {
        $emailArray = @()
        foreach ($slot in ($emailSlots.Keys | Sort-Object)) {
            $s        = $emailSlots[$slot]
            $dispName = if ($s.name) { $s.name } else { $s.address }
            $emailObj = @{}
            if ($s.address) { $emailObj.address = $s.address }
            if ($s.name)    { $emailObj.name    = $s.name    }
            $emailArray += $emailObj

            # Match EWS sync pattern behaviour ensuring Outlook can resolve properties natively
            $n = $slot
            $extProps.Add(@{ id = (New-GCSExtendedPropertyId String "dispidEmail${n}AddrType");            value = 'SMTP'        })
            $extProps.Add(@{ id = (New-GCSExtendedPropertyId String "dispidEmail${n}DisplayName");         value = $dispName     })
            $extProps.Add(@{ id = (New-GCSExtendedPropertyId String "dispidEmail${n}EmailAddress");        value = $s.address    })
            $extProps.Add(@{ id = (New-GCSExtendedPropertyId String "dispidEmail${n}OriginalDisplayName"); value = $s.address    })
        }
        $body.emailAddresses = $emailArray
    }

    if ($businessPhones.Count -gt 0) { $body.businessPhones = $businessPhones.ToArray() }
    if ($homePhones.Count     -gt 0) { $body.homePhones     = $homePhones.ToArray()     }
    foreach ($key in $addresses.Keys) { $body[$key] = $addresses[$key] }
    if ($extProps.Count -gt 0) { $body.singleValueExtendedProperties = $extProps.ToArray() }

    $body
}

# ─────────────────────────────────────────────────────────────────────────────
# Public cmdlets
# ─────────────────────────────────────────────────────────────────────────────

function New-GCSContact {
    <#
    .SYNOPSIS
        Creates a new contact in a user's mailbox from a property bag.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$UserId,

        [Parameter(Position = 1)]
        [System.Collections.IDictionary]$PropertyBag,

        [Parameter()]
        [string]$FolderId
    )

    if (-not $PropertyBag) {
        if (-not (Get-Variable -Name ContactProps -Scope Script -ErrorAction SilentlyContinue)) {
            throw "No PropertyBag supplied and no script-level ContactProps found. Call New-GCSContactPropertyBag first."
        }
        $PropertyBag = $script:ContactProps
    }

    $body = ConvertTo-GCSContactBody -PropertyBag $PropertyBag

    $displayName = if ($body.Contains('displayName')) { $body.displayName }
                   elseif ($body.Contains('fileAs'))   { $body.fileAs }
                   else                                { 'New Contact' }

    if ($PSCmdlet.ShouldProcess($displayName, 'Create Graph contact')) {
        try {
            $params = @{
                UserId        = $UserId
                BodyParameter = $body
                ErrorAction   = 'Stop'
            }
            if ($FolderId) {
                $contact = New-MgUserContactFolderContact @params -ContactFolderId $FolderId
            } else {
                $contact = New-MgUserContact @params
            }
            Write-Verbose "Contact created: $($contact.DisplayName)  (id: $($contact.Id))"
            $contact
        }
        catch {
            Write-Error "Failed to create contact '$displayName': $_"
        }
    }
}

function Set-GCSContact {
    <#
    .SYNOPSIS
        Updates an existing contact from a property bag via PATCH.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$UserId,

        [Parameter(Mandatory, Position = 1)]
        [string]$ContactId,

        [Parameter(Mandatory, Position = 2)]
        [System.Collections.IDictionary]$PropertyBag
    )

    $body = ConvertTo-GCSContactBody -PropertyBag $PropertyBag

    if ($PSCmdlet.ShouldProcess($ContactId, 'Update Graph contact')) {
        try {
            Update-MgUserContact `
                -UserId    $UserId `
                -ContactId $ContactId `
                -BodyParameter $body `
                -ErrorAction Stop
            Write-Verbose "Contact $ContactId updated"
        }
        catch {
            Write-Error "Failed to update contact '$ContactId': $_"
        }
    }
}

function Get-GCSContactProperty {
    <#
    .SYNOPSIS
        Retrieves a contact and expands one or more extended properties.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$UserId,

        [Parameter(Mandatory, Position = 1)]
        [string]$ContactId,

        [Parameter()]
        [string[]]$ExtendedPropertyIds
    )

    $expandParts = @()
    foreach ($id in $ExtendedPropertyIds) {
        $expandParts += "singleValueExtendedProperties(`$filter=id eq '$id')"
    }

    $params = @{
        UserId    = $UserId
        ContactId = $ContactId
    }
    if ($expandParts.Count -gt 0) {
        $params.ExpandProperty = $expandParts -join ','
    }

    Get-MgUserContact @params
}

function Get-GCSContactIndex {
    <#
    .SYNOPSIS
        Builds a hashtable of existing contacts in a mailbox keyed on lowercase email address.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserId,

        [Parameter()]
        [string]$FolderId = 'Contacts'
    )

    Write-Verbose "Fetching existing contacts to build evaluation index..."
    $index = @{}

    $params = @{
        UserId   = $UserId
        All      = $true
        Property = @('id', 'displayName', 'emailAddresses')
    }

    try {
        $contacts = if ($FolderId -and $FolderId -ne 'Contacts') {
            Get-MgUserContactFolderContact @params -ContactFolderId $FolderId
        } else {
            Get-MgUserContact @params
        }

        foreach ($c in $contacts) {
            if ($null -eq $c.EmailAddresses) { continue }
            
            foreach ($email in $c.EmailAddresses) {
                if (-not [string]::IsNullOrWhiteSpace($email.Address)) {
                    $key = $email.Address.Trim().ToLower()
                    if (-not $index.ContainsKey($key)) {
                        $index[$key] = $c
                    }
                }
            }
        }
        Write-Verbose "Index built containing $($index.Count) unique email mappings."
    }
    catch {
        Write-Error "Failed to build contact index mapping table: $_"
    }

    return $index
}