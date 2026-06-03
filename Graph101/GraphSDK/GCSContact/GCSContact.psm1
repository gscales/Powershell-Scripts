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
    # First-class Graph fields
    MobilePhone      = @{ Field = 'mobilePhone';    Kind = 'direct' }
    BusinessPhone    = @{ Field = 'businessPhones'; Kind = 'array'  }
    BusinessPhone2   = @{ Field = 'businessPhones'; Kind = 'array'  }
    HomePhone        = @{ Field = 'homePhones';     Kind = 'array'  }
    HomePhone2       = @{ Field = 'homePhones';     Kind = 'array'  }
    # No first-class Graph field — stored as MAPI proptag extended properties
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
    <#
    .SYNOPSIS
        Builds a Graph singleValueExtendedProperty id for a MAPI named property
        in the PSETID_Address property set.

    .PARAMETER DataType
        MAPI data type string: String, Integer, Double, Boolean, DateTime, etc.

    .PARAMETER Name
        Canonical property name, e.g. dispidEmail1AddrType

    .PARAMETER Guid
        Optional. Property set GUID. Defaults to PSETID_Address
        ({00062004-0000-0000-C000-000000000046}).

    .EXAMPLE
        New-GCSExtendedPropertyId -DataType String -Name dispidEmail1AddrType
        # Returns: "String {00062004-0000-0000-C000-000000000046} Name dispidEmail1AddrType"

    .EXAMPLE
        ExtPropId String dispidEmail1AddrType   # alias form
    #>
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
    <#
    .SYNOPSIS
        Builds a Graph singleValueExtendedProperty id from a MAPI LID (numeric
        property identifier) in the PSETID_Address property set.

    .PARAMETER DataType
        MAPI data type string.

    .PARAMETER Lid
        MAPI Long ID (LID) as an integer. e.g. 0x8080 for PidLidEmail1DisplayName.

    .PARAMETER Guid
        Optional. Defaults to PSETID_Address.

    .EXAMPLE
        New-GCSExtendedPropertyLid -DataType String -Lid 0x8080
        # Returns: "String {00062004-0000-0000-C000-000000000046} Id 0x8080"
    #>
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
    <#
    .SYNOPSIS
        Builds a Graph singleValueExtendedProperty id from a raw MAPI property
        tag (for properties that are not in a named-property set).

    .PARAMETER DataType
        MAPI data type string.

    .PARAMETER Tag
        MAPI property tag as an integer. e.g. 0x3A2C for PR_TELEX_NUMBER.

    .EXAMPLE
        New-GCSExtendedPropertyTag -DataType String -Tag 0x3A2C
        # Returns: "String 0x3A2C"
    #>
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
    <#
    .SYNOPSIS
        Creates a new empty contact property bag hashtable.

    .DESCRIPTION
        Returns an ordered hashtable ready to be populated with
        New-GCSContactProperty calls and passed to New-GCSContact or
        Set-GCSContact.

    .EXAMPLE
        $bag = New-GCSContactPropertyBag
        New-GCSContactProperty -Bag $bag -Type Normal -Name GivenName -Value "John"
        New-GCSContact -UserId user@contoso.com -PropertyBag $bag
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.IDictionary])]
    param()
    # Return a plain Hashtable — [ordered]@{} returns an OrderedDictionary
    # which does not bind to [System.Collections.Hashtable] parameters in PowerShell,
    # causing $bag to appear empty when passed to New-GCSContactProperty.
    $h = @{}
    $h
}

function New-GCSContactProperty {
    <#
    .SYNOPSIS
        Adds a typed property entry to a contact property bag.

    .DESCRIPTION
        Equivalent to SetProp in the original EWS script.  Call this once per
        property, then pass the accumulated bag to New-GCSContact or
        Set-GCSContact.

    .PARAMETER Bag
        The property bag hashtable created by New-GCSContactPropertyBag.
        When omitted, the cmdlet writes to $script:ContactProps which is
        created automatically if it does not exist.

    .PARAMETER Type
        Property category:
          Normal   – top-level Graph contact field (GivenName, Surname, etc.)
          Email    – email address slot  (Email1.Address, Email2.Name, etc.)
          Phone    – phone number  (MobilePhone, BusinessPhone, etc.)
          Address  – physical address  (Home.City, Business.Street, etc.)
          Extended – singleValueExtendedProperty (use ExtPropId/Lid/Tag for Name)

    .PARAMETER Name
        Property name within its type category, or an extended property id string.

    .PARAMETER Value
        Property value.

    .EXAMPLE
        $bag = New-GCSContactPropertyBag
        New-GCSContactProperty $bag Normal GivenName        "John"
        New-GCSContactProperty $bag Normal Surname          "Doe"
        New-GCSContactProperty $bag Email  Email1.Address   "john@example.com"
        New-GCSContactProperty $bag Phone  MobilePhone      "0400000000"
        New-GCSContactProperty $bag Address Home.City       "Sydney"

    .EXAMPLE
        # Extended property (MAPI named property)
        New-GCSContactProperty $bag Extended (ExtPropId String dispidEmail1AddrType) "SMTP"
    #>
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
    <#
    .SYNOPSIS
        Converts a property bag into the JSON body hashtable for a Graph API call.
    #>
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
                if (-not $graphField) {
                    Write-Warning "GCSContact: Normal property '$($entry.Name)' has no Graph mapping — skipping"
                    continue
                }
                $val = $entry.Value
                if ($entry.Name -in @('BirthDay','WeddingAnniversary') -and $val -is [datetime]) {
                    $val = $val.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
                }
                $body[$graphField] = $val
            }

            'Email' {
                # Accepts  "Email1.Address" / "Email1.Name"
                # also     "EmailAddress1.Address" (EWS form) — strip prefix
                $cleanName = $entry.Name -replace '^EmailAddress','Email'
                $parts  = $cleanName.Split('.')
                $slot   = $parts[0] -replace '^Email',''   # "1","2","3"
                $field  = $parts[1]                         # "Address" or "Name"

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
                if (-not $mapEntry) {
                    Write-Warning "GCSContact: Phone key '$($entry.Name)' unknown — skipping"
                    continue
                }
                switch ($mapEntry.Kind) {
                    'direct'   { $body[$mapEntry.Field] = $entry.Value }
                    'array'    {
                        if ($mapEntry.Field -eq 'businessPhones') { $businessPhones.Add($entry.Value) }
                        else                                       { $homePhones.Add($entry.Value)     }
                    }
                    'extended' { $extProps.Add(@{ id = $mapEntry.Field; value = $entry.Value }) }
                }
            }

            'Address' {
                $parts      = $entry.Name.Split('.')
                $addressKey = $script:AddressKeyMap[$parts[0]]
                if (-not $addressKey) {
                    Write-Warning "GCSContact: Address key '$($parts[0])' unknown — skipping"
                    continue
                }
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
                $extProps.Add(@{ id = $entry.Name.ToString(); value = $entry.Value.ToString() })
            }
        }
    }

    # Collapse email slots
    if ($emailSlots.Count -gt 0) {
        $emailArray = @()
        foreach ($slot in ($emailSlots.Keys | Sort-Object)) {
            $s        = $emailSlots[$slot]
            $dispName = if ($s.name) { $s.name } else { $s.address }
            $emailObj = @{}
            if ($s.address) { $emailObj.address = $s.address }
            if ($s.name)    { $emailObj.name    = $s.name    }
            $emailArray += $emailObj

            # Inject MAPI Email<N> extended properties so Outlook renders correctly
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

    .DESCRIPTION
        Converts the property bag built with New-GCSContactProperty into a
        single POST request to the Microsoft Graph contacts endpoint, including
        all standard fields and MAPI extended properties in one call.

    .PARAMETER UserId
        UPN or object ID of the target mailbox (e.g. user@contoso.com).

    .PARAMETER PropertyBag
        Hashtable built with New-GCSContactPropertyBag and populated with
        New-GCSContactProperty.  When omitted, uses $script:ContactProps.

    .PARAMETER FolderId
        Optional. Contact folder ID to create the contact in. When omitted,
        the contact is created in the default Contacts folder.

    .OUTPUTS
        Microsoft.Graph.PowerShell.Models.MicrosoftGraphContact

    .EXAMPLE
        $bag = New-GCSContactPropertyBag
        New-GCSContactProperty $bag Normal GivenName       "John"
        New-GCSContactProperty $bag Normal Surname         "Doe"
        New-GCSContactProperty $bag Email  Email1.Address  "john@example.com"
        New-GCSContact -UserId user@contoso.com -PropertyBag $bag
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
            throw "No PropertyBag supplied and no script-level ContactProps found. " +
                  "Call New-GCSContactPropertyBag first."
        }
        $PropertyBag = $script:ContactProps
    }

    $body = ConvertTo-GCSContactBody -PropertyBag $PropertyBag

    $displayName = if ($body.Contains('displayName')) { $body.displayName }
                   elseif ($body.Contains('fileAs'))   { $body.fileAs }
                   else                               { 'New Contact' }

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
        Updates an existing contact from a property bag.

    .DESCRIPTION
        Issues a PATCH request against an existing contact, applying only the
        properties present in the property bag.

    .PARAMETER UserId
        UPN or object ID of the target mailbox.

    .PARAMETER ContactId
        ID of the contact to update (from New-GCSContact or Get-MgUserContact).

    .PARAMETER PropertyBag
        Hashtable of properties to update.  Unspecified properties are unchanged.

    .EXAMPLE
        $bag = New-GCSContactPropertyBag
        New-GCSContactProperty $bag Extended (ExtPropId String dispidEmail1DisplayName) "dD"
        Set-GCSContact -UserId user@contoso.com -ContactId $contact.Id -PropertyBag $bag
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

    .PARAMETER UserId
        UPN or object ID of the mailbox.

    .PARAMETER ContactId
        Contact ID to retrieve.

    .PARAMETER ExtendedPropertyIds
        One or more extended property id strings (from ExtPropId / ExtPropLid /
        ExtPropTag) to expand on the returned contact object.

    .EXAMPLE
        Get-GCSContactProperty `
            -UserId    user@contoso.com `
            -ContactId $contact.Id `
            -ExtendedPropertyIds (ExtPropId String dispidEmail1DisplayName),
                                 (ExtPropId String dispidEmail1EmailAddress)
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

# ─────────────────────────────────────────────────────────────────────────────
# Batch contact creation
# ─────────────────────────────────────────────────────────────────────────────

function New-GCSContactBatch {
    <#
    .SYNOPSIS
        Creates multiple contacts in batches of up to 20 using the Graph
        JSON batch endpoint ($batch).

    .DESCRIPTION
        Takes a collection of property bags (each built with
        New-GCSContactPropertyBag / New-GCSContactProperty) and submits them
        to the Microsoft Graph \$batch endpoint in groups of up to 20 requests.

        Using \$batch reduces round-trips from one HTTP call per contact to one
        call per 20 contacts, which is significantly faster for bulk imports.

        Authentication is handled by the existing Invoke-MgGraphRequest session
        (Connect-MgGraph) — no manual token management is required.

        Throttling (HTTP 429) is handled automatically: when any response in a
        batch returns 429 the function honours the Retry-After header and
        re-submits the failed requests.

    .PARAMETER UserId
        UPN or object ID of the target mailbox (e.g. user@contoso.com).

    .PARAMETER PropertyBags
        An array or list of property bag hashtables, each built with
        New-GCSContactPropertyBag and populated with New-GCSContactProperty.

    .PARAMETER FolderId
        Optional. Contact folder ID to create contacts in.  When omitted,
        contacts are created in the default Contacts folder.

    .PARAMETER BatchSize
        Number of requests per batch call.  Defaults to 20, which is the
        Graph \$batch maximum.  Reduce for troubleshooting.

    .OUTPUTS
        PSCustomObject with properties:
          SuccessCount  – number of contacts successfully created
          ErrorCount    – number of failed requests
          ThrottleCount – number of 429 throttle responses handled
          TimeToRun     – total elapsed seconds

    .EXAMPLE
        $bags = 1..50 | ForEach-Object {
            $b = New-GCSContactPropertyBag
            New-GCSContactProperty $b Normal GivenName  "User$_"
            New-GCSContactProperty $b Normal Surname    "Test"
            New-GCSContactProperty $b Email Email1.Address "user$_@contoso.com"
            $b
        }
        New-GCSContactBatch -UserId admin@contoso.com -PropertyBags $bags

    .EXAMPLE
        # Create into a specific contact folder
        New-GCSContactBatch -UserId admin@contoso.com -PropertyBags $bags -FolderId $folderId
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$UserId,

        [Parameter(Mandatory, Position = 1)]
        [System.Collections.IList]$PropertyBags,

        [Parameter()]
        [string]$FolderId,

        [Parameter()]
        [ValidateRange(1, 20)]
        [int]$BatchSize = 20
    )

    $report = [PSCustomObject]@{
        SuccessCount  = 0
        ErrorCount    = 0
        ThrottleCount = 0
        TimeToRun     = 0
    }
    $startTime = Get-Date

    # Build the relative URL for each contact POST
    $contactUrl = if ($FolderId) {
        "/users/$UserId/contactFolders/$FolderId/contacts"
    } else {
        "/users/$UserId/contacts"
    }

    # Convert every property bag to a Graph body hashtable up front
    $bodies = $PropertyBags | ForEach-Object {
        ConvertTo-GCSContactBody -PropertyBag $_
    }

    # ── Slice into batches of $BatchSize ─────────────────────────────────────
    $total     = $bodies.Count
    $processed = 0

    while ($processed -lt $total) {

        $slice = $bodies[$processed..([Math]::Min($processed + $BatchSize - 1, $total - 1))]
        $processed += $slice.Count

        if ($PSCmdlet.ShouldProcess("$UserId ($($slice.Count) contacts)", 'Batch create contacts')) {
            Invoke-GCSBatchSlice `
                -UserId     $UserId `
                -Slice      $slice `
                -ContactUrl $contactUrl `
                -Report     $report
        }
    }

    $report.TimeToRun = [Math]::Round(
        (New-TimeSpan -Start $startTime -End (Get-Date)).TotalSeconds, 2)

    Write-Verbose ("Batch complete — Success: $($report.SuccessCount)  " +
                   "Errors: $($report.ErrorCount)  " +
                   "Throttles: $($report.ThrottleCount)  " +
                   "Elapsed: $($report.TimeToRun)s")
    $report
}

# ── Internal: submit one batch slice, handling 429 retry ─────────────────────

function Invoke-GCSBatchSlice {
    <#
    .SYNOPSIS
        Submits a single slice of up to 20 contact POST requests via \$batch
        and processes the responses.  Retries throttled requests automatically.
    #>
    [CmdletBinding()]
    param(
        [string]                          $UserId,
        [object[]]                        $Slice,
        [string]                          $ContactUrl,
        [PSCustomObject]                  $Report
    )

    # Build the \$batch request body
    $requests = for ($i = 0; $i -lt $Slice.Count; $i++) {
        @{
            id      = ($i + 1).ToString()
            method  = 'POST'
            url     = $ContactUrl
            headers = @{ 'Content-Type' = 'application/json' }
            body    = $Slice[$i]
        }
    }

    $batchBody = @{ requests = $requests }

    # Invoke-MgGraphRequest uses the SDK's active session — no manual token needed
    try {
        $response = Invoke-MgGraphRequest `
            -Method      POST `
            -Uri         'https://graph.microsoft.com/v1.0/$batch' `
            -Body        ($batchBody | ConvertTo-Json -Depth 20 -Compress) `
            -ContentType 'application/json' `
            -ErrorAction Stop
    }
    catch {
        Write-Error "Batch request failed: $_"
        $Report.ErrorCount += $Slice.Count
        return
    }

    if (-not $response.responses) {
        Write-Error 'Batch returned no responses'
        $Report.ErrorCount += $Slice.Count
        return
    }

    # Track which items need a retry due to throttling
    $retryBodies    = [System.Collections.Generic.List[object]]::new()
    $retryAfterSecs = 0

    foreach ($r in $response.responses) {

        $status = [int]$r.status

        switch ($status) {

            201 {
                # Created successfully
                $Report.SuccessCount++
                $displayName = if ($r.body.displayName) { $r.body.displayName }
                               else                      { "(id $($r.id))" }
                Write-Verbose "Contact created: $displayName"
            }

            429 {
                # Throttled — queue for retry
                $Report.ThrottleCount++
                $retryAfterSecs = [Math]::Max(
                    $retryAfterSecs,
                    [int]($r.headers.'Retry-After' ?? 10))

                # The id is 1-based and matches the slice index
                $originalIndex = [int]$r.id - 1
                if ($originalIndex -ge 0 -and $originalIndex -lt $Slice.Count) {
                    $retryBodies.Add($Slice[$originalIndex])
                }
                Write-Warning "Request $($r.id) throttled (429) — will retry after ${retryAfterSecs}s"
            }

            default {
                $Report.ErrorCount++
                $errorMsg = if ($r.body.error.message) { $r.body.error.message }
                            else                        { "HTTP $status" }
                Write-Warning "Request $($r.id) failed: $errorMsg"
            }
        }
    }

    # ── Throttle retry ────────────────────────────────────────────────────────
    if ($retryBodies.Count -gt 0) {
        Write-Verbose "Sleeping ${retryAfterSecs}s before retrying $($retryBodies.Count) throttled request(s)"
        Start-Sleep -Seconds $retryAfterSecs

        # Recursive call — retried items are a new slice so ids reset to 1-based
        Invoke-GCSBatchSlice `
            -UserId     $UserId `
            -Slice      $retryBodies.ToArray() `
            -ContactUrl $ContactUrl `
            -Report     $Report
    }
}


Export-ModuleMember -Function @(
    'New-GCSContactProperty',
    'New-GCSContactPropertyBag',
    'New-GCSContact',
    'New-GCSContactBatch',
    'Set-GCSContact',
    'Get-GCSContactProperty',
    'New-GCSExtendedPropertyId',
    'New-GCSExtendedPropertyLid',
    'New-GCSExtendedPropertyTag'
) -Alias @(
    'SetProp',
    'ExtPropId',
    'ExtPropLid',
    'ExtPropTag'
)