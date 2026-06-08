<#
.SYNOPSIS
    Syncs Exchange Online users (with phone numbers) to a Graph contact folder,
    skipping contacts that already exist.

.PARAMETER UserId
    UPN or object ID of the mailbox to create contacts in.

.PARAMETER FolderId
    Contact folder ID to target. Defaults to the well-known 'Contacts' folder.

.PARAMETER BatchSize
    Number of contacts per Graph batch request. Max 20. Defaults to 20.

.EXAMPLE
    .\Sync-EXOUsersToContacts.ps1 -UserId gscales@datarumble.com

.EXAMPLE
    .\Sync-EXOUsersToContacts.ps1 -UserId gscales@datarumble.com -FolderId $folderId -Verbose
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$UserId,

    [Parameter()]
    [string]$FolderId = 'Contacts',

    [Parameter()]
    [ValidateRange(1, 20)]
    [int]$BatchSize = 20
)

# ── Constants ─────────────────────────────────────────────────────────────────

# Public strings namespace GUID — used for custom named extended properties
# that are not in a MAPI property set (i.e. not PSETID_Address)
$PublicStringsGuid = '00020329-0000-0000-C000-000000000046'
$GCSContactsTag    = 'GCSContacts'

# ── Retrieve source data from EXO ─────────────────────────────────────────────

Write-Verbose 'Retrieving users from Exchange Online...'

$ContactData = Get-User `
                   -ResultSize Unlimited `
                   -Filter     "Phone -ne `$null -or MobilePhone -ne `$null" |
               Select-Object DisplayName, FirstName, LastName, Title, Company,
                             Department, WindowsEmailAddress, Phone, MobilePhone

Write-Verbose "$($ContactData.Count) user(s) retrieved from EXO"

# ── Build index of existing contacts ─────────────────────────────────────────

Write-Verbose "Building existing contact index for $UserId..."

$ExistingContacts = Get-GCSContactIndex `
                        -UserId   $UserId `
                        -FolderId $FolderId `
                        -Verbose:$VerbosePreference

# ── Build property bags ───────────────────────────────────────────────────────

$BagsToCreate = [System.Collections.Generic.List[System.Collections.IDictionary]]::new()
$skippedCount = 0

foreach ($contact in $ContactData) {

    # Normalise the email key the same way Get-GCSContactIndex does
    $emailKey = $contact.WindowsEmailAddress?.ToString().Trim().ToLower()

    # Skip if no email address — can't safely deduplicate without one
    if ([string]::IsNullOrWhiteSpace($emailKey)) {
        Write-Warning "Skipping '$($contact.DisplayName)' — no WindowsEmailAddress"
        $skippedCount++
        continue
    }

    # Skip if already exists in the target folder
    if ($ExistingContacts.ContainsKey($emailKey)) {
        Write-Verbose "Already exists: $emailKey — skipping"
        $skippedCount++
        continue
    }

    $bag = New-GCSContactPropertyBag

    # ── Name fields ───────────────────────────────────────────────────────────
    if (-not [string]::IsNullOrWhiteSpace($contact.DisplayName)) { New-GCSContactProperty $bag Normal DisplayName $contact.DisplayName.Trim()
            $extSuectIbjd = "String 0x0037"
            New-GCSContactProperty $bag Extended $extSuectIbjd $contact.DisplayName.Trim()            
    }   
    if (-not [string]::IsNullOrWhiteSpace($contact.FirstName))   { New-GCSContactProperty $bag Normal GivenName   $contact.FirstName.Trim()   }
    if (-not [string]::IsNullOrWhiteSpace($contact.LastName))    { New-GCSContactProperty $bag Normal Surname     $contact.LastName.Trim()    }

    # ── Organisational fields ─────────────────────────────────────────────────
    if (-not [string]::IsNullOrWhiteSpace($contact.Title))       { New-GCSContactProperty $bag Normal JobTitle    $contact.Title.Trim()       }
    if (-not [string]::IsNullOrWhiteSpace($contact.Company))     { New-GCSContactProperty $bag Normal CompanyName $contact.Company.Trim()     }
    if (-not [string]::IsNullOrWhiteSpace($contact.Department))  { New-GCSContactProperty $bag Normal Department  $contact.Department.Trim()  }

    # ── Email ─────────────────────────────────────────────────────────────────
    if (-not [string]::IsNullOrWhiteSpace($contact.WindowsEmailAddress)) {
        New-GCSContactProperty $bag Email Email1.Address $contact.WindowsEmailAddress.ToString().Trim()
        New-GCSContactProperty $bag Email Email1.Name    $contact.DisplayName.Trim()
    }

    # ── Phone numbers (Unrolling MultiValuedProperties) ───────────────────────
    $mobileStr = if ($contact.MobilePhone) { @($contact.MobilePhone)[0].ToString().Trim() } else { $null }
    if (-not [string]::IsNullOrWhiteSpace($mobileStr)) { 
        New-GCSContactProperty $bag Phone MobilePhone $mobileStr 
    }

    $businessPhones = if ($contact.Phone) { 
        @($contact.Phone) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.ToString().Trim() } 
    } else { @() }
    
    foreach ($p in $businessPhones) {
        New-GCSContactProperty $bag Phone BusinessPhone $p
    }

    # ── Custom extended property (Public Strings namespace) ───────────────────
    # Id format: "String {00020329-0000-0000-C000-000000000046} Name GCSContacts"
    $extPropId = "String {$PublicStringsGuid} Name $GCSContactsTag"
    New-GCSContactProperty $bag Extended $extPropId "$GCSContactsTag"

    $BagsToCreate.Add($bag)
}

Write-Host "$($BagsToCreate.Count) contact(s) to create, $skippedCount skipped"

# ── Batch create ──────────────────────────────────────────────────────────────

if ($BagsToCreate.Count -gt 0) {

    if ($PSCmdlet.ShouldProcess("$UserId ($($BagsToCreate.Count) contacts)", 'Batch create contacts')) {

        $result = New-GCSContactBatch `
                      -UserId       $UserId `
                      -FolderId     $FolderId `
                      -PropertyBags $BagsToCreate `
                      -BatchSize    $BatchSize `
                      -Verbose:$VerbosePreference

        Write-Host ("Done — Created: $($result.SuccessCount)  " +
                    "Errors: $($result.ErrorCount)  " +
                    "Throttles: $($result.ThrottleCount)  " +
                    "Time: $($result.TimeToRun)s")
    }
} else {
    Write-Host 'Nothing to create — all contacts already exist in the target folder'
}