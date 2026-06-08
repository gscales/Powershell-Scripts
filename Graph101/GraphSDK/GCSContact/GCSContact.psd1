@{
    # ── Identity ──────────────────────────────────────────────────────────────
    ModuleVersion     = '1.5.0'
    GUID              = 'a3f7e2d1-5b84-4c19-9e36-0f2a7d8b1c45'
    Author            = 'Glen Scales'
    CompanyName       = 'MSGDevelop'
    Copyright         = '(c) Glen Scales. All rights reserved.'
    Description       = 'Flexible Microsoft Graph Contact Management module.'
    # ── PowerShell requirements ───────────────────────────────────────────────
    PowerShellVersion = '5.1'
    # ── Dependencies ─────────────────────────────────────────────────────────
    RequiredModules   = @()
    # ── Contents ──────────────────────────────────────────────────────────────
    RootModule        = 'GCSContact.psm1'
    FunctionsToExport = @(
        'New-GCSContactProperty',
        'New-GCSContactPropertyBag',
        'New-GCSContact',
        'New-GCSContactBatch',
        'Set-GCSContact',
        'Get-GCSContactProperty',
        'Get-GCSContactIndex',
        'New-GCSExtendedPropertyId',
        'New-GCSExtendedPropertyLid',
        'New-GCSExtendedPropertyTag',
        'ConvertTo-GCSContactBody'
    )
    AliasesToExport   = @(
        'SetProp',
        'ExtPropId',
        'ExtPropLid',
        'ExtPropTag'
    )
    VariablesToExport = @()
    CmdletsToExport   = @()
    
    # ── Metadata ──────────────────────────────────────────────────────────────
    PrivateData = @{
        PSData = @{
            Tags         = @('Graph', 'Exchange', 'Contact', 'MAPI', 'EWS', 'Batch')
            ProjectUri   = ''
            ReleaseNotes = '1.5.0 — Exported ConvertTo-GCSContactBody to resolve visibility scope issues. Added strict type casting and array unrolling.'
        }
    }
}