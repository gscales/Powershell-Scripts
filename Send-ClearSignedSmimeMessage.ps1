<#
.SYNOPSIS
    Builds and sends a clear-signed S/MIME message via Microsoft Graph.

.DESCRIPTION
    Loads a DER/Base64 .cer file, finds the matching private key in the Windows
    certificate store, then builds a multipart/mixed MIME message and clear-signs
    it using .NET's System.Security.Cryptography.Pkcs.SignedCms — which supports
    both legacy CAPI and modern CNG private keys on all .NET runtimes.

    The signed EML is saved locally and then sent via the Microsoft Graph
    sendMail API (POST /v1.0/users/{from}/sendMail) using the raw MIME approach
    described at https://learn.microsoft.com/en-us/graph/outlook-send-mime-message.

    Prerequisites:
      - Mailozaurr module 1.0.7  : Install-Module Mailozaurr -RequiredVersion 1.0.7
      - Microsoft.Graph module   : Install-Module Microsoft.Graph
      - An Entra ID app (or delegated sign-in) with Mail.Send permission granted.

.PARAMETER CerPath
    Full path to the .cer file (public certificate, DER or Base64 encoded).
    Used to locate the matching certificate with private key in CurrentUser\My.

.PARAMETER FromEmail
    Sender email address. Must match the mailbox used when connecting to Graph.

.PARAMETER ToEmail
    Recipient email address.

.PARAMETER Subject
    Email subject line.

.PARAMETER Body
    HTML body of the message.

.PARAMETER AttachmentPath
    Optional path to a file to attach. MIME content-type is detected automatically
    from the file extension.

.PARAMETER OutputDirectory
    Directory where the .eml will be saved before sending. Defaults to .\output.

.PARAMETER GraphClientId
    Entra ID application (client) ID used to authenticate to Microsoft Graph.

.PARAMETER GraphTenantId
    Entra ID tenant ID.

.PARAMETER GraphClientSecret
    Client secret for app-only (client credentials) authentication.
    If omitted, interactive delegated authentication is used instead.

.PARAMETER SaveOnly
    If specified, builds and saves the EML but does not send via Graph.

.EXAMPLE
    # Interactive (delegated) auth — prompts for sign-in
    .\Send-ClearSignedSmimeMessage.ps1 `
        -CerPath       "C:\certs\test-smime.cer" `
        -FromEmail     "sender@contoso.com" `
        -ToEmail       "recipient@contoso.com" `
        -Subject       "Hello from PowerShell" `
        -Body          "<p>This message is clear-signed.</p>" `
        -GraphClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -GraphTenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

.EXAMPLE
    # App-only auth with client secret
    .\Send-ClearSignedSmimeMessage.ps1 `
        -CerPath             "C:\certs\test-smime.cer" `
        -FromEmail           "sender@contoso.com" `
        -ToEmail             "recipient@contoso.com" `
        -Subject             "Signed message" `
        -Body                "<p>Hello!</p>" `
        -AttachmentPath      "C:\docs\report.pdf" `
        -GraphClientId       "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -GraphTenantId       "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -GraphClientSecret   "your-client-secret"

.EXAMPLE
    # Build and save EML only, no Graph send
    .\Send-ClearSignedSmimeMessage.ps1 `
        -CerPath         "C:\certs\test-smime.cer" `
        -FromEmail       "sender@contoso.com" `
        -ToEmail         "recipient@contoso.com" `
        -Subject         "Test" `
        -Body            "<p>Test</p>" `
        -OutputDirectory "C:\output" `
        -SaveOnly
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $CerPath,
    [Parameter(Mandatory)][string] $FromEmail,
    [Parameter(Mandatory)][string] $ToEmail,
    [string] $Subject            = "",
    [string] $Body               = "",
    [string] $AttachmentPath     = "",
    [string] $OutputDirectory    = ".\output",
    [string] $GraphClientId      = "",
    [string] $GraphTenantId      = "",
    [string] $GraphClientSecret  = "",
    [switch] $SaveOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# 1. Locate Mailozaurr and load MimeKit
# ---------------------------------------------------------------------------
$module = Get-Module -ListAvailable -Name Mailozaurr |
          Sort-Object Version -Descending |
          Select-Object -First 1

if (-not $module) {
    throw "Mailozaurr module not found. Run: Install-Module Mailozaurr -RequiredVersion 1.0.7"
}

$libPath    = Join-Path $module.ModuleBase 'Lib\Default'
$mimeKitDll = Join-Path $libPath 'MimeKit.dll'
if (-not (Test-Path $mimeKitDll)) {
    throw "MimeKit.dll not found at: $mimeKitDll"
}
Add-Type -Path $mimeKitDll
Write-Verbose "MimeKit loaded from: $mimeKitDll"

# ---------------------------------------------------------------------------
# 2. Load the .cer and locate the matching cert + private key in the Windows store
# ---------------------------------------------------------------------------
if (-not (Test-Path $CerPath)) {
    throw ".cer file not found: $CerPath"
}

$cerCert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CerPath)
Write-Verbose "Loaded .cer: Subject=$($cerCert.Subject) Thumbprint=$($cerCert.Thumbprint)"

$certStore = [System.Security.Cryptography.X509Certificates.X509Store]::new(
    [System.Security.Cryptography.X509Certificates.StoreName]::My,
    [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser
)
$certStore.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)

$signingCert = $certStore.Certificates |
               Where-Object { $_.Thumbprint -eq $cerCert.Thumbprint -and $_.HasPrivateKey } |
               Select-Object -First 1

$certStore.Close()

if (-not $signingCert) {
    throw "No certificate with a private key matching thumbprint '$($cerCert.Thumbprint)' found in CurrentUser\My."
}

Write-Verbose "Found signing cert in store: $($signingCert.Subject)"

# ---------------------------------------------------------------------------
# 3. Build the inner MIME body using MimeKit (HTML + optional attachment)
# ---------------------------------------------------------------------------
$null = [System.IO.Directory]::CreateDirectory($OutputDirectory)

$htmlBody      = [MimeKit.TextPart]::new("html")
$htmlBody.Text = $Body

if ($AttachmentPath) {
    if (-not (Test-Path $AttachmentPath)) {
        throw "Attachment file not found: $AttachmentPath"
    }

    $fileName    = [System.IO.Path]::GetFileName($AttachmentPath)
    $contentType = [MimeKit.ContentType]::Parse(
                       [MimeKit.MimeTypes]::GetMimeType($fileName))

    $attachStream = [System.IO.File]::OpenRead($AttachmentPath)
    $attachment   = [MimeKit.MimePart]::new($contentType.MediaType, $contentType.MediaSubtype)
    $attachment.Content                 = [MimeKit.MimeContent]::new($attachStream)
    $attachment.ContentDisposition      = [MimeKit.ContentDisposition]::new(
                                              [MimeKit.ContentDisposition]::Attachment)
    $attachment.ContentTransferEncoding = [MimeKit.ContentEncoding]::Base64
    $attachment.FileName                = $fileName

    $innerBody = [MimeKit.Multipart]::new("mixed")
    $innerBody.Add($htmlBody)
    $innerBody.Add($attachment)
    Write-Verbose "Attachment added: $fileName ($($contentType.MediaType)/$($contentType.MediaSubtype))"
} else {
    $innerBody = $htmlBody
}

# Serialise the inner body to bytes — these are exactly the bytes that get signed.
$innerStream = [System.IO.MemoryStream]::new()
$innerBody.WriteTo($innerStream)
$innerBytes = $innerStream.ToArray()
$innerStream.Dispose()

# ---------------------------------------------------------------------------
# 4. Sign with System.Security.Cryptography.Pkcs.SignedCms (detached signature)
#    Works with both CAPI and CNG private keys on .NET Core / .NET 5+
# ---------------------------------------------------------------------------
$contentInfo = [System.Security.Cryptography.Pkcs.ContentInfo]::new($innerBytes)
$signedCms   = [System.Security.Cryptography.Pkcs.SignedCms]::new($contentInfo, $true)

$cmsSigner                 = [System.Security.Cryptography.Pkcs.CmsSigner]::new($signingCert)
$cmsSigner.DigestAlgorithm = [System.Security.Cryptography.Oid]::new("2.16.840.1.101.3.4.2.1") # SHA-256
$cmsSigner.IncludeOption   = [System.Security.Cryptography.X509Certificates.X509IncludeOption]::EndCertOnly

$signedCms.ComputeSignature($cmsSigner)
$signatureBytes = $signedCms.Encode()

Write-Verbose "CMS detached signature computed, $($signatureBytes.Length) bytes"

# ---------------------------------------------------------------------------
# 5. Assemble multipart/signed
#    RFC 5751: boundary, protocol="application/pkcs7-signature", micalg="sha-256"
# ---------------------------------------------------------------------------
$boundary = "----=_Part_$(([System.Guid]::NewGuid().ToString('N')))"

$sigStream = [System.IO.MemoryStream]::new($signatureBytes)
$sigPart   = [MimeKit.MimePart]::new("application", "pkcs7-signature")
$sigPart.Content                 = [MimeKit.MimeContent]::new($sigStream)
$sigPart.ContentDisposition      = [MimeKit.ContentDisposition]::new(
                                       [MimeKit.ContentDisposition]::Attachment)
$sigPart.ContentTransferEncoding = [MimeKit.ContentEncoding]::Base64
$sigPart.FileName                = "smime.p7s"

$multipartSigned = [MimeKit.Multipart]::new("signed")
$multipartSigned.ContentType.Parameters.Add("protocol", "application/pkcs7-signature")
$multipartSigned.ContentType.Parameters.Add("micalg",   "sha-256")
$multipartSigned.ContentType.Boundary = $boundary
$multipartSigned.Add($innerBody)
$multipartSigned.Add($sigPart)

# ---------------------------------------------------------------------------
# 6. Build the MimeMessage and save to .eml
# ---------------------------------------------------------------------------
$message         = [MimeKit.MimeMessage]::new()
$message.From.Add([MimeKit.MailboxAddress]::new($FromEmail, $FromEmail))
$message.To.Add([MimeKit.MailboxAddress]::new($ToEmail, $ToEmail))
$message.Subject = $Subject
$message.Body    = $multipartSigned

$emlPath = Join-Path $OutputDirectory "clearsigned.eml"
$message.WriteTo($emlPath)
Write-Host "EML saved to: $emlPath"

if ($SaveOnly) {
    Write-Host "SaveOnly specified — skipping Graph send."
    return
}

# ---------------------------------------------------------------------------
# 7. Connect to Microsoft Graph and send via raw MIME
#
#    Graph requires the EML bytes to be base64-encoded and POSTed to:
#      POST /v1.0/users/{from}/sendMail
#    with Content-Type: text/plain
#    per https://learn.microsoft.com/en-us/graph/outlook-send-mime-message
# ---------------------------------------------------------------------------
if (-not $GraphClientId -or -not $GraphTenantId) {
    throw "GraphClientId and GraphTenantId are required to send. Use -SaveOnly to skip sending."
}

# Check Microsoft.Graph module is available
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users.Actions)) {
    throw "Microsoft.Graph module not found. Run: Install-Module Microsoft.Graph"
}

Write-Verbose "Connecting to Microsoft Graph (tenant: $GraphTenantId)"

if ($GraphClientSecret) {
    # App-only: client credentials flow
    $secureSecret = ConvertTo-SecureString $GraphClientSecret -AsPlainText -Force
    $credential   = [System.Management.Automation.PSCredential]::new($GraphClientId, $secureSecret)
    Connect-MgGraph -TenantId $GraphTenantId -ClientSecretCredential $credential -NoWelcome
    Write-Verbose "Connected to Graph using app-only (client credentials)"
} else {
    # Delegated: interactive browser sign-in
    Connect-MgGraph -TenantId $GraphTenantId -ClientId $GraphClientId `
                    -Scopes "Mail.Send" -NoWelcome
    Write-Verbose "Connected to Graph using delegated (interactive) auth"
}

# Read the saved EML bytes and base64-encode them
$emlBytes  = [System.IO.File]::ReadAllBytes($emlPath)
$emlBase64 = [System.Convert]::ToBase64String($emlBytes)

# POST to /v1.0/users/{from}/sendMail with Content-Type: text/plain
# The body must be the raw base64 string — Graph decodes it server-side
$uri = "https://graph.microsoft.com/v1.0/users/$FromEmail/sendMail"

Invoke-MgGraphRequest `
    -Method      POST `
    -Uri         $uri `
    -Body        $emlBase64 `
    -ContentType "text/plain"

Write-Host "Message sent via Microsoft Graph."
Write-Host ""
Write-Host "Summary:"
Write-Host "  From       : $FromEmail"
Write-Host "  To         : $ToEmail"
Write-Host "  Subject    : $Subject"
Write-Host "  Signing    : SHA-256 / multipart/signed (clear-signed)"
Write-Host "  Cert       : $($signingCert.Subject)"
Write-Host "  Thumbprint : $($signingCert.Thumbprint)"
Write-Host "  EML saved  : $emlPath"
