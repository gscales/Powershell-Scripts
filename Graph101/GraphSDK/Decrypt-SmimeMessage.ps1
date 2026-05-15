<#
.SYNOPSIS
    Fetches the latest S/MIME message from a mailbox via Graph and decrypts it.

.DESCRIPTION
    Connects to Microsoft Graph and finds the latest S/MIME message in the inbox,
    downloads the smime.p7m attachment, locates the matching decryption certificate
    in CurrentUser\My, and decrypts it using EnvelopedCms.

.PARAMETER MailboxEmail
    The mailbox to search (e.g. user@contoso.com).

.PARAMETER OutputDirectory
    Directory to save the decrypted content. Defaults to .\output.

.EXAMPLE
    .\Decrypt-SmimeMessage.ps1 -MailboxEmail "user@contoso.com"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $MailboxEmail,
    [string] $OutputDirectory = "c:\temp\"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# 1. Connect to Microsoft Graph (skipped if already connected)
# ---------------------------------------------------------------------------
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users.Actions)) {
    throw "Microsoft.Graph module not found. Run: Install-Module Microsoft.Graph"
}

try {
    $mgContext = Get-MgContext
} catch {
    $mgContext = $null
}

if (-not $mgContext) {
    Connect-MgGraph -Scopes "Mail.Read" -NoWelcome
    Write-Verbose "Connected to Microsoft Graph"
} else {
    Write-Verbose "Using existing Microsoft Graph connection ($($mgContext.Account))"
}

# ---------------------------------------------------------------------------
# 2. Find the latest S/MIME message in the inbox using a server-side filter
#    on PidTagMessageClass (0x001A) via singleValueExtendedProperties.
# ---------------------------------------------------------------------------
Write-Host "Searching for latest S/MIME message in inbox..."

$filterSmime = "singleValueExtendedProperties/Any(ep: ep/id eq 'String 0x001A' and ep/value eq 'IPM.Note.SMIME')"

$response     = Get-MgUserMessage -UserId $MailboxEmail -Filter $filterSmime -Select "id,subject,receivedDateTime,hasAttachments" -ExpandProperty "singleValueExtendedProperties(`$filter=id eq 'String 0x001A')"
$smimeMessage = $response | Sort-Object receivedDateTime -Descending | Select-Object -First 1

if (-not $smimeMessage) {
    throw "No S/MIME message found in the inbox."
}

$msgClass = ($smimeMessage.SingleValueExtendedProperties |
             Where-Object { $_.Id -eq "String 0x1A" }).Value

Write-Host "Found message:"
Write-Host "  Subject  : $($smimeMessage.subject)"
Write-Host "  Received : $($smimeMessage.receivedDateTime)"
Write-Host "  Class    : $msgClass"

# ---------------------------------------------------------------------------
# 3. Download the smime.p7m attachment
# ---------------------------------------------------------------------------
# List attachments to find the smime.p7m attachment ID
$attachList = Get-MgUserMessageAttachment -UserId $MailboxEmail -MessageId $smimeMessage.Id -Property "id,name"
$p7mMeta    = $attachList | Where-Object { $_.Name -eq "smime.p7m" } | Select-Object -First 1

if (-not $p7mMeta) {
    throw "No smime.p7m attachment found on message '$($smimeMessage.subject)'."
}

# Fetch the full attachment by ID to get contentBytes
$p7mAttach = Get-MgUserMessageAttachment -UserId $MailboxEmail -MessageId $smimeMessage.Id -AttachmentId $p7mMeta.Id
$p7mBytes  = [System.Convert]::FromBase64String($p7mAttach.AdditionalProperties["contentBytes"])
Write-Host "  Attachment: smime.p7m ($($p7mBytes.Length) bytes)"


# ---------------------------------------------------------------------------
# 4. Decrypt EnvelopedData using System.Security.Cryptography.Pkcs.EnvelopedCms
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "Decrypting EnvelopedData..."

$envelopedCms = [System.Security.Cryptography.Pkcs.EnvelopedCms]::new()
$envelopedCms.Decode($p7mBytes)

$certStore = [System.Security.Cryptography.X509Certificates.X509Store]::new(
    [System.Security.Cryptography.X509Certificates.StoreName]::My,
    [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser
)
$certStore.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)

try {
    $envelopedCms.Decrypt($certStore.Certificates)
} catch {
    throw "Decryption failed. Ensure the private key for the recipient certificate is in CurrentUser\My. Error: $_"
} finally {
    $certStore.Close()
}

$decryptedBytes = $envelopedCms.ContentInfo.Content
Write-Host "Decryption succeeded. Decrypted content: $($decryptedBytes.Length) bytes"

# ---------------------------------------------------------------------------
# 5. Save and display the decrypted content
# ---------------------------------------------------------------------------
$null = [System.IO.Directory]::CreateDirectory($OutputDirectory)

$decryptedPath = Join-Path $OutputDirectory "smime-decrypted.eml"
[System.IO.File]::WriteAllBytes($decryptedPath, $decryptedBytes)
Write-Host "Decrypted bytes saved to: $decryptedPath"

try {
    $text = [System.Text.Encoding]::UTF8.GetString($decryptedBytes)
    Write-Host ""
    Write-Host "--- Decrypted content (UTF-8) ---"
    Write-Host $text
} catch {
    Write-Host "Content is not UTF-8 text - inspect the saved .bin file."
}
 