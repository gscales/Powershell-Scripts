<#
.SYNOPSIS
    Sends a message and files the sent copy into \Inbox\SentByApp, emulating the
    EWS SendItem SavedItemFolderId behaviour that Microsoft Graph has no direct
    equivalent for.

.DESCRIPTION
    Graph cannot send a message into a nominated folder. Neither sendMail nor
    messages/{id}/send accepts a destination, so the only route is:

        create draft -> send -> wait for it to land in Sent Items -> move it

    The wait is the awkward part. Send is asynchronous: the item leaves Drafts
    and appears in Sent Items some time later, and a move issued too early
    either fails or - the failure mode people actually hit - leaves an empty
    draft sitting in the destination folder while the real sent message is
    nowhere to be found.

    Two things make this reliable enough to use:

      1. Immutable IDs. By default a message id changes when the item moves
         between folders, so the draft id is worthless after the send. Sending
         Prefer: IdType="ImmutableId" on EVERY call keeps one id valid across
         the whole sequence. Mixing immutable and default ids across calls is
         the single most common way to break this.

      2. Polling on ParentFolderId rather than on a fixed sleep. The script
         waits until the message reports itself in Sent Items before moving it.

.NOTES
    Delegated permissions:   Mail.ReadWrite, Mail.Send
    Application permissions: Mail.ReadWrite, Mail.Send, scoped with
                             New-ApplicationAccessPolicy

    Even with all this, the move is a second operation that can fail after the
    mail has already gone out. Treat "sent" and "filed" as separate outcomes -
    the recipient has the message either way.
#>

#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Mail, Microsoft.Graph.Users.Actions

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $Mailbox,

    [Parameter(Mandatory)]
    [string] $To,

    [string] $Subject = 'Filed by Send-AndFileMessage',

    [string] $Body = 'This message was sent and then filed into \Inbox\SentByApp.',

    [string] $FolderName = 'SentByApp',

    [string] $ParentFolder = 'inbox',

    [int] $TimeoutSeconds = 60
)

$ErrorActionPreference = 'Stop'

# Applied to every request. Consistency matters more than the header itself.
$ImmutableId = @{ Prefer = 'IdType="ImmutableId"' }

function Resolve-TargetFolder {
    param([string] $UserId, [string] $Parent, [string] $Name)

    $existing = Get-MgUserMailFolderChildFolder -UserId $UserId `
        -MailFolderId $Parent -Filter "displayName eq '$Name'" `
        -Headers $ImmutableId -ErrorAction SilentlyContinue

    if ($existing) {
        Write-Verbose "Found existing folder '$Name'."
        return $existing[0]
    }

    Write-Verbose "Creating folder '$Name' under '$Parent'."
    return New-MgUserMailFolderChildFolder -UserId $UserId -MailFolderId $Parent `
        -DisplayName $Name -Headers $ImmutableId
}

function Wait-ForSentItem {
    <#
        Polls until the message reports itself outside Drafts. Returns the
        message, or throws if it never arrives.
    #>
    param([string] $UserId, [string] $MessageId, [string] $SentItemsId, [int] $Timeout)

    $deadline = (Get-Date).AddSeconds($Timeout)
    $delay = 500

    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds $delay

        # Briefly unresolvable mid-send; that is expected, not an error.
        $message = Get-MgUserMessage -UserId $UserId -MessageId $MessageId `
            -Property 'id,parentFolderId,isDraft,sentDateTime' `
            -Headers $ImmutableId -ErrorAction SilentlyContinue

        if ($message -and -not $message.IsDraft -and $message.ParentFolderId -eq $SentItemsId) {
            Write-Verbose "Message reached Sent Items after $([int]((Get-Date) - $start).TotalMilliseconds) ms."
            return $message
        }

        # Back off gently rather than hammering the mailbox.
        $delay = [Math]::Min($delay * 2, 4000)
    }

    throw "The message did not appear in Sent Items within $Timeout seconds. " +
          "It has almost certainly been sent - check the mailbox before resending."
}

# --- 1. Connect -------------------------------------------------------------
# Delegated. For app-only use:
#   Connect-MgGraph -ClientId $id -TenantId $tid -CertificateThumbprint $thumb
if (-not (Get-MgContext)) {
    Connect-MgGraph -Scopes 'Mail.ReadWrite', 'Mail.Send' | Out-Null
}

$start = Get-Date

# --- 2. Resolve \Inbox\SentByApp and Sent Items -----------------------------
$target = Resolve-TargetFolder -UserId $Mailbox -Parent $ParentFolder -Name $FolderName
Write-Verbose "Target folder id: $($target.Id)"

$sentItems = Get-MgUserMailFolder -UserId $Mailbox -MailFolderId 'sentitems' `
    -Headers $ImmutableId

# --- 3. Create the draft ----------------------------------------------------
$draftParams = @{
    Subject      = $Subject
    Body         = @{ ContentType = 'Text'; Content = $Body }
    ToRecipients = @(
        @{ EmailAddress = @{ Address = $To } }
    )
}

$draft = New-MgUserMessage -UserId $Mailbox -BodyParameter $draftParams `
    -Headers $ImmutableId

Write-Verbose "Draft created: $($draft.Id)"

# --- 4. Send it -------------------------------------------------------------
# Note there is no saveToSentItems switch on this endpoint - a draft send
# always files a copy. That is precisely why the move below is needed.
Send-MgUserMessage -UserId $Mailbox -MessageId $draft.Id -Headers $ImmutableId

Write-Verbose 'Send accepted; waiting for the item to land in Sent Items.'

# --- 5. Wait, then move -----------------------------------------------------
$sent = Wait-ForSentItem -UserId $Mailbox -MessageId $draft.Id `
    -SentItemsId $sentItems.Id -Timeout $TimeoutSeconds

$moved = Move-MgUserMessage -UserId $Mailbox -MessageId $sent.Id `
    -DestinationId $target.Id -Headers $ImmutableId

[pscustomobject]@{
    MessageId     = $moved.Id
    ParentFolder  = "\$ParentFolder\$FolderName"
    FolderId      = $target.Id
    Sent          = $true
    Filed         = $moved.ParentFolderId -eq $target.Id
    ElapsedMs     = [int]((Get-Date) - $start).TotalMilliseconds
}