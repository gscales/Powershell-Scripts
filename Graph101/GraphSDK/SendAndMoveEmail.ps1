<#
.SYNOPSIS
    Sends a message and files the sent copy into a folder nominated by path,
    emulating the EWS SendItem SavedItemFolderId behaviour that Microsoft Graph
    has no direct equivalent for.

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

.PARAMETER FolderPath
    Backslash-delimited path from the mailbox root, leading backslash included:
    '\Inbox\SentByApp'. '\' resolves to MsgFolderRoot itself.

.EXAMPLE
    .\Send-AndFileMessage.ps1 -Mailbox gscales@datarumble.com `
        -To recipient@example.com -FolderPath '\Inbox\SentByApp' -CreateMissing -Verbose

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

    [string] $Body = 'This message was sent and then filed by path.',

    [string] $FolderPath = '\Inbox\SentByApp',

    # Create any level of the path that does not exist. Off by default so a
    # typo'd path fails instead of quietly creating a new folder tree.
    [switch] $CreateMissing,

    [int] $TimeoutSeconds = 60
)

$ErrorActionPreference = 'Stop'

# Applied to every request. Consistency matters more than the header itself.
$ImmutableId = @{ Prefer = 'IdType="ImmutableId"' }

function Get-MailBoxFolderFromPath {
    <#
        Walks a backslash-delimited folder path from the mailbox root, resolving
        one level at a time by display name.

        Two additions over the original: -Headers, so the caller can keep the
        immutable-id Prefer header consistent across every call in a sequence,
        and -CreateMissing, so a path can be provisioned rather than only found.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string]
        $FolderPath,

        [Parameter(Position = 1, Mandatory = $true)]
        [String]
        $MailboxName,

        [Parameter(Position = 2, Mandatory = $false)]
        [String]
        $WellKnownSearchRoot = "MsgFolderRoot",

        [Parameter(Mandatory = $false)]
        [hashtable]
        $Headers = @{},

        [Parameter(Mandatory = $false)]
        [switch]
        $CreateMissing
    )
    process {
        if ($FolderPath -eq '\') {
            return Get-MgUserMailFolder -UserId $MailboxName -MailFolderId msgFolderRoot -Headers $Headers
        }

        $fldArray = $FolderPath.Split("\")

        #Loop through the Split Array and do a Search for each level of folder
        $folderId = $WellKnownSearchRoot

        for ($lint = 1; $lint -lt $fldArray.Length; $lint++) {
            #Perform search based on the displayname of each folder level
            $FolderName = $fldArray[$lint]

            if ([string]::IsNullOrWhiteSpace($FolderName)) {
                throw ("Folder path '$FolderPath' contains an empty level")
            }

            # A folder named O'Brien would otherwise terminate the OData string
            # literal and produce a filter syntax error rather than a miss.
            $filterName = $FolderName.Replace("'", "''")

            $tfTargetFolder = Get-MgUserMailFolderChildFolder -UserId $MailboxName `
                -Filter "DisplayName eq '$filterName'" -MailFolderId $folderId -All -Headers $Headers

            # Display names are not unique among siblings. Taking .Id off an
            # array yields an array, which fails further down in a way that does
            # not name the cause - so pick one deliberately and say so.
            if (@($tfTargetFolder).Count -gt 1) {
                Write-Warning ("Folder level '$FolderName' matched " +
                    "$(@($tfTargetFolder).Count) sibling folders; using the first.")
                $tfTargetFolder = @($tfTargetFolder)[0]
            }

            if ($tfTargetFolder -and $tfTargetFolder.DisplayName -eq $FolderName) {
                $folderId = $tfTargetFolder.Id.ToString()
            }
            elseif ($CreateMissing) {
                Write-Verbose "Creating missing folder level '$FolderName'."
                $tfTargetFolder = New-MgUserMailFolderChildFolder -UserId $MailboxName `
                    -MailFolderId $folderId -DisplayName $FolderName -Headers $Headers
                $folderId = $tfTargetFolder.Id.ToString()
            }
            else {
                throw ("Folder Not found")
            }
        }

        return $tfTargetFolder
    }
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
            return $message
        }

        # Back off gently rather than hammering the mailbox.
        $delay = [Math]::Min($delay * 2, 4000)
    }

    throw ("The message did not appear in Sent Items within $Timeout seconds. " +
           "It has almost certainly been sent - check the mailbox before resending.")
}

# --- 1. Connect -------------------------------------------------------------
# Delegated. For app-only use:
#   Connect-MgGraph -ClientId $id -TenantId $tid -CertificateThumbprint $thumb
if (-not (Get-MgContext)) {
    Connect-MgGraph -Scopes 'Mail.ReadWrite', 'Mail.Send' | Out-Null
}

$start = Get-Date

# --- 2. Resolve the target folder by path, and Sent Items -------------------
$target = Get-MailBoxFolderFromPath -MailboxName $Mailbox -FolderPath $FolderPath `
    -Headers $ImmutableId -CreateMissing:$CreateMissing

Write-Verbose "Resolved '$FolderPath' to folder id $($target.Id)."

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
    MessageId    = $moved.Id
    FolderPath   = $FolderPath
    FolderId     = $target.Id
    Sent         = $true
    Filed        = $moved.ParentFolderId -eq $target.Id
    ElapsedMs    = [int]((Get-Date) - $start).TotalMilliseconds
}
