# SimpleMailboxClient-ImportExport

A WinForms mailbox browser built on the Microsoft Graph **Mailbox Import/Export API**
(`/beta/admin/exchange/mailboxes`) rather than the normal Graph Mail API. It browses folders,
lists items, shows message detail from MAPI properties, and moves items around as Fast Transfer
Stream (FTS) blobs — export to disk, import from disk, or copy folder-to-folder across mailboxes.

Script:
<https://github.com/gscales/Powershell-Scripts/blob/master/Graph101/GraphSDK/SimpleMailboxClient-ImportExport.ps1>

## Why this API instead of the Mail API

The import/export endpoints exist for high-fidelity mailbox migration, not for reading mail, and
that shapes what the client can and can't do:

- **Mailbox ids aren't addresses.** Every call needs `MBX:{ExchangeMailboxGuid}@{TenantId}`,
  resolved once per mailbox from `/beta/users/{upn}/settings/exchange` → `primaryMailboxId`.
- **Items are thin.** A `mailboxItem` carries only id, type, size, timestamps, changeKey and
  categories. Subject, sender, recipients, body — all of it comes back as tagged MAPI
  `singleValueExtendedProperties`.
- **No `$search`.** There's no KQL/`-Search` equivalent and no documented `$filter` on extended
  properties, so the search box in this client filters client-side over the page of items already
  retrieved.
- **No attachment navigation property.** You can read the body (PR_HTML / PR_BODY), but individual
  attachment bytes aren't reachable. Export the item to get everything.
- **Export output is opaque.** `exportItems` returns a full-fidelity stream meant to be re-imported
  with `createImportSession`/`import`. It is not an `.eml` and shouldn't be parsed as one — this
  client saves it with a `.fts` extension as a reminder.

## Prerequisites

- **PowerShell 5.1 or 7** on Windows. The client is WinForms, so it's Windows-only.
- Run in an **STA** host. `powershell.exe` is STA by default; under `pwsh` use `pwsh -sta`. The
  message viewer uses a `WebBrowser` control to render HTML bodies and needs STA — without it the
  viewer silently falls back to the plain-text body.
- **Microsoft.Graph.Authentication** module. That's the only module needed; every call is a raw
  REST call via `Invoke-MgGraphRequest`, except Send Message which uses `Send-MgUserMail`.

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

## Permissions

Graph permissions are the whole story. What
you do need to decide is **which permission mode** you're running in, because that decides which
mailboxes the client can open.

| | Delegated | Application (app-only) |
|---|---|---|
| Import/export permission | `MailboxItem.ImportExport` | `MailboxItem.ImportExport.All` |
| Resolving `primaryMailboxId` | `MailboxSettings.Read` | `MailboxSettings.Read` |
| New Message / Send Message | `Mail.Send` | `Mail.Send` |
| Which mailboxes | **Only the signed-in account's own** primary and archive | Any mailbox in the tenant |
| Admin consent | Required | Required |

### Delegated

Straightforward to set up, and fine when you're working with your own mailbox. The catch is the
one that trips people coming from EWS: **these are not shared-mailbox permissions.** Delegated
access here only reaches the primary and archive mailboxes of the account you signed in as. Full
access to another mailbox, or being a delegate on it, does not extend the delegated token to it —
`Get-SMCMailboxId` for someone else's UPN will fail regardless of your Exchange rights.

That means with delegated permissions, **Add Mailbox** is only useful for adding your own archive
alongside your primary, and **Copy Folder to...** is limited to folders within your own mailbox.

### Application (app-only)

This is the mode for anything cross-mailbox: opening several users' mailboxes in one tree,
migrating, or copying items from one person's folder into another's. `MailboxItem.ImportExport.All`
grants tenant-wide access, so treat the credential accordingly — and if you want to narrow that
blast radius, scope the app with an [application access
policy](https://learn.microsoft.com/graph/auth-limit-mailbox-access) so it can only touch a defined
set of mailboxes.

## Application registration

`MailboxItem.ImportExport` isn't on the default Graph PowerShell app's consent list, so either way
you'll want your own app registration.

**For delegated:**

1. Azure portal → **App registrations** → **New registration**.
2. Under **Authentication**, add the redirect URI `http://localhost` and enable the app as a
   **public client**.
3. Under **API permissions** → **Delegated permissions**, add `MailboxItem.ImportExport`,
   `MailboxSettings.Read`, and `Mail.Send` if you want the send feature. Grant admin consent.
4. Note the **Application (client) ID** and your **Directory (tenant) ID**.

**For application:**

1. Register the app as above, but skip the public client / redirect URI — app-only doesn't need it.
2. Under **API permissions** → **Application permissions**, add `MailboxItem.ImportExport.All`,
   `MailboxSettings.Read`, and `Mail.Send` if needed. Grant admin consent.
3. Under **Certificates & secrets**, create a client secret, or upload a certificate if you'd
   rather not handle a secret.

## Connecting

Sign in *before* loading or running the client — it assumes a live Graph connection.

**Delegated:**

```powershell
Connect-MgGraph -ClientId "<your-app-id>" -TenantId "<your-tenant-id>" `
    -Scopes "MailboxSettings.Read","MailboxItem.ImportExport","Mail.Send"
```

**Application, with a client secret:**

```powershell
$secret = ConvertTo-SecureString "<your-client-secret>" -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential("<your-app-id>", $secret)
Connect-MgGraph -TenantId "<your-tenant-id>" -ClientSecretCredential $credential
```

**Application, with a certificate:**

```powershell
Connect-MgGraph -ClientId "<your-app-id>" -TenantId "<your-tenant-id>" `
    -CertificateThumbprint "<thumbprint>"
```

App-only takes no `-Scopes`; the token carries whatever application permissions were consented.

Confirm what you actually got — a permission that didn't come through is the fastest explanation
for a 403 later:

```powershell
Get-MgContext | Select-Object AuthType, Scopes
```

## Loading and starting the client

```powershell
Import-Module ./SimpleMailboxClient-ImportExport.ps1
Start-SMCMailClient -MailboxName user@yourtenant.onmicrosoft.com
```

`Import-Module` on a `.ps1` works and keeps the functions in their own module scope; dot-sourcing
(`. ./SimpleMailboxClient-ImportExport.ps1`) works equally well and is handier if you want to call
the individual functions from the console afterwards.

`Start-SMCMailClient` opens the window and loads that mailbox's folder tree straight away. The
window is modal to your session — the prompt comes back when you close it.

## Using the window

**Row 1 — mailbox**

- **Open Mailbox** clears the tree and loads the UPN in the text box.
- **Add Mailbox** prompts for another UPN and appends it as a second root, leaving what's already
  there alone. This is what makes cross-mailbox copies possible: every folder node remembers which
  mailbox it came from, so operations always target the right one regardless of what else is open.
  Note this needs application permissions — under delegated you can only open your own mailbox.
- **# of Items** caps how many items the list view fetches per folder (default 100). It affects the
  list only — folder export and folder copy always process every item.

Selecting a folder in the tree lists its items.

**Row 2 — item actions**

- **Show Message** opens the viewer for the selected row (double-clicking a row does the same).
  It pulls a wide MAPI property set and adapts to the item's message class: mail shows
  From/To/Cc/Sent/Received, a contact shows name, title, company and phone numbers, an appointment
  shows organiser and start/end. The body renders PR_HTML when the item has one, with PR_BODY as
  the fallback; **Show plain text** toggles between them, and **Properties...** dumps every
  property that came back — useful when a field is blank and you want to know whether the property
  is missing or just not displayed.
- **Show Header** shows PR_TRANSPORT_MESSAGE_HEADERS. Items that never arrived over SMTP — calendar
  items, contacts, tasks, anything created directly in the mailbox — legitimately don't have any.
- **New Message** / Send goes through the regular Mail API, not this one. The import/export API has
  no concept of sending.
- **Export Item (.fts)** saves the selected item as a single FTS blob.
- **Export Folder (.fts)...** exports every item in the selected folder to a directory, one file
  per item, named from the subject plus a short hash of the item id so identical or blank subjects
  can't collide. Runs in batches of 20 (the `exportItems` per-request maximum) with a progress bar;
  a failed batch is reported and skipped rather than aborting the rest.
- **Import from FTS...** imports selected `.fts` files into the folder currently selected in the
  tree.
- **Import Folder (.fts)...** imports every `.fts` file found directly in a chosen directory (not
  recursive).
- **Update** re-lists the current folder.

**Row 3 — search**

Tick the checkbox to enable it, pick Subject, From or Body, and type a term. This filters
client-side over the items already retrieved, so it only sees as many items as **# of Items**
fetched. Body search is the slow one — it fetches PR_BODY per candidate item, so keep the item
count small when using it.

**Row 4 — Copy Folder to...**

Copies every item from the selected folder into any other folder in any open mailbox. Pick the
source in the tree, click the button, choose the destination from the picker (double-click a folder
to select it), and confirm.

It works in chunks of 20: export a chunk, import that chunk, discard it, move to the next — so peak
memory is one chunk of blobs no matter how large the folder is, and nothing touches the filesystem.
Two things worth knowing before you run it:

- It's a **copy, not a move**. The destination gets new items with new ids; nothing is deleted from
  the source, and running it twice duplicates everything.
- It copies **items in the folder, not subfolders**. Nested folders need their own pass.

If the destination import session expires mid-copy, a failed item is retried once against a fresh
session before being counted as failed.

## Scripting it without the GUI

The functions are usable directly once the script is loaded. For example, importing exported items
into a mailbox's Calendar regardless of what's selected in the tree:

```powershell
$mailboxId = Get-SMCMailboxId -Upn user@yourtenant.onmicrosoft.com
Import-SMCItemsToCalendar -MailboxId $mailboxId -FilePaths (Get-ChildItem C:\exports\*.fts).FullName
```

Or a folder-to-folder copy with no window at all:

```powershell
$srcId  = Get-SMCMailboxId -Upn source@yourtenant.onmicrosoft.com
$dstId  = Get-SMCMailboxId -Upn dest@yourtenant.onmicrosoft.com
$srcBox = Invoke-SMCGetMailboxFolder -MailboxId $srcId -FolderId "Inbox"
$dstBox = Invoke-SMCGetMailboxFolder -MailboxId $dstId -FolderId "Inbox"

Invoke-SMCCopyFolderItems -SourceMailboxId $srcId -SourceFolderId $srcBox.id `
    -DestinationMailboxId $dstId -DestinationFolderId $dstBox.id
```

Most functions take `-Verbose`, which is where retry, decode and property-group failures get
reported.

## Troubleshooting

**"Could not resolve a primaryMailboxId"** — either `MailboxSettings.Read` wasn't consented, or
you're connected with delegated permissions and asked for a mailbox that isn't your own. Check
`Get-MgContext | Select-Object AuthType, Scopes` first.

**403 on folder or item calls** — usually the same delegated-scope limit: the token only reaches
the signed-in account's own primary and archive mailboxes. Reconnect app-only for anything else.
If you're already app-only, check that admin consent was actually granted for
`MailboxItem.ImportExport.All`, and that an application access policy isn't excluding the mailbox.

**"exportItems accepts a maximum of 20 item ids per request"** — the batching is meant to prevent
this. If it appears, the item list came back nested one level deeper than expected;
`Get-SMCFlatItemList` normalises that, so make sure you're running a build that includes it.

**Message body is empty** — open **Properties...** in the viewer. If neither `Binary 0x1013` nor
`String 0x1000` is listed, the store didn't return a body for that item; if PR_HTML is there but
nothing renders, you're likely in a non-STA host and looking at the text fallback.

**Body renders with mojibake** — under PowerShell 7, single-byte code pages like 1252 aren't
registered by default on .NET, so the HTML decode falls back to UTF-8. Register the provider first:

```powershell
[System.Text.Encoding]::RegisterProvider([System.Text.CodePagesEncodingProvider]::Instance)
```

**Throttling (HTTP 429)** — imports honour the server's `Retry-After` and retry up to three times.
A large copy against a busy tenant will simply take a while.