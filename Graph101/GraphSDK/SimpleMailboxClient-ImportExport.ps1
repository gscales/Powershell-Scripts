<#
.SYNOPSIS
    Simple Mailbox Client - converted to use the Microsoft Graph Mailbox Import/Export API
    (admin/exchange/mailboxes) instead of the regular Get-MgUserMailFolder* / Get-MgUserMailFolderMessage
    cmdlets for listing folders and items.

.NOTES
    WHY THIS LOOKS DIFFERENT FROM THE ORIGINAL

    The admin/exchange/mailboxes endpoints are a different API surface (currently beta) built for
    high-fidelity mailbox migration, not for reading mail. That has some real consequences here:

    1. MailboxId is NOT the UPN/SMTP address. It's an opaque identifier in the form
       "MBX:{ExchangeMailboxGuid}@{TenantId}". You resolve it once per mailbox via:
           GET https://graph.microsoft.com/beta/users/{upn}/settings/exchange
       and use the returned .primaryMailboxId property. See Get-SMCMailboxId below.
       Reference: https://learn.microsoft.com/graph/api/resources/mailbox-import-export-api-overview

    2. mailboxFolder objects DO natively include displayName, childFolderCount, parentFolderId,
       totalItemCount and type - no extended properties needed for folders.

    3. mailboxItem objects are much thinner than a Graph "message". The only native properties are:
           id, type, size, createdDateTime, lastModifiedDateTime, changeKey, categories
       There is no subject, sender, receivedDateTime, hasAttachments, or body on the base object.
       All of those have to be pulled back as MAPI singleValueExtendedProperties. The list view
       pulls the handful it needs for its columns; the message viewer pulls a much wider set - see
       Get-SMCItemViewerDetail.

    4. There is no $search support on this API (unlike the regular mail API's -Search / KQL support),
       and no documented support for filtering on extended properties via $filter. The "Search by
       Property" panel in this version filters client-side against the extended properties already
       retrieved for the current page of items, after they've been expanded. $filter is still used
       for the handful of native fields (createdDateTime, type) where that's documented to work.

    5. There is no per-item Attachments navigation property, so attachment BYTES still aren't
       reachable item-by-item. The body, however, is: PR_HTML (0x1013) comes back as a Binary
       extended property, and the viewer decodes it with PR_INTERNET_CPID (0x3FDE) and renders it
       in a WebBrowser control, falling back to PR_BODY (0x1000) plain text when an item has no
       HTML. Full fidelity including attachments still means exporting the item with exportItems,
       which returns an opaque stream (.fts) for re-import - not something to parse as an .eml.

    6. Sending mail (New Message / Send Message) is untouched and still goes through the regular Graph
       Mail API (Send-MgUserMail), because the import/export API has no concept of sending - it's
       import/export only.

    7. "Copy Folder to..." copies every item from the selected folder into any other folder in the
       tree - same mailbox or a different one - by round-tripping FTS blobs through memory in
       20-item chunks. See Invoke-SMCCopyFolderItems.

    STA REQUIREMENT
    The message viewer uses System.Windows.Forms.WebBrowser to render HTML bodies, and that control
    requires a single-threaded-apartment host. Run this from powershell.exe (which is STA by
    default) or `powershell -sta`; under a host that isn't STA the viewer falls back to showing the
    plain-text body instead of failing.

    PERMISSIONS
    Delegated scopes needed:
        MailboxSettings.Read                (to resolve primaryMailboxId)
        MailboxItem.ImportExport            (to list folders/items, export, and import)
        Mail.Send                           (only if you use the New/Send Message feature)
    The calling account also needs to be assigned whatever Exchange Online RBAC role currently grants
    mailbox import/export rights on top of Graph consent - check the overview doc for current
    requirements, this has moved around while the API has been in beta:
    https://learn.microsoft.com/graph/mailbox-import-export-concept-overview

    Example connect:
        Connect-MgGraph -Scopes "MailboxSettings.Read","MailboxItem.ImportExport","Mail.Send"

    This script only needs the Microsoft.Graph.Authentication module (Connect-MgGraph /
    Invoke-MgGraphRequest) - it no longer needs Microsoft.Graph.Mail, since every call here is a raw
    Graph REST call against the admin/exchange/mailboxes surface.
#>

[VOID][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[VOID][System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
[VOID][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
$Script:form = new-object System.Windows.Forms.form

#region ---------------------------- Mailbox Import/Export API helpers ----------------------------

function Get-SMCMailboxId {
	<#
	Resolves a UPN/SMTP address to the MBX:{ExchangeGuid}@{TenantId} identifier required by every
	admin/exchange/mailboxes/{mailboxId}/... call. Do this once per mailbox and cache the result -
	don't try to construct the id yourself.
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [String]$Upn
	)
	Process {
		$RequestURL = "https://graph.microsoft.com/beta/users/$Upn/settings/exchange"
		$Result = Invoke-MgGraphRequest -Method Get -Uri $RequestURL
		if ([String]::IsNullOrEmpty($Result.primaryMailboxId)) {
			throw "Could not resolve a primaryMailboxId for $Upn - check the MailboxSettings.Read scope/consent and that the mailbox exists."
		}
		return $Result.primaryMailboxId
	}
}

function Get-SMCTaggedProperty {
	<# Builds a tagged-MAPI-property descriptor, e.g. Id=0x37 DataType=String -> PR_SUBJECT. #>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [String]$DataType,
		[Parameter(Position = 1, Mandatory = $true)] [String]$Id
	)
	Begin {
		$Property = "" | Select-Object Id, DataType
		$Property.Id = $Id
		$Property.DataType = $DataType
		return , $Property
	}
}

function Get-SMCExtendedPropertyFilter {
	<#
	Turns a list of Get-SMCTaggedProperty descriptors into the OData $filter clause used inside
	$expand=singleValueExtendedProperties($filter=...), e.g.:
	    (Id eq 'String 0x37') or (Id eq 'SystemTime 0xE06')
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [Object[]]$PropertyList
	)
	Begin {
		$clauses = foreach ($Prop in $PropertyList) {
			"(Id eq '" + $Prop.DataType + " " + $Prop.Id + "')"
		}
		return ($clauses -join " or ")
	}
}

function Expand-SMCItemProperties {
	<#
	Maps the raw singleValueExtendedProperties collection on a mailboxItem back onto friendly
	note properties (Subject, ReceivedDateTime, SenderName, SenderAddress, DisplayTo, HasAttachments,
	BodyPreview, TransportHeaders) so the rest of the script can just read $item.Subject etc.
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [PSObject]$Item
	)
	Process {
		foreach ($Prop in $Item.singleValueExtendedProperties) {
			switch ($Prop.Id) {
				"String 0x37" { Add-Member -InputObject $Item -NotePropertyName "Subject" -NotePropertyValue $Prop.Value -Force }
				"SystemTime 0xE06" { Add-Member -InputObject $Item -NotePropertyName "ReceivedDateTime" -NotePropertyValue $Prop.Value -Force }
				"String 0xC1A" { Add-Member -InputObject $Item -NotePropertyName "SenderName" -NotePropertyValue $Prop.Value -Force }
				"String 0xC1F" { Add-Member -InputObject $Item -NotePropertyName "SenderAddress" -NotePropertyValue $Prop.Value -Force }
				"String 0xE04" { Add-Member -InputObject $Item -NotePropertyName "DisplayTo" -NotePropertyValue $Prop.Value -Force }
				"Boolean 0xE1B" { Add-Member -InputObject $Item -NotePropertyName "HasAttachments" -NotePropertyValue ([System.Convert]::ToBoolean($Prop.Value)) -Force }
				"String 0x1000" { Add-Member -InputObject $Item -NotePropertyName "BodyPreview" -NotePropertyValue $Prop.Value -Force }
				"String 0x7D" { Add-Member -InputObject $Item -NotePropertyName "TransportHeaders" -NotePropertyValue $Prop.Value -Force }
				"String 0x1A" { Add-Member -InputObject $Item -NotePropertyName "MessageClass" -NotePropertyValue $Prop.Value -Force }
			}
		}
		# Fill in blanks so downstream binding never chokes on a missing property.
		foreach ($name in @("Subject", "ReceivedDateTime", "SenderName", "SenderAddress", "DisplayTo", "MessageClass")) {
			if (-not (Get-Member -InputObject $Item -Name $name -ErrorAction SilentlyContinue)) {
				Add-Member -InputObject $Item -NotePropertyName $name -NotePropertyValue "" -Force
			}
		}
		if (-not (Get-Member -InputObject $Item -Name "HasAttachments" -ErrorAction SilentlyContinue)) {
			Add-Member -InputObject $Item -NotePropertyName "HasAttachments" -NotePropertyValue $false -Force
		}
		return $Item
	}
}

function Invoke-SMCGetMailboxFolder {
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [String]$MailboxId,
		[Parameter(Position = 1, Mandatory = $false)] [String]$FolderId = "MsgFolderRoot"
	)
	Process {
		$RequestURL = "https://graph.microsoft.com/beta/admin/exchange/mailboxes/$MailboxId/folders/$FolderId"
		return Invoke-MgGraphRequest -Method Get -Uri $RequestURL
	}
}

function Invoke-SMCGetChildFolders {
	<# Single level of child mailboxFolders (paged). displayName/childFolderCount/type are native. #>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [String]$MailboxId,
		[Parameter(Position = 1, Mandatory = $true)] [String]$FolderId
	)
	Process {
		$Folders = @()
		$RequestURL = "https://graph.microsoft.com/beta/admin/exchange/mailboxes/$MailboxId/folders/$FolderId/childFolders?`$top=999"
		do {
			$Results = Invoke-MgGraphRequest -Method Get -Uri $RequestURL
			$RequestURL = $null
			if ($Results) {
				$Folders += $Results.Value
				$RequestURL = $Results.'@odata.nextLink'
			}
		} until ([String]::IsNullOrEmpty($RequestURL))
		# Comma operator guards against PowerShell unwrapping a 1-element array on return (see the
		# longer explanation on Invoke-SMCExportItems) - cheap insurance even though today's only
		# consumer (foreach-based tree building) happens to tolerate the unwrapped form fine.
		return , @($Folders)
	}
}

function Invoke-SMCListFolderItems {
	<#
	Lists items in a folder, expanding the extended properties needed for a mail-list view
	(subject/received time/sender/to/hasAttachments). Size, id, type, createdDateTime are native.
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [String]$MailboxId,
		[Parameter(Position = 1, Mandatory = $true)] [String]$FolderId,
		[Parameter(Position = 2, Mandatory = $false)] [Int]$ItemCount = 100
	)
	Process {
		$Props = @()
		$Props += (Get-SMCTaggedProperty -DataType String -Id "0x37")
		$Props += (Get-SMCTaggedProperty -DataType SystemTime -Id "0xE06")
		$Props += (Get-SMCTaggedProperty -DataType String -Id "0xC1A")
		$Props += (Get-SMCTaggedProperty -DataType String -Id "0xC1F")
		$Props += (Get-SMCTaggedProperty -DataType String -Id "0xE04")
		$Props += (Get-SMCTaggedProperty -DataType Boolean -Id "0xE1B")
		$Props += (Get-SMCTaggedProperty -DataType String -Id "0x1A")
		$expandFilter = Get-SMCExtendedPropertyFilter -PropertyList $Props
		$expandClause = "singleValueExtendedProperties(`$filter=$expandFilter)"

		# ItemCount <= 0 means "no limit, fetch every item in the folder" (used by the whole-
		# folder export/copy) - use the max page size for that case rather than clamping to a
		# minimum of 1, which would force one-item-per-request paging and make a large folder crawl.
		$TopVal = if ($ItemCount -le 0) { 999 } else { [Math]::Min(999, $ItemCount) }
		$RequestURL = "https://graph.microsoft.com/beta/admin/exchange/mailboxes/$MailboxId/folders/$FolderId/items?`$top=$TopVal&`$expand=$expandClause"

		$Items = @()
		do {
			$Results = Invoke-MgGraphRequest -Method Get -Uri $RequestURL
			$RequestURL = $null
			if ($Results) {
				foreach ($rawItem in $Results.Value) {
					$Items += (Expand-SMCItemProperties -Item ([PSCustomObject]$rawItem))
					if ($ItemCount -gt 0 -and $Items.Count -ge $ItemCount) {
						# Comma operator guards against PowerShell unwrapping a 1-element array on
						# return - see Invoke-SMCExportItems for the full explanation. Without this,
						# a folder with exactly one item comes back as a bare object instead of a
						# 1-element array, which foreach/.Count tolerate fine but range-indexing
						# ($items[$i..$j], used by the whole-folder export) does not - that's what
						# silently turned a real item into an empty export batch with no error.
						return , @($Items)
					}
				}
				$RequestURL = $Results.'@odata.nextLink'
			}
		} until ([String]::IsNullOrEmpty($RequestURL))
		return , @($Items)
	}
}

function Invoke-SMCGetItemDetail {
	<# Single item, expanding whatever tagged properties the caller asks for. #>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [String]$MailboxId,
		[Parameter(Position = 1, Mandatory = $true)] [String]$FolderId,
		[Parameter(Position = 2, Mandatory = $true)] [String]$ItemId,
		[Parameter(Position = 3, Mandatory = $true)] [Object[]]$PropertyList
	)
	Process {
		$expandFilter = Get-SMCExtendedPropertyFilter -PropertyList $PropertyList
		$expandClause = "singleValueExtendedProperties(`$filter=$expandFilter)"
		$RequestURL = "https://graph.microsoft.com/beta/admin/exchange/mailboxes/$MailboxId/folders/$FolderId/items/$ItemId`?`$expand=$expandClause"
		$rawItem = Invoke-MgGraphRequest -Method Get -Uri $RequestURL
		return (Expand-SMCItemProperties -Item ([PSCustomObject]$rawItem))
	}
}

function Invoke-SMCExportItems {
	<#
	Exports up to 20 mailboxItem ids at a time as opaque, full-fidelity blobs. This is NOT a
	.eml/.msg export - the returned data is meant to be round-tripped back in with
	createImportSession/import, not parsed. Saved here with a .fts extension for that reason.
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [String]$MailboxId,
		[Parameter(Position = 1, Mandatory = $true)] [String[]]$ItemIds
	)
	Process {
		if ($ItemIds.Count -gt 20) {
			throw "exportItems accepts a maximum of 20 item ids per request."
		}
		$RequestURL = "https://graph.microsoft.com/beta/admin/exchange/mailboxes/$MailboxId/exportItems"
		$Body = @{ itemIds = $ItemIds } | ConvertTo-Json -Depth 5
		$Result = Invoke-MgGraphRequest -Method Post -Uri $RequestURL -Body $Body -ContentType "application/json"
		# The comma operator here is deliberate, not decoration: PowerShell unwraps a single-
		# element array when it's returned from a function, so exporting exactly one item would
		# otherwise hand the caller the bare Hashtable itself instead of a 1-element array. Then
		# $exported[0] on that bare Hashtable does a *key* lookup for a key named "0" - which
		# doesn't exist - silently returning $null instead of throwing, which is exactly the "good
		# request, no data" symptom. @() guarantees array shape; the leading , stops it being
		# re-flattened on the way out, so the caller always gets a real, indexable array back.
		return , @($Result.value)
	}
}

function New-SMCProgressForm {
	<#
	A small non-modal progress window with a status label and a bar. Non-modal (.Show, not
	.ShowDialog) so the script can keep running the export/import loop and just poke this window's
	controls directly between steps. PowerShell WinForms scripts are single-threaded - there's no
	background worker here - so Update-SMCProgressForm calls Application.DoEvents() after each
	update to pump the message loop and actually repaint; without that the window would just sit
	there frozen-looking until the whole operation finished.
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)] [String]$Title = "Working..."
	)
	$frm = New-Object System.Windows.Forms.Form
	$frm.Text = $Title
	$frm.Size = New-Object System.Drawing.Size(440, 140)
	$frm.FormBorderStyle = "FixedDialog"
	$frm.StartPosition = "CenterParent"
	$frm.ControlBox = $false
	$frm.MinimizeBox = $false
	$frm.MaximizeBox = $false
	$frm.TopMost = $true

	$lbl = New-Object System.Windows.Forms.Label
	$lbl.Location = New-Object System.Drawing.Point(15, 15)
	$lbl.Size = New-Object System.Drawing.Size(400, 20)
	$lbl.Text = "Starting..."
	[void]$frm.Controls.Add($lbl)

	$bar = New-Object System.Windows.Forms.ProgressBar
	$bar.Location = New-Object System.Drawing.Point(15, 45)
	$bar.Size = New-Object System.Drawing.Size(400, 25)
	$bar.Minimum = 0; $bar.Maximum = 100; $bar.Value = 0
	[void]$frm.Controls.Add($bar)

	Add-Member -InputObject $frm -NotePropertyName StatusLabel -NotePropertyValue $lbl -Force
	Add-Member -InputObject $frm -NotePropertyName Bar -NotePropertyValue $bar -Force

	[void]$frm.Show($Script:form)
	[System.Windows.Forms.Application]::DoEvents()
	return $frm
}

function Update-SMCProgressForm {
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [System.Windows.Forms.Form]$ProgressForm,
		[Parameter(Position = 1, Mandatory = $true)] [String]$Status,
		[Parameter(Position = 2, Mandatory = $false)] [Int]$PercentComplete = -1
	)
	if (-not $ProgressForm -or $ProgressForm.IsDisposed) { return }
	$ProgressForm.StatusLabel.Text = $Status
	if ($PercentComplete -ge 0) {
		$ProgressForm.Bar.Value = [Math]::Max(0, [Math]::Min(100, $PercentComplete))
	}
	[System.Windows.Forms.Application]::DoEvents()
}

function Close-SMCProgressForm {
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)] [System.Windows.Forms.Form]$ProgressForm
	)
	if ($ProgressForm -and -not $ProgressForm.IsDisposed) {
		$ProgressForm.Close()
		$ProgressForm.Dispose()
	}
}

function Get-SMCSafeFileName {
	<#
	Builds a filesystem-safe, unique .fts filename from an item's subject: strips characters
	Windows won't allow in filenames, collapses whitespace, truncates an overly long subject, and
	appends an index plus a short hash of the item's own id - the hash guarantees uniqueness even
	when many items share an identical (or blank) subject, without embedding the raw Graph item id
	(which routinely contains characters like / and + that are unsafe or path-breaking) directly
	in the filename.
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [Int]$Index,
		[Parameter(Position = 1, Mandatory = $false)] [String]$Subject,
		[Parameter(Position = 2, Mandatory = $true)] [String]$ItemId
	)
	Begin {
		if ([String]::IsNullOrWhiteSpace($Subject)) { $Subject = "item" }
		$invalidChars = [Regex]::Escape([String]([System.IO.Path]::GetInvalidFileNameChars() -join ''))
		$safe = [Regex]::Replace($Subject, "[$invalidChars]", "_")
		$safe = ($safe -replace '\s+', ' ').Trim()
		if ($safe.Length -gt 60) { $safe = $safe.Substring(0, 60).Trim() }
		if ([String]::IsNullOrWhiteSpace($safe)) { $safe = "item" }

		$md5 = [System.Security.Cryptography.MD5]::Create()
		try {
			$hashBytes = $md5.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($ItemId))
		}
		finally {
			$md5.Dispose()
		}
		$shortHash = -join ($hashBytes[0..3] | ForEach-Object { $_.ToString("x2") })

		return "{0:D5}_{1}_{2}.fts" -f $Index, $safe, $shortHash
	}
}

function Get-SMCFlatItemList {
	<#
	Normalises whatever Invoke-SMCListFolderItems handed back into a genuinely flat, indexable
	array of item objects.

	Why this exists: the list function ends with `return , @($Items)`, and depending on how the
	result is captured that can arrive at the caller either as a flat array of N items or as a
	1-element array whose single element is the real N-item array. In the nested case every count
	in the export/copy loops reads 1 - TotalItems shows 1, the batching loop runs a single pass so
	the progress bar never moves off its first update, and `$batch | ForEach-Object { $_.id }`
	member-enumerates the inner array to produce all N ids in one request, which exportItems
	rejects with "a maximum of 20 item ids per request".

	Rather than depend on the shape being one thing or the other, unwrap while the collection is
	a single element that is itself a non-string, non-dictionary collection. A folder holding
	exactly one item is unaffected: a mailboxItem PSCustomObject isn't enumerable, so the loop
	stops immediately and the single item is preserved.
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)] $InputObject
	)
	Begin {
		$result = @($InputObject)
		while ($result.Count -eq 1 -and
			$null -ne $result[0] -and
			$result[0] -is [System.Collections.IEnumerable] -and
			$result[0] -isnot [String] -and
			$result[0] -isnot [System.Collections.IDictionary]) {
			Write-Verbose "Get-SMCFlatItemList: unwrapping a nested single-element collection."
			$result = @($result[0])
		}
		return , $result
	}
}

function Invoke-SMCExportFolderToDirectory {
	<#
	Exports every item in a folder as individual .fts files into $DestinationPath. Fetches the
	full item list first (ignoring the GUI's "# of Items" display cap - this is a full-folder
	operation, not the list view), batches exportItems calls at 20 ids per request (its max), and
	writes one uniquely-named file per item. A batch failure is reported and skipped rather than
	aborting the whole export, so one bad batch can't stop the rest of the folder from exporting.
	Returns a report object with counts.
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [String]$MailboxId,
		[Parameter(Position = 1, Mandatory = $true)] [String]$FolderId,
		[Parameter(Position = 2, Mandatory = $true)] [String]$DestinationPath,
		[Parameter(Position = 3, Mandatory = $false)] [System.Windows.Forms.Form]$ProgressForm
	)
	Process {
		$report = [PSCustomObject]@{ TotalItems = 0; Exported = 0; Failed = 0; Errors = @() }

		Write-Progress -Activity "Exporting folder" -Status "Listing items..."
		if ($ProgressForm) { Update-SMCProgressForm -ProgressForm $ProgressForm -Status "Listing items..." -PercentComplete 0 }
		# Normalise the returned list to a flat array before anything counts or slices it - see
		# Get-SMCFlatItemList for why the shape can't be assumed.
		$items = Get-SMCFlatItemList (Invoke-SMCListFolderItems -MailboxId $MailboxId -FolderId $FolderId -ItemCount 0)
		$report.TotalItems = $items.Count
		if ($items.Count -eq 0) {
			Write-Progress -Activity "Exporting folder" -Completed
			return $report
		}

		$index = 0
		for ($i = 0; $i -lt $items.Count; $i += 20) {
			$batch = $items[$i..([Math]::Min($i + 19, $items.Count - 1))]
			$percent = [Math]::Min(100, [int](($i / $items.Count) * 100))
			$statusText = "Items $($i + 1)-$($i + $batch.Count) of $($items.Count)"
			Write-Progress -Activity "Exporting folder" -Status $statusText -PercentComplete $percent
			if ($ProgressForm) { Update-SMCProgressForm -ProgressForm $ProgressForm -Status $statusText -PercentComplete $percent }
			try {
				$exported = Invoke-SMCExportItems -MailboxId $MailboxId -ItemIds @($batch | ForEach-Object { $_.id })
			}
			catch {
				$report.Failed += $batch.Count
				$report.Errors += "Batch starting at item $($i + 1): $($_.Exception.Message)"
				continue
			}
			if ($exported.Count -ne $batch.Count) {
				$report.Errors += "Batch starting at item $($i + 1): expected $($batch.Count) result(s), got $($exported.Count) - some items in this batch may be missing below."
			}
			# Match exported blobs back to their source item by id so the filename can use the
			# real subject - exportItems' response order isn't guaranteed to match the request.
			$bySubject = @{}
			foreach ($srcItem in $batch) { $bySubject[$srcItem.id] = $srcItem.Subject }
			foreach ($exportedItem in $exported) {
				$index++
				if (-not $exportedItem.data) {
					$report.Failed++
					$report.Errors += "Item $($exportedItem.itemId): export returned no data."
					continue
				}
				try {
					$fileName = Get-SMCSafeFileName -Index $index -Subject $bySubject[$exportedItem.itemId] -ItemId $exportedItem.itemId
					$fullPath = Join-Path -Path $DestinationPath -ChildPath $fileName
					[IO.File]::WriteAllBytes($fullPath, [Convert]::FromBase64String($exportedItem.data))
					$report.Exported++
				}
				catch {
					$report.Failed++
					$report.Errors += "Item $($exportedItem.itemId): $($_.Exception.Message)"
				}
			}
			# Second update, at the END of the batch, based on items actually completed. The
			# update above is computed from the batch's start index, so on a folder small enough
			# to fit in one batch it is always 0 and nothing ever moves it; reporting completed
			# work here is what makes the bar advance and finish at 100.
			$done = [Math]::Min($i + $batch.Count, $items.Count)
			$donePercent = [Math]::Min(100, [int](($done / $items.Count) * 100))
			$doneText = "$done of $($items.Count) item(s) processed"
			Write-Progress -Activity "Exporting folder" -Status $doneText -PercentComplete $donePercent
			if ($ProgressForm) { Update-SMCProgressForm -ProgressForm $ProgressForm -Status $doneText -PercentComplete $donePercent }
		}
		Write-Progress -Activity "Exporting folder" -Completed
		return $report
	}
}

function Invoke-SMCCreateImportSession {
	<#
	Starts an import session for the mailbox and returns the mailboxItemImportSession object
	(.importUrl, .expirationDateTime). importUrl is opaque and preauthenticated - see
	Invoke-SMCImportItemData for why it's called with no Authorization header.
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [String]$MailboxId
	)
	Process {
		$RequestURL = "https://graph.microsoft.com/beta/admin/exchange/mailboxes/$MailboxId/createImportSession"
		return Invoke-MgGraphRequest -Method Post -Uri $RequestURL
	}
}

function Invoke-SMCImportItemData {
	<#
	Uploads one already-base64-encoded FTS blob into $FolderId via an import session's opaque
	importUrl. This is the in-memory half of the import path: Invoke-SMCImportItem reads a .fts
	file off disk and hands the encoded bytes here, and Invoke-SMCCopyFolderItems hands over the
	base64 string exportItems just returned without ever touching the filesystem.

	The importUrl already carries its own auth token, so - per the Graph docs - this deliberately
	does NOT send an Authorization header, and uses Invoke-RestMethod rather than
	Invoke-MgGraphRequest (which would try to attach the signed-in Graph token instead).
	Retries on HTTP 429 using the server's Retry-After hint.
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [String]$ImportUrl,
		[Parameter(Position = 1, Mandatory = $true)] [String]$FolderId,
		[Parameter(Position = 2, Mandatory = $true)] [String]$Base64Data,
		[Parameter(Position = 3, Mandatory = $false)] [Int]$RetryCount = 0
	)
	Process {
		$Request = @{
			FolderId = $FolderId
			Mode     = "create"
			Data     = $Base64Data
		}
		try {
			return Invoke-RestMethod -Method Post -Uri $ImportUrl -Body ($Request | ConvertTo-Json -Depth 5) -ContentType "application/json"
		}
		catch {
			$resp = $_.Exception.Response
			if ($resp -and [int]$resp.StatusCode -eq 429 -and $RetryCount -lt 3) {
				$retryAfter = 5
				if ($resp.Headers -and $resp.Headers["Retry-After"]) { $retryAfter = [int]$resp.Headers["Retry-After"] }
				Write-Verbose "Throttled on import - waiting $retryAfter seconds (retry $($RetryCount + 1)/3)"
				Start-Sleep -Seconds $retryAfter
				return Invoke-SMCImportItemData -ImportUrl $ImportUrl -FolderId $FolderId -Base64Data $Base64Data -RetryCount ($RetryCount + 1)
			}
			throw
		}
	}
}

function Invoke-SMCImportItem {
	<#
	Uploads a single previously-exported .fts file into $FolderId via an import session's opaque
	importUrl. Thin wrapper over Invoke-SMCImportItemData - reads the file, encodes it, and lets
	that function deal with the request and the 429 retry.
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [String]$ImportUrl,
		[Parameter(Position = 1, Mandatory = $true)] [String]$FolderId,
		[Parameter(Position = 2, Mandatory = $true)] [String]$FilePath
	)
	Process {
		$encoded = [Convert]::ToBase64String([IO.File]::ReadAllBytes($FilePath))
		return Invoke-SMCImportItemData -ImportUrl $ImportUrl -FolderId $FolderId -Base64Data $encoded
	}
}

function Invoke-SMCCopyFolderItems {
	<#
	.SYNOPSIS
	Copies every item from a source folder into a destination folder, in either the same mailbox
	or a different one, by round-tripping FTS blobs through memory - nothing is written to disk.

	.DESCRIPTION
	Lists the source folder once, then works in chunks of 20 (exportItems' per-request maximum):
	export a chunk, immediately import each blob in that chunk into the destination, then let the
	chunk go out of scope before fetching the next. Peak memory is one chunk of blobs regardless
	of how big the folder is, which is the point of interleaving rather than exporting everything
	up front.

	Because the copy is a create-import, the destination gets NEW items with new ids - this is a
	copy, not a move, and re-running it will duplicate. Nothing is deleted from the source.

	One import session is created for the destination mailbox up front. Sessions expire, so a
	failed import triggers exactly one retry against a freshly created session before the item is
	counted as failed; a genuine failure (bad blob, permissions) then fails twice and is recorded
	rather than retried forever.

	A chunk whose export fails is recorded and skipped rather than aborting the copy, matching how
	the folder export behaves.

	.OUTPUTS
	A report object with TotalItems, Copied, Failed and Errors.
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [String]$SourceMailboxId,
		[Parameter(Position = 1, Mandatory = $true)] [String]$SourceFolderId,
		[Parameter(Position = 2, Mandatory = $true)] [String]$DestinationMailboxId,
		[Parameter(Position = 3, Mandatory = $true)] [String]$DestinationFolderId,
		[Parameter(Position = 4, Mandatory = $false)] [System.Windows.Forms.Form]$ProgressForm
	)
	Process {
		$report = [PSCustomObject]@{ TotalItems = 0; Copied = 0; Failed = 0; Errors = @() }

		Write-Progress -Activity "Copying folder" -Status "Listing source items..."
		if ($ProgressForm) { Update-SMCProgressForm -ProgressForm $ProgressForm -Status "Listing source items..." -PercentComplete 0 }
		# Same normalisation as the export path - see Get-SMCFlatItemList.
		$items = Get-SMCFlatItemList (Invoke-SMCListFolderItems -MailboxId $SourceMailboxId -FolderId $SourceFolderId -ItemCount 0)
		$report.TotalItems = $items.Count
		if ($items.Count -eq 0) {
			Write-Progress -Activity "Copying folder" -Completed
			return $report
		}

		if ($ProgressForm) { Update-SMCProgressForm -ProgressForm $ProgressForm -Status "Starting import session..." -PercentComplete 0 }
		$session = Invoke-SMCCreateImportSession -MailboxId $DestinationMailboxId

		for ($i = 0; $i -lt $items.Count; $i += 20) {
			$batch = $items[$i..([Math]::Min($i + 19, $items.Count - 1))]
			$startPercent = [Math]::Min(100, [int](($i / $items.Count) * 100))
			$exportText = "Exporting items $($i + 1)-$($i + $batch.Count) of $($items.Count)"
			Write-Progress -Activity "Copying folder" -Status $exportText -PercentComplete $startPercent
			if ($ProgressForm) { Update-SMCProgressForm -ProgressForm $ProgressForm -Status $exportText -PercentComplete $startPercent }

			try {
				$exported = Invoke-SMCExportItems -MailboxId $SourceMailboxId -ItemIds @($batch | ForEach-Object { $_.id })
			}
			catch {
				$report.Failed += $batch.Count
				$report.Errors += "Chunk starting at item $($i + 1): export failed - $($_.Exception.Message)"
				continue
			}
			if ($exported.Count -ne $batch.Count) {
				$report.Errors += "Chunk starting at item $($i + 1): exported $($exported.Count) of $($batch.Count) requested item(s)."
			}

			$importText = "Importing items $($i + 1)-$($i + $batch.Count) of $($items.Count)"
			Write-Progress -Activity "Copying folder" -Status $importText -PercentComplete $startPercent
			if ($ProgressForm) { Update-SMCProgressForm -ProgressForm $ProgressForm -Status $importText -PercentComplete $startPercent }

			foreach ($exportedItem in $exported) {
				if (-not $exportedItem.data) {
					$report.Failed++
					$report.Errors += "Item $($exportedItem.itemId): export returned no data."
					continue
				}
				try {
					[void](Invoke-SMCImportItemData -ImportUrl $session.importUrl -FolderId $DestinationFolderId -Base64Data $exportedItem.data)
					$report.Copied++
				}
				catch {
					# Most likely cause on a long copy is an expired import session, which isn't
					# distinguishable from a genuine failure without parsing error bodies - so just
					# take a fresh session and try the item once more.
					Write-Verbose "Import failed for $($exportedItem.itemId), retrying with a new import session: $($_.Exception.Message)"
					try {
						$session = Invoke-SMCCreateImportSession -MailboxId $DestinationMailboxId
						[void](Invoke-SMCImportItemData -ImportUrl $session.importUrl -FolderId $DestinationFolderId -Base64Data $exportedItem.data)
						$report.Copied++
					}
					catch {
						$report.Failed++
						$report.Errors += "Item $($exportedItem.itemId): import failed - $($_.Exception.Message)"
					}
				}
			}

			# Release this chunk's blobs before the next export - they can be several MB each and
			# holding them costs nothing useful once they're imported.
			$exported = $null

			$done = [Math]::Min($i + $batch.Count, $items.Count)
			$donePercent = [Math]::Min(100, [int](($done / $items.Count) * 100))
			$doneText = "$done of $($items.Count) item(s) copied"
			Write-Progress -Activity "Copying folder" -Status $doneText -PercentComplete $donePercent
			if ($ProgressForm) { Update-SMCProgressForm -ProgressForm $ProgressForm -Status $doneText -PercentComplete $donePercent }
		}
		Write-Progress -Activity "Copying folder" -Completed
		return $report
	}
}

function Import-SMCItemsToCalendar {
	<#
	.SYNOPSIS
	Imports one or more previously-exported .fts item files straight into the mailbox's Calendar
	folder, regardless of whatever folder is currently selected in the GUI's tree view.

	.EXAMPLE
	Import-SMCItemsToCalendar -MailboxId $Script:MailboxId -FilePaths "C:\exports\meeting1.fts","C:\exports\meeting2.fts"
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [String]$MailboxId,
		[Parameter(Position = 1, Mandatory = $true)] [String[]]$FilePaths
	)
	Process {
		try {
			$calendarFolder = Invoke-SMCGetMailboxFolder -MailboxId $MailboxId -FolderId "calendar"
		}
		catch {
			# Fall back to finding it by folder type if the "calendar" well-known id isn't recognized
			# for this mailbox - same pattern used to locate Sticky Note folders by type.
			Write-Verbose "Well-known 'calendar' folder id didn't resolve - falling back to a type search."
			$RequestURL = "https://graph.microsoft.com/beta/admin/exchange/mailboxes/$MailboxId/folders?`$filter=type eq 'IPF.Appointment'"
			$Results = Invoke-MgGraphRequest -Method Get -Uri $RequestURL
			$calendarFolder = $Results.Value | Select-Object -First 1
			if (-not $calendarFolder) {
				throw "Could not resolve the Calendar folder for mailbox $MailboxId."
			}
		}

		$session = Invoke-SMCCreateImportSession -MailboxId $MailboxId
		$results = foreach ($file in $FilePaths) {
			Write-Verbose "Importing $file into Calendar ($($calendarFolder.id))"
			Invoke-SMCImportItem -ImportUrl $session.importUrl -FolderId $calendarFolder.id -FilePath $file
		}
		return $results
	}
}

#endregion

#region ---------------------------- Item viewer property plumbing ----------------------------
# Everything the message viewer needs to turn a thin mailboxItem into something readable. The API
# gives back id/type/size/timestamps and nothing else, so every human-visible field below is a
# tagged MAPI property fetched through $expand=singleValueExtendedProperties.

function Get-SMCViewerPropertyGroups {
	<#
	Returns the tagged properties the viewer asks for, split into independent groups.

	Splitting matters: these go out as separate GETs, so one group being rejected (an unsupported
	tag, a type mismatch, a property the store won't hand back for this item class) costs that
	group only. A single combined request would fail whole and leave the viewer with nothing,
	which is roughly what "doesn't show the message correctly" looks like from the outside.

	Group order is deliberate - core first, so the header is populated even if later groups fail.
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)] [String]$MessageClass
	)
	Process {
		$groups = @()

		# Core envelope: subject, class, delivery time, sender, recipients, attachment flag.
		$core = @()
		$core += (Get-SMCTaggedProperty -DataType String -Id "0x37")       # PR_SUBJECT
		$core += (Get-SMCTaggedProperty -DataType String -Id "0x1A")       # PR_MESSAGE_CLASS
		$core += (Get-SMCTaggedProperty -DataType SystemTime -Id "0xE06")  # PR_MESSAGE_DELIVERY_TIME
		$core += (Get-SMCTaggedProperty -DataType String -Id "0xC1A")      # PR_SENDER_NAME
		$core += (Get-SMCTaggedProperty -DataType String -Id "0xC1F")      # PR_SENDER_EMAIL_ADDRESS
		$core += (Get-SMCTaggedProperty -DataType String -Id "0xE04")      # PR_DISPLAY_TO
		$core += (Get-SMCTaggedProperty -DataType Boolean -Id "0xE1B")     # PR_HASATTACH
		$groups += , $core

		# Secondary envelope: sent time, on-behalf-of, Cc, importance, size, body code page.
		$extra = @()
		$extra += (Get-SMCTaggedProperty -DataType SystemTime -Id "0x39")  # PR_CLIENT_SUBMIT_TIME
		$extra += (Get-SMCTaggedProperty -DataType String -Id "0x42")      # PR_SENT_REPRESENTING_NAME
		$extra += (Get-SMCTaggedProperty -DataType String -Id "0x65")      # PR_SENT_REPRESENTING_EMAIL_ADDRESS
		$extra += (Get-SMCTaggedProperty -DataType String -Id "0xE03")     # PR_DISPLAY_CC
		$extra += (Get-SMCTaggedProperty -DataType Integer -Id "0x17")     # PR_IMPORTANCE
		$extra += (Get-SMCTaggedProperty -DataType Integer -Id "0xE08")    # PR_MESSAGE_SIZE
		$extra += (Get-SMCTaggedProperty -DataType Integer -Id "0x3FDE")   # PR_INTERNET_CPID
		$groups += , $extra

		# Plain-text body on its own - it's the fallback when there's no HTML, and it's big enough
		# that it's worth not dragging into the envelope requests.
		$groups += , @(Get-SMCTaggedProperty -DataType String -Id "0x1000")  # PR_BODY

		# PR_HTML is Binary and can be large; isolate it so a failure just means "no HTML body".
		$groups += , @(Get-SMCTaggedProperty -DataType Binary -Id "0x1013")  # PR_HTML

		# Class-specific tagged properties. Most appointment/contact/task detail lives in NAMED
		# properties, which this API's $filter syntax has no documented way to request - so this
		# is deliberately limited to the tagged ones that are commonly populated.
		$class = if ($MessageClass) { $MessageClass.ToUpperInvariant() } else { "" }
		if ($class.StartsWith("IPM.APPOINTMENT") -or $class.StartsWith("IPM.SCHEDULE")) {
			$appt = @()
			$appt += (Get-SMCTaggedProperty -DataType SystemTime -Id "0x60")  # PR_START_DATE
			$appt += (Get-SMCTaggedProperty -DataType SystemTime -Id "0x61")  # PR_END_DATE
			$appt += (Get-SMCTaggedProperty -DataType String -Id "0x3A")      # PR_RCVD_REPRESENTING_NAME
			$groups += , $appt
		}
		elseif ($class.StartsWith("IPM.CONTACT")) {
			$contact = @()
			$contact += (Get-SMCTaggedProperty -DataType String -Id "0x3001")  # PR_DISPLAY_NAME
			$contact += (Get-SMCTaggedProperty -DataType String -Id "0x3A16")  # PR_COMPANY_NAME
			$contact += (Get-SMCTaggedProperty -DataType String -Id "0x3A17")  # PR_TITLE
			$contact += (Get-SMCTaggedProperty -DataType String -Id "0x3A08")  # PR_BUSINESS_TELEPHONE_NUMBER
			$contact += (Get-SMCTaggedProperty -DataType String -Id "0x3A1C")  # PR_MOBILE_TELEPHONE_NUMBER
			$groups += , $contact
		}
		elseif ($class.StartsWith("IPM.TASK")) {
			$task = @()
			$task += (Get-SMCTaggedProperty -DataType SystemTime -Id "0x60")   # PR_START_DATE
			$task += (Get-SMCTaggedProperty -DataType SystemTime -Id "0x61")   # PR_END_DATE (due)
			$groups += , $task
		}

		return , $groups
	}
}

function Get-SMCItemViewerDetail {
	<#
	Fetches an item's viewer property set as a bag keyed by the same "DataType 0xTag" string Graph
	uses, e.g. $bag["String 0x37"]. Each property group is a separate request wrapped in its own
	try/catch, so a group the store won't serve degrades that section of the viewer instead of
	blanking the whole window.

	Returns an object with:
	    Item       - the native mailboxItem from the first successful request (id, size, timestamps)
	    Props      - hashtable of tagged property values
	    Failed     - descriptions of any group that couldn't be retrieved
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [String]$MailboxId,
		[Parameter(Position = 1, Mandatory = $true)] [String]$FolderId,
		[Parameter(Position = 2, Mandatory = $true)] [String]$ItemId,
		[Parameter(Position = 3, Mandatory = $false)] [String]$MessageClass
	)
	Process {
		$bag = @{}
		$nativeItem = $null
		$failed = @()

		$groups = Get-SMCViewerPropertyGroups -MessageClass $MessageClass
		foreach ($group in $groups) {
			if (-not $group -or @($group).Count -eq 0) { continue }
			try {
				$fetched = Invoke-SMCGetItemDetail -MailboxId $MailboxId -FolderId $FolderId -ItemId $ItemId -PropertyList $group
			}
			catch {
				$ids = (@($group) | ForEach-Object { $_.Id }) -join ", "
				Write-Verbose "Viewer property group ($ids) failed: $($_.Exception.Message)"
				$failed += $ids
				continue
			}
			if (-not $nativeItem) { $nativeItem = $fetched }
			foreach ($prop in $fetched.singleValueExtendedProperties) {
				# Only record properties the store actually returned a value for - an absent
				# property comes back either missing entirely or with a null value, and keeping
				# nulls out of the bag lets the viewer just test ContainsKey.
				if ($null -ne $prop.Value -and "$($prop.Value)" -ne "") {
					$bag["$($prop.Id)"] = $prop.Value
				}
			}
		}

		return [PSCustomObject]@{
			Item   = $nativeItem
			Props  = $bag
			Failed = $failed
		}
	}
}

function Get-SMCProp {
	<# Reads one tagged property out of a viewer bag, returning $null when it wasn't retrieved. #>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [Hashtable]$Bag,
		[Parameter(Position = 1, Mandatory = $true)] [String]$Key
	)
	Begin {
		if ($Bag.ContainsKey($Key)) { return $Bag[$Key] }
		return $null
	}
}

function Format-SMCPropDate {
	<# SystemTime properties arrive as ISO strings; show them in local time, or blank if unparseable. #>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)] $Value
	)
	Begin {
		if ($null -eq $Value -or "$Value" -eq "") { return "" }
		try { return ([DateTime]$Value).ToLocalTime().ToString("dddd, d MMMM yyyy h:mm tt") }
		catch { return "$Value" }
	}
}

function ConvertFrom-SMCHtmlProperty {
	<#
	PR_HTML (0x1013) is a Binary property, so Graph hands it back base64-encoded. The bytes are
	the raw HTML in whatever code page PR_INTERNET_CPID (0x3FDE) says - commonly 65001 (UTF-8) or
	1252 - so decoding blind as UTF-8 is what turns curly quotes and accented characters into
	mojibake.

	Note for PowerShell 7: single-byte code pages like 1252 aren't registered by default on .NET
	Core, so GetEncoding throws; the catch falls back to UTF-8 rather than failing the viewer.
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [String]$Base64Data,
		[Parameter(Position = 1, Mandatory = $false)] $CodePage
	)
	Process {
		$bytes = [Convert]::FromBase64String($Base64Data)
		$encoding = [System.Text.Encoding]::UTF8
		if ($CodePage) {
			try { $encoding = [System.Text.Encoding]::GetEncoding([int]$CodePage) }
			catch { Write-Verbose "Code page $CodePage unavailable, decoding as UTF-8 instead." }
		}
		$html = $encoding.GetString($bytes)
		# Strip a leading BOM if the encoding left one in - WebBrowser renders it as a stray glyph.
		return $html.TrimStart([char]0xFEFF)
	}
}

function Get-SMCImportanceText {
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)] $Value
	)
	Begin {
		if ($null -eq $Value -or "$Value" -eq "") { return "" }
		switch ([int]$Value) {
			0 { return "Low" }
			1 { return "Normal" }
			2 { return "High" }
			default { return "$Value" }
		}
	}
}

function Get-SMCViewerHeaderFields {
	<#
	Builds the ordered label/value pairs shown above the body, varying by message class: a contact
	gets company and phone numbers, an appointment gets start/end, everything else gets the mail
	envelope. Empty values are dropped so the header doesn't carry rows of blank labels.
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [Hashtable]$Props,
		[Parameter(Position = 1, Mandatory = $false)] $NativeItem
	)
	Process {
		$fields = [System.Collections.Generic.List[Object]]::new()
		$addField = {
			param($label, $value)
			if ($null -ne $value -and "$value".Trim() -ne "") {
				$fields.Add([PSCustomObject]@{ Label = $label; Value = "$value" })
			}
		}

		$class = "$(Get-SMCProp -Bag $Props -Key 'String 0x1A')"

		$senderName = Get-SMCProp -Bag $Props -Key "String 0xC1A"
		$senderMail = Get-SMCProp -Bag $Props -Key "String 0xC1F"
		$fromText = if ($senderName -and $senderMail) { "$senderName <$senderMail>" }
		elseif ($senderName) { "$senderName" }
		else { "$senderMail" }

		$onBehalfName = Get-SMCProp -Bag $Props -Key "String 0x42"
		$onBehalfMail = Get-SMCProp -Bag $Props -Key "String 0x65"
		# Only show "on behalf of" when it's actually a different party - it's populated with the
		# sender's own details on ordinary mail, where repeating it is just noise.
		$onBehalfText = ""
		if ($onBehalfMail -and "$onBehalfMail" -ne "$senderMail") {
			$onBehalfText = if ($onBehalfName) { "$onBehalfName <$onBehalfMail>" } else { "$onBehalfMail" }
		}

		if ($class -match '^IPM\.Contact') {
			& $addField "Name" (Get-SMCProp -Bag $Props -Key "String 0x3001")
			& $addField "Job title" (Get-SMCProp -Bag $Props -Key "String 0x3A17")
			& $addField "Company" (Get-SMCProp -Bag $Props -Key "String 0x3A16")
			& $addField "Business" (Get-SMCProp -Bag $Props -Key "String 0x3A08")
			& $addField "Mobile" (Get-SMCProp -Bag $Props -Key "String 0x3A1C")
			& $addField "File as" (Get-SMCProp -Bag $Props -Key "String 0x37")
		}
		elseif ($class -match '^IPM\.(Appointment|Schedule)') {
			& $addField "Subject" (Get-SMCProp -Bag $Props -Key "String 0x37")
			& $addField "Organiser" $fromText
			& $addField "Start" (Format-SMCPropDate (Get-SMCProp -Bag $Props -Key "SystemTime 0x60"))
			& $addField "End" (Format-SMCPropDate (Get-SMCProp -Bag $Props -Key "SystemTime 0x61"))
			& $addField "Required" (Get-SMCProp -Bag $Props -Key "String 0xE04")
			& $addField "Optional" (Get-SMCProp -Bag $Props -Key "String 0xE03")
		}
		elseif ($class -match '^IPM\.Task') {
			& $addField "Subject" (Get-SMCProp -Bag $Props -Key "String 0x37")
			& $addField "Owner" $fromText
			& $addField "Start" (Format-SMCPropDate (Get-SMCProp -Bag $Props -Key "SystemTime 0x60"))
			& $addField "Due" (Format-SMCPropDate (Get-SMCProp -Bag $Props -Key "SystemTime 0x61"))
		}
		else {
			& $addField "From" $fromText
			& $addField "On behalf of" $onBehalfText
			& $addField "To" (Get-SMCProp -Bag $Props -Key "String 0xE04")
			& $addField "Cc" (Get-SMCProp -Bag $Props -Key "String 0xE03")
			& $addField "Subject" (Get-SMCProp -Bag $Props -Key "String 0x37")
			$sent = Format-SMCPropDate (Get-SMCProp -Bag $Props -Key "SystemTime 0x39")
			$received = Format-SMCPropDate (Get-SMCProp -Bag $Props -Key "SystemTime 0xE06")
			& $addField "Sent" $sent
			& $addField "Received" $received
		}

		$importance = Get-SMCImportanceText (Get-SMCProp -Bag $Props -Key "Integer 0x17")
		if ($importance -and $importance -ne "Normal") { & $addField "Importance" $importance }

		$hasAttach = Get-SMCProp -Bag $Props -Key "Boolean 0xE1B"
		if ($null -ne $hasAttach) {
			try { $attachFlag = [System.Convert]::ToBoolean($hasAttach) } catch { $attachFlag = $false }
			if ($attachFlag) { & $addField "Attachments" "Yes - export the item to get them" }
		}

		$size = Get-SMCProp -Bag $Props -Key "Integer 0xE08"
		if (-not $size -and $NativeItem) { $size = $NativeItem.size }
		if ($size) { & $addField "Size" ("{0:N0} KB" -f [math]::Round([double]$size / 1KB, 0)) }

		& $addField "Message class" $class

		return , $fields
	}
}

#endregion

#region ---------------------------- Standard Graph Mail API (sending only) ----------------------------
# The import/export API has no send capability - New Message / Send Message still goes through the
# regular Mail API (Send-MgUserMail) against the mailbox's UPN, which is unrelated to the MailboxId
# resolved above. Carried over unchanged from the original GUI script.

function New-SMCEmailAddress {
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)] [string]$Name,
		[Parameter(Position = 1, Mandatory = $true)] [string]$Address
	)
	Begin {
		$EmailAddress = "" | Select-Object Name, Address
		$EmailAddress.Name = if ([String]::IsNullOrEmpty($Name)) { $Address } else { $Name }
		$EmailAddress.Address = $Address
		return , $EmailAddress
	}
}

function Send-SMCMessageREST {
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [string]$MailboxUpn,
		[Parameter(Position = 1, Mandatory = $true)] [String]$Subject,
		[Parameter(Position = 2, Mandatory = $false)] [String]$Body,
		[Parameter(Position = 3, Mandatory = $false)] [psobject]$ToRecipients,
		[Parameter(Position = 4, Mandatory = $false)] [psobject]$Attachments
	)
	Begin {
		$MessageBody = @{
			message = @{
				subject      = $Subject
				body         = @{ contentType = "HTML"; content = $Body }
				toRecipients = @($ToRecipients | ForEach-Object { @{ emailAddress = @{ name = $_.Name; address = $_.Address } } })
			}
			saveToSentItems = $true
		}
		if ($Attachments -and $Attachments.Count -gt 0) {
			$MessageBody.message.attachments = @($Attachments | ForEach-Object {
					$Item = Get-Item $_
					@{
						"@odata.type" = "#microsoft.graph.fileAttachment"
						name          = $Item.Name
						contentBytes  = [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($_))
					}
				})
		}
		Send-MgUserMail -UserId $MailboxUpn -BodyParameter ($MessageBody | ConvertTo-Json -Depth 10)
	}
}

#endregion

#region ---------------------------- GUI ----------------------------

function Add-SMCFolderTreeNodes {
	<#
	Recursively populates a TreeNode with child mailboxFolders. Each folder object gets tagged
	with the MailboxId/MailboxUpn it belongs to (via Add-Member) before being stored as the
	TreeNode's Tag - this is what lets Get-SMCClientFolderItems and friends always query the
	right mailbox for whatever node is selected, once more than one mailbox is in the same tree,
	instead of relying on a single "current mailbox" variable that only ever pointed at one.
	It's also what makes a cross-mailbox copy work: the folder picker hands back a Tag that
	already knows which mailbox it belongs to.
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [System.Windows.Forms.TreeNode]$ParentNode,
		[Parameter(Position = 1, Mandatory = $true)] [String]$MailboxId,
		[Parameter(Position = 2, Mandatory = $true)] [String]$MailboxUpn,
		[Parameter(Position = 3, Mandatory = $true)] [String]$FolderId
	)
	Process {
		$children = Invoke-SMCGetChildFolders -MailboxId $MailboxId -FolderId $FolderId
		foreach ($child in $children) {
			Add-Member -InputObject $child -NotePropertyName MailboxId -NotePropertyValue $MailboxId -Force
			Add-Member -InputObject $child -NotePropertyName MailboxUpn -NotePropertyValue $MailboxUpn -Force
			$node = New-Object System.Windows.Forms.TreeNode($child.displayName)
			$node.Name = $child.displayName
			$node.Tag = $child
			[void]$ParentNode.Nodes.Add($node)
			if ($child.childFolderCount -gt 0) {
				Add-SMCFolderTreeNodes -ParentNode $node -MailboxId $MailboxId -MailboxUpn $MailboxUpn -FolderId $child.id
			}
		}
	}
}

function Copy-SMCTreeNodesForPicker {
	<#
	Clones a set of TreeNodes (text + Tag, recursively) into another node collection. Used to fill
	the destination-folder picker from the main tree without re-querying Graph: the Tag objects are
	shared by reference, so a picked node hands back exactly the same folder object - MailboxId and
	all - that the main tree is holding.
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] $SourceNodes,
		[Parameter(Position = 1, Mandatory = $true)] $TargetNodes
	)
	Process {
		foreach ($srcNode in $SourceNodes) {
			$clone = New-Object System.Windows.Forms.TreeNode($srcNode.Text)
			$clone.Tag = $srcNode.Tag
			[void]$TargetNodes.Add($clone)
			if ($srcNode.Nodes.Count -gt 0) {
				Copy-SMCTreeNodesForPicker -SourceNodes $srcNode.Nodes -TargetNodes $clone.Nodes
			}
		}
	}
}

function Show-SMCFolderPicker {
	<#
	Modal picker showing every mailbox/folder currently in the main tree, so a copy destination can
	be any folder in any open mailbox. Returns the selected folder object (the node's Tag, carrying
	id/displayName/MailboxId/MailboxUpn) or $null if cancelled or nothing usable was selected.
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)] [String]$Title = "Choose a destination folder",
		[Parameter(Position = 1, Mandatory = $false)] [String]$Prompt = "Select the folder to copy into:"
	)
	Process {
		$script:smcPickedFolder = $null

		$pickerForm = New-Object System.Windows.Forms.Form
		$pickerForm.Text = $Title
		$pickerForm.Size = New-Object System.Drawing.Size(460, 560)
		$pickerForm.StartPosition = "CenterParent"
		$pickerForm.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::None
		$pickerForm.MinimizeBox = $false
		$pickerForm.MaximizeBox = $false

		$lbl = New-Object System.Windows.Forms.Label
		$lbl.Dock = "Top"
		$lbl.Height = 24
		$lbl.Text = $Prompt
		$lbl.Padding = New-Object System.Windows.Forms.Padding(6, 4, 0, 0)

		$buttonPanel = New-Object System.Windows.Forms.Panel
		$buttonPanel.Dock = "Bottom"
		$buttonPanel.Height = 42

		$btnOk = New-Object System.Windows.Forms.Button
		$btnOk.Text = "OK"
		$btnOk.Size = New-Object System.Drawing.Size(90, 26)
		$btnOk.Location = New-Object System.Drawing.Point(250, 8)

		$btnCancel = New-Object System.Windows.Forms.Button
		$btnCancel.Text = "Cancel"
		$btnCancel.Size = New-Object System.Drawing.Size(90, 26)
		$btnCancel.Location = New-Object System.Drawing.Point(348, 8)

		[void]$buttonPanel.Controls.Add($btnOk)
		[void]$buttonPanel.Controls.Add($btnCancel)

		$pickerTree = New-Object System.Windows.Forms.TreeView
		$pickerTree.Dock = "Fill"
		Copy-SMCTreeNodesForPicker -SourceNodes $tvTreView.Nodes -TargetNodes $pickerTree.Nodes
		foreach ($rootNode in $pickerTree.Nodes) { $rootNode.Expand() }

		# Dock=Fill must be added first so it doesn't claim the space the docked edges need.
		[void]$pickerForm.Controls.Add($pickerTree)
		[void]$pickerForm.Controls.Add($buttonPanel)
		[void]$pickerForm.Controls.Add($lbl)

		[void]$btnOk.Add_Click({
				$sel = $pickerTree.SelectedNode
				if (-not $sel -or -not $sel.Tag -or -not $sel.Tag.id) {
					[System.Windows.Forms.MessageBox]::Show("Select a folder first.", "No Folder Selected")
					return
				}
				$script:smcPickedFolder = $sel.Tag
				$pickerForm.Close()
			})
		[void]$btnCancel.Add_Click({
				$script:smcPickedFolder = $null
				$pickerForm.Close()
			})
		# Double-clicking a folder is the same as selecting it and pressing OK.
		[void]$pickerTree.Add_DoubleClick({
				$sel = $pickerTree.SelectedNode
				if ($sel -and $sel.Tag -and $sel.Tag.id) {
					$script:smcPickedFolder = $sel.Tag
					$pickerForm.Close()
				}
			})

		[void]$pickerForm.ShowDialog()
		$pickerForm.Dispose()
		return $script:smcPickedFolder
	}
}

function Invoke-SMCAddMailboxToTree {
	<#
	Resolves a mailbox and adds it as a new top-level root in the existing tree, alongside
	whatever mailboxes are already there - this is the "Add Mailbox" button's logic. Returns the
	new root TreeNode (or $null on failure) so the caller can decide what to do with it, e.g.
	scroll it into view.
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [String]$Upn
	)
	Process {
		Write-Progress -Activity "Resolving mailbox" -Status $Upn
		try {
			$mailboxId = Get-SMCMailboxId -Upn $Upn
			$rootFolder = Invoke-SMCGetMailboxFolder -MailboxId $mailboxId -FolderId "MsgFolderRoot"
		}
		catch {
			[System.Windows.Forms.MessageBox]::Show("Couldn't open mailbox '$Upn':`r`n`r`n$($_.Exception.Message)", "Open Mailbox Failed")
			Write-Progress -Activity "Executing Request" -Completed
			return $null
		}
		Add-Member -InputObject $rootFolder -NotePropertyName MailboxId -NotePropertyValue $mailboxId -Force
		Add-Member -InputObject $rootFolder -NotePropertyName MailboxUpn -NotePropertyValue $Upn -Force

		# Kept as a convenience default (e.g. for New Message) reflecting the most recently
		# added/opened mailbox - per-folder operations should use the folder's own tagged
		# MailboxId/MailboxUpn instead, since that's correct regardless of which mailbox's branch
		# is currently selected.
		$Script:MailboxId = $mailboxId
		$Script:MailboxUpn = $Upn

		$TNRoot = New-Object System.Windows.Forms.TreeNode("Mailbox - $Upn")
		$TNRoot.Name = "Mailbox"
		$TNRoot.Tag = $rootFolder
		[void]$tvTreView.Nodes.Add($TNRoot)

		Write-Progress -Activity "Enumerating folders" -Status $Upn
		Add-SMCFolderTreeNodes -ParentNode $TNRoot -MailboxId $mailboxId -MailboxUpn $Upn -FolderId $rootFolder.id
		$TNRoot.Expand()

		Write-Progress -Activity "Executing Request" -Completed
		return $TNRoot
	}
}

function Invoke-SMCOpenMailbox {
	<# "Open Mailbox" button: clears the tree first, then adds just this one mailbox. #>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [String]$Upn
	)
	Process {
		$tvTreView.Nodes.Clear()
		$TNRoot = Invoke-SMCAddMailboxToTree -Upn $Upn
		if ($TNRoot) {
			# Force the view back to the very top - belt-and-braces alongside the layout fixes, in
			# case scroll position (not just available space) was part of why top items were hidden.
			$tvTreView.TopNode = $TNRoot
		}
	}
}

function Invoke-SMCAddMailboxButtonClick {
	<# "Add Mailbox" button: prompts for a UPN and appends it as another root, leaving whatever's already in the tree alone. #>
	[CmdletBinding()]
	Param ()
	$Upn = [Microsoft.VisualBasic.Interaction]::InputBox("Mailbox (UPN) to add:", "Add Mailbox", "")
	if ([String]::IsNullOrWhiteSpace($Upn)) { return }
	$TNRoot = Invoke-SMCAddMailboxToTree -Upn $Upn
	if ($TNRoot) {
		$TNRoot.EnsureVisible()
		$tvTreView.SelectedNode = $TNRoot
	}
}

function Get-SMCClientFolderItems {
	[CmdletBinding()]
	Param ()
	if (-not $Script:lfFolderID) { return }
	$mbtable.Rows.Clear()
	# Rebind from scratch each time rather than relying on the DataGridView noticing the
	# DataTable changed underneath the same reference - cheap, and removes one more variable
	# if items ever appear to "not enumerate".
	$dgDataGrid.DataSource = $null

	try {
		$items = Invoke-SMCListFolderItems -MailboxId $Script:lfFolderID.MailboxId -FolderId $Script:lfFolderID.id -ItemCount ([int]$neResultCheckNum.Value)
	}
	catch {
		[System.Windows.Forms.MessageBox]::Show("Couldn't list items in '$($Script:lfFolderID.displayName)' ($($Script:lfFolderID.id)):`r`n`r`n$($_.Exception.Message)", "List Items Failed")
		$Script:form.Text = "Simple Exchange Mailbox Client (Import/Export API) - $($Script:lfFolderID.MailboxUpn) : failed listing '$($Script:lfFolderID.displayName)'"
		return
	}

	if ($seSearchCheck.Checked -and -not [String]::IsNullOrEmpty($sbSearchTextBox.Text)) {
		$term = $sbSearchTextBox.Text
		switch ($snSearchPropDrop.SelectedItem) {
			"Subject" { $items = $items | Where-Object { $_.Subject -like "*$term*" } }
			"From" { $items = $items | Where-Object { $_.SenderName -like "*$term*" -or $_.SenderAddress -like "*$term*" } }
			"Body" {
				# PR_BODY isn't pulled back in the list view for performance reasons, so a body search
				# has to fetch each candidate item individually - keep this list small.
				$bodyProp = @(Get-SMCTaggedProperty -DataType String -Id "0x1000")
				$items = $items | Where-Object {
					try {
						(Invoke-SMCGetItemDetail -MailboxId $Script:lfFolderID.MailboxId -FolderId $Script:lfFolderID.id -ItemId $_.id -PropertyList $bodyProp).BodyPreview -like "*$term*"
					}
					catch {
						Write-Verbose "Body lookup failed for item $($_.id): $($_.Exception.Message)"
						$false
					}
				}
			}
		}
	}

	foreach ($mail in $items) {
		try {
			$fromDisplay = if ($mail.SenderName) { $mail.SenderName } else { "N/A" }
			$subjectDisplay = if ($mail.Subject) { $mail.Subject } else { "N/A" }
			$sizeKb = [math]::round($mail.size / 1Kb, 0)
			# The "Received" column is a strongly-typed [DATETIME] DataTable column. Not every
			# item actually has PR_MESSAGE_DELIVERY_TIME (meeting responses, hidden associated
			# items, etc. often don't) - Expand-SMCItemProperties then leaves ReceivedDateTime as
			# an empty string, and handing DataTable an empty string for a DateTime column is a
			# real way for Rows.Add() to throw and quietly abort the rest of the batch. Coerce it
			# to DBNull explicitly instead of hoping .NET does something forgiving with "".
			$receivedValue = if ($mail.ReceivedDateTime) {
				try { [DateTime]$mail.ReceivedDateTime } catch { [DBNull]::Value }
			}
			else { [DBNull]::Value }
			[void]$mbtable.Rows.Add($fromDisplay, $subjectDisplay, $receivedValue, $sizeKb, $mail.id, $mail.HasAttachments, $mail.MessageClass)
		}
		catch {
			Write-Verbose "Skipped item $($mail.id) - couldn't add it as a grid row: $($_.Exception.Message)"
		}
	}
	$dgDataGrid.ColumnHeadersVisible = $true
	$dgDataGrid.DataSource = $mbtable
	if ($dgDataGrid.Rows.Count -gt 0) {
		# Belt-and-braces alongside the layout fixes: force the view back to row 0 in case
		# scroll position (not just available space) was part of why top rows were hidden.
		$dgDataGrid.FirstDisplayedScrollingRowIndex = 0
	}
	$dgDataGrid.Refresh()
	$Script:form.Text = "Simple Exchange Mailbox Client (Import/Export API) - $($Script:lfFolderID.MailboxUpn) : $($Script:lfFolderID.displayName) ($($items.Count) item(s), $($mbtable.Rows.Count) shown)"
}

function Get-SMCSelectedGridRow {
	<#
	Returns the DataRowView behind the current grid selection, or $null when nothing is selected.
	Every per-item button used to index DefaultView with $dgDataGrid.CurrentCell.RowIndex directly,
	which throws on a null CurrentCell (no row clicked yet, or an empty folder) instead of saying
	"pick a message first".
	#>
	[CmdletBinding()]
	Param ()
	if (-not $dgDataGrid.CurrentCell) { return $null }
	$rowIndex = $dgDataGrid.CurrentCell.RowIndex
	if ($rowIndex -lt 0 -or $rowIndex -ge $mbtable.DefaultView.Count) { return $null }
	return $mbtable.DefaultView[$rowIndex]
}

function Invoke-SMCExportMessage {
	<# Exports the selected item as an opaque .fts blob - see Invoke-SMCExportItems above. #>
	[CmdletBinding()]
	Param ()
	$row = Get-SMCSelectedGridRow
	if (-not $row) {
		[System.Windows.Forms.MessageBox]::Show("Select a message in the list first.")
		return
	}
	$MessageID = $row[4]
	$saveFileDialog = [System.Windows.Forms.SaveFileDialog]@{
		CheckPathExists  = $true
		OverwritePrompt  = $true
		InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
		FileName         = $row[1]
		Title            = 'Choose where to save the exported item (full-fidelity, not a plain .eml)'
		Filter           = "Fast Transfer stream (*.fts)|*.fts"
	}
	if ($saveFileDialog.ShowDialog() -eq 'OK') {
		$exported = Invoke-SMCExportItems -MailboxId $Script:lfFolderID.MailboxId -ItemIds @($MessageID)
		if ($exported -and $exported[0].data) {
			[IO.File]::WriteAllBytes($saveFileDialog.FileName, [Convert]::FromBase64String($exported[0].data))
		}
		else {
			[System.Windows.Forms.MessageBox]::Show("Export returned no data for this item.")
		}
	}
}

function Invoke-SMCImportFromFts {
	<# GUI handler for the "Import from FTS" button - imports into whichever folder is currently selected in the tree. #>
	[CmdletBinding()]
	Param ()
	if (-not $Script:lfFolderID) {
		[System.Windows.Forms.MessageBox]::Show("Select a destination folder in the tree first.")
		return
	}
	$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{
		Multiselect      = $true
		InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
		Title            = "Choose exported .fts item(s) to import into '$($Script:lfFolderID.displayName)'"
		Filter           = "Fast Transfer stream (*.fts)|*.fts"
	}
	if ($openFileDialog.ShowDialog() -eq 'OK' -and $openFileDialog.FileNames.Count -gt 0) {
		$session = Invoke-SMCCreateImportSession -MailboxId $Script:lfFolderID.MailboxId
		$imported = 0
		$errors = @()
		$progressForm = New-SMCProgressForm -Title "Import from FTS"
		try {
			for ($i = 0; $i -lt $openFileDialog.FileNames.Count; $i++) {
				$file = $openFileDialog.FileNames[$i]
				$percent = [int](($i / $openFileDialog.FileNames.Count) * 100)
				Update-SMCProgressForm -ProgressForm $progressForm -Status "File $($i + 1) of $($openFileDialog.FileNames.Count): $(Split-Path $file -Leaf)" -PercentComplete $percent
				try {
					[void](Invoke-SMCImportItem -ImportUrl $session.importUrl -FolderId $Script:lfFolderID.id -FilePath $file)
					$imported++
				}
				catch {
					$errors += "$(Split-Path $file -Leaf): $($_.Exception.Message)"
				}
			}
		}
		finally {
			Close-SMCProgressForm -ProgressForm $progressForm
		}
		$summary = "Imported $imported of $($openFileDialog.FileNames.Count) item(s) into '$($Script:lfFolderID.displayName)'."
		if ($errors.Count -gt 0) {
			$shownErrors = ($errors | Select-Object -First 5) -join "`r`n"
			$summary += "`r`n`r`n$($errors.Count) file(s) failed:`r`n$shownErrors"
			if ($errors.Count -gt 5) { $summary += "`r`n...and $($errors.Count - 5) more." }
		}
		[System.Windows.Forms.MessageBox]::Show($summary)
		Get-SMCClientFolderItems
	}
}

function Invoke-SMCExportFolderButtonClick {
	<# GUI handler for "Export Folder (.fts)..." - exports every item in the selected folder as individual .fts files. #>
	[CmdletBinding()]
	Param ()
	if (-not $Script:lfFolderID) {
		[System.Windows.Forms.MessageBox]::Show("Select a folder in the tree first.")
		return
	}
	$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
		Description         = "Choose a destination folder for the exported .fts files from '$($Script:lfFolderID.displayName)'"
		ShowNewFolderButton = $true
	}
	if ($folderBrowser.ShowDialog() -ne 'OK') { return }

	$progressForm = New-SMCProgressForm -Title "Export Folder"
	try {
		$report = Invoke-SMCExportFolderToDirectory -MailboxId $Script:lfFolderID.MailboxId -FolderId $Script:lfFolderID.id -DestinationPath $folderBrowser.SelectedPath -ProgressForm $progressForm
	}
	catch {
		[System.Windows.Forms.MessageBox]::Show("Folder export failed:`r`n`r`n$($_.Exception.Message)", "Export Folder Failed")
		return
	}
	finally {
		Close-SMCProgressForm -ProgressForm $progressForm
	}

	$summary = "Exported $($report.Exported) of $($report.TotalItems) item(s) from '$($Script:lfFolderID.displayName)' to:`r`n$($folderBrowser.SelectedPath)"
	if ($report.Failed -gt 0) {
		$shownErrors = ($report.Errors | Select-Object -First 5) -join "`r`n"
		$summary += "`r`n`r`n$($report.Failed) item(s) failed:`r`n$shownErrors"
		if ($report.Errors.Count -gt 5) { $summary += "`r`n...and $($report.Errors.Count - 5) more." }
	}
	[System.Windows.Forms.MessageBox]::Show($summary, "Export Folder Complete")
}

function Invoke-SMCCopyFolderButtonClick {
	<#
	GUI handler for "Copy Folder to...". Source is whatever folder is selected in the tree;
	destination is picked from a modal copy of the same tree, so it can be another folder in the
	same mailbox or a folder in any other mailbox that's been added. Confirms first, since this
	writes real items into someone's mailbox and re-running it duplicates them.
	#>
	[CmdletBinding()]
	Param ()
	if (-not $Script:lfFolderID -or -not $Script:lfFolderID.id) {
		[System.Windows.Forms.MessageBox]::Show("Select the source folder in the tree first.")
		return
	}
	$source = $Script:lfFolderID

	$destination = Show-SMCFolderPicker -Title "Copy Folder - choose destination" -Prompt "Copy items from '$($source.displayName)' into:"
	if (-not $destination) { return }

	if ($destination.MailboxId -eq $source.MailboxId -and $destination.id -eq $source.id) {
		[System.Windows.Forms.MessageBox]::Show("Source and destination are the same folder - that would just duplicate every item back into itself.", "Copy Folder")
		return
	}

	$confirmText = "Copy all items from:`r`n  $($source.MailboxUpn) : $($source.displayName)`r`n`r`nto:`r`n  $($destination.MailboxUpn) : $($destination.displayName)`r`n`r`nThis creates new copies in the destination and leaves the source untouched. Running it twice will duplicate the items.`r`n`r`nContinue?"
	$answer = [System.Windows.Forms.MessageBox]::Show($confirmText, "Copy Folder", [System.Windows.Forms.MessageBoxButtons]::YesNo)
	if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }

	$progressForm = New-SMCProgressForm -Title "Copy Folder"
	try {
		$report = Invoke-SMCCopyFolderItems -SourceMailboxId $source.MailboxId -SourceFolderId $source.id `
			-DestinationMailboxId $destination.MailboxId -DestinationFolderId $destination.id -ProgressForm $progressForm
	}
	catch {
		[System.Windows.Forms.MessageBox]::Show("Folder copy failed:`r`n`r`n$($_.Exception.Message)", "Copy Folder Failed")
		return
	}
	finally {
		Close-SMCProgressForm -ProgressForm $progressForm
	}

	$summary = "Copied $($report.Copied) of $($report.TotalItems) item(s)`r`nfrom '$($source.displayName)' ($($source.MailboxUpn))`r`nto '$($destination.displayName)' ($($destination.MailboxUpn))."
	if ($report.Failed -gt 0 -or $report.Errors.Count -gt 0) {
		$shownErrors = ($report.Errors | Select-Object -First 5) -join "`r`n"
		$summary += "`r`n`r`n$($report.Failed) item(s) failed:`r`n$shownErrors"
		if ($report.Errors.Count -gt 5) { $summary += "`r`n...and $($report.Errors.Count - 5) more." }
	}
	[System.Windows.Forms.MessageBox]::Show($summary, "Copy Folder Complete")

	# If the destination happens to be the folder on screen, refresh so the new items show up.
	if ($Script:lfFolderID.MailboxId -eq $destination.MailboxId -and $Script:lfFolderID.id -eq $destination.id) {
		Get-SMCClientFolderItems
	}
}

function Invoke-SMCImportFolderButtonClick {
	<# GUI handler for "Import Folder (.fts)..." - imports every *.fts file found directly in a chosen directory into the selected folder. #>
	[CmdletBinding()]
	Param ()
	if (-not $Script:lfFolderID) {
		[System.Windows.Forms.MessageBox]::Show("Select a destination folder in the tree first.")
		return
	}
	$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
		Description = "Choose a directory of exported .fts files to import into '$($Script:lfFolderID.displayName)'"
	}
	if ($folderBrowser.ShowDialog() -ne 'OK') { return }

	$ftsFiles = @(Get-ChildItem -Path $folderBrowser.SelectedPath -Filter "*.fts" -File)
	if ($ftsFiles.Count -eq 0) {
		[System.Windows.Forms.MessageBox]::Show("No .fts files found directly in:`r`n$($folderBrowser.SelectedPath)")
		return
	}

	$session = Invoke-SMCCreateImportSession -MailboxId $Script:lfFolderID.MailboxId
	$imported = 0
	$errors = @()
	$progressForm = New-SMCProgressForm -Title "Import Folder"
	try {
		for ($i = 0; $i -lt $ftsFiles.Count; $i++) {
			$percent = [int]((($i) / $ftsFiles.Count) * 100)
			$statusText = "File $($i + 1) of $($ftsFiles.Count): $($ftsFiles[$i].Name)"
			Write-Progress -Activity "Importing folder" -Status $statusText -PercentComplete $percent
			Update-SMCProgressForm -ProgressForm $progressForm -Status $statusText -PercentComplete $percent
			try {
				[void](Invoke-SMCImportItem -ImportUrl $session.importUrl -FolderId $Script:lfFolderID.id -FilePath $ftsFiles[$i].FullName)
				$imported++
			}
			catch {
				$errors += "$($ftsFiles[$i].Name): $($_.Exception.Message)"
			}
		}
	}
	finally {
		Write-Progress -Activity "Importing folder" -Completed
		Close-SMCProgressForm -ProgressForm $progressForm
	}

	$summary = "Imported $imported of $($ftsFiles.Count) file(s) into '$($Script:lfFolderID.displayName)'."
	if ($errors.Count -gt 0) {
		$shownErrors = ($errors | Select-Object -First 5) -join "`r`n"
		$summary += "`r`n`r`n$($errors.Count) file(s) failed:`r`n$shownErrors"
		if ($errors.Count -gt 5) { $summary += "`r`n...and $($errors.Count - 5) more." }
	}
	[System.Windows.Forms.MessageBox]::Show($summary, "Import Folder Complete")
	Get-SMCClientFolderItems
}

function Show-SMCRawPropertyWindow {
	<#
	Dumps every tagged property the viewer managed to retrieve into a two-column grid - the
	"what did the store actually give me" view, for when a field in the header is blank and the
	question is whether the property is missing or just not being displayed. Binary values are
	shown as a byte count rather than pages of base64.
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [Hashtable]$Props,
		[Parameter(Position = 1, Mandatory = $false)] [String]$Title = "Retrieved MAPI properties"
	)
	Process {
		$propTable = New-Object System.Data.DataTable
		[void]$propTable.Columns.Add("Property")
		[void]$propTable.Columns.Add("Value")
		foreach ($key in ($Props.Keys | Sort-Object)) {
			$value = $Props[$key]
			if ($key -like "Binary *") {
				try { $value = "<binary, $(([Convert]::FromBase64String($value)).Length) bytes>" }
				catch { $value = "<binary>" }
			}
			else {
				$value = "$value"
				if ($value.Length -gt 500) { $value = $value.Substring(0, 500) + " ..." }
			}
			[void]$propTable.Rows.Add($key, $value)
		}

		$propForm = New-Object System.Windows.Forms.Form
		$propForm.Text = $Title
		$propForm.Size = New-Object System.Drawing.Size(820, 520)
		$propForm.StartPosition = "CenterParent"
		$propGrid = New-Object System.Windows.Forms.DataGridView
		$propGrid.Dock = "Fill"
		$propGrid.ReadOnly = $true
		$propGrid.AllowUserToAddRows = $false
		$propGrid.AllowUserToDeleteRows = $false
		$propGrid.AutoSizeColumnsMode = "Fill"
		$propGrid.DataSource = $propTable
		[void]$propForm.Controls.Add($propGrid)
		[void]$propForm.ShowDialog()
		$propForm.Dispose()
	}
}

function Invoke-SMCShowClientMessage {
	<#
	Opens the item viewer for the selected row.

	What was wrong before: the viewer asked for a thin property set and then rendered whatever
	came back into fixed-position labels and a plain TextBox. Three separate problems fell out of
	that. It only ever asked for PR_BODY, so any message whose body is HTML-only showed an empty
	pane. Every field was a hardcoded label at a fixed pixel offset, so long recipient lists and
	subjects were clipped and nothing reflowed when the window was resized. And the header was
	always the mail envelope, so a contact or an appointment displayed a row of blank From/To
	lines instead of its own fields.

	Now: properties are fetched in resilient groups (Get-SMCItemViewerDetail), the header is built
	from whatever is actually populated for that message class (Get-SMCViewerHeaderFields), and
	the body renders PR_HTML through a WebBrowser when present, with PR_BODY as the fallback. A
	toggle switches between rendered HTML and the underlying text, and "Properties" shows the raw
	bag for diagnosis.
	#>
	[CmdletBinding()]
	Param ()
	$row = Get-SMCSelectedGridRow
	if (-not $row) {
		[System.Windows.Forms.MessageBox]::Show("Select a message in the list first.")
		return
	}
	$MessageID = $row[4]
	$rowClass = "$($row[6])"

	try {
		$detail = Get-SMCItemViewerDetail -MailboxId $Script:lfFolderID.MailboxId -FolderId $Script:lfFolderID.id -ItemId $MessageID -MessageClass $rowClass
	}
	catch {
		[System.Windows.Forms.MessageBox]::Show("Couldn't read this item:`r`n`r`n$($_.Exception.Message)", "Show Message Failed")
		return
	}
	# Kept on the script scope so the item can be poked at from the console after viewing.
	$script:msMessage = $detail

	$props = $detail.Props
	$fields = Get-SMCViewerHeaderFields -Props $props -NativeItem $detail.Item

	$plainBody = "$(Get-SMCProp -Bag $props -Key 'String 0x1000')"
	$htmlRaw = Get-SMCProp -Bag $props -Key "Binary 0x1013"
	$htmlBody = ""
	if ($htmlRaw) {
		try {
			$htmlBody = ConvertFrom-SMCHtmlProperty -Base64Data "$htmlRaw" -CodePage (Get-SMCProp -Bag $props -Key "Integer 0x3FDE")
		}
		catch {
			Write-Verbose "PR_HTML present but couldn't be decoded: $($_.Exception.Message)"
		}
	}

	$subject = "$(Get-SMCProp -Bag $props -Key 'String 0x37')"
	if ([String]::IsNullOrWhiteSpace($subject)) { $subject = "(no subject)" }

	$msgform = New-Object System.Windows.Forms.Form
	$msgform.Text = $subject
	$msgform.Size = New-Object System.Drawing.Size(940, 760)
	$msgform.StartPosition = "CenterParent"
	# Same reasoning as the main form: this layout is computed in code, so font-based autoscaling
	# would silently shift every pixel value below.
	$msgform.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::None

	# ---- Header ----
	# One label pair per populated field, stacked at a computed Y. Value labels are Anchored
	# Left+Right so they widen with the window instead of truncating a long recipient list, and
	# the header panel's height is derived from how many fields there actually are.
	$lineHeight = 22
	$headerPanel = New-Object System.Windows.Forms.Panel
	$headerPanel.Dock = "Top"
	$headerPanel.Height = ($fields.Count * $lineHeight) + 16
	$headerPanel.Padding = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)

	$y = 8
	foreach ($field in $fields) {
		$labelCtl = New-Object System.Windows.Forms.Label
		$labelCtl.Location = New-Object System.Drawing.Point(14, $y)
		$labelCtl.Size = New-Object System.Drawing.Size(105, 18)
		$labelCtl.Text = $field.Label
		$labelCtl.Font = New-Object System.Drawing.Font($msgform.Font, [System.Drawing.FontStyle]::Bold)
		[void]$headerPanel.Controls.Add($labelCtl)

		$valueCtl = New-Object System.Windows.Forms.Label
		$valueCtl.Location = New-Object System.Drawing.Point(124, $y)
		$valueCtl.Size = New-Object System.Drawing.Size(($msgform.ClientSize.Width - 140), 18)
		$valueCtl.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
		$valueCtl.AutoEllipsis = $true
		$valueCtl.Text = $field.Value
		# Long values get the full text on hover rather than being lost to the ellipsis.
		$tip = New-Object System.Windows.Forms.ToolTip
		$tip.SetToolTip($valueCtl, $field.Value)
		[void]$headerPanel.Controls.Add($valueCtl)

		$y += $lineHeight
	}

	# ---- Toolbar ----
	$toolPanel = New-Object System.Windows.Forms.Panel
	$toolPanel.Dock = "Top"
	$toolPanel.Height = 34

	$btnToggle = New-Object System.Windows.Forms.Button
	$btnToggle.Location = New-Object System.Drawing.Point(14, 4)
	$btnToggle.Size = New-Object System.Drawing.Size(150, 25)
	$btnToggle.Text = "Show plain text"

	$btnProps = New-Object System.Windows.Forms.Button
	$btnProps.Location = New-Object System.Drawing.Point(174, 4)
	$btnProps.Size = New-Object System.Drawing.Size(150, 25)
	$btnProps.Text = "Properties..."
	[void]$btnProps.Add_Click({ Show-SMCRawPropertyWindow -Props $props -Title "Properties - $subject" })

	$noteLabel = New-Object System.Windows.Forms.Label
	$noteLabel.Location = New-Object System.Drawing.Point(334, 8)
	$noteLabel.Size = New-Object System.Drawing.Size(560, 18)
	$noteLabel.ForeColor = [System.Drawing.Color]::DimGray
	$noteLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

	[void]$toolPanel.Controls.Add($btnToggle)
	[void]$toolPanel.Controls.Add($btnProps)
	[void]$toolPanel.Controls.Add($noteLabel)

	# ---- Body ----
	$bodyPanel = New-Object System.Windows.Forms.Panel
	$bodyPanel.Dock = "Fill"

	$textView = New-Object System.Windows.Forms.TextBox
	$textView.Dock = "Fill"
	$textView.Multiline = $true
	$textView.ScrollBars = "Both"
	$textView.WordWrap = $true
	$textView.ReadOnly = $true
	$textView.Font = New-Object System.Drawing.Font("Consolas", 9)
	$textView.Text = if ($plainBody) { $plainBody -replace "`r?`n", "`r`n" } else { "" }

	# WebBrowser needs an STA host. Creating it inside a try/catch means a non-STA session
	# degrades to the text view rather than throwing in the user's face.
	$htmlView = $null
	if ($htmlBody) {
		try {
			$htmlView = New-Object System.Windows.Forms.WebBrowser
			$htmlView.Dock = "Fill"
			$htmlView.ScriptErrorsSuppressed = $true
			$htmlView.IsWebBrowserContextMenuEnabled = $false
			$htmlView.WebBrowserShortcutsEnabled = $false
			# Stop clicks in the message body from navigating the control away from the item.
			[void]$htmlView.Add_Navigating({
					param($sender, $e)
					if ("$($e.Url)" -notlike "about:*") { $e.Cancel = $true }
				})
			$htmlView.DocumentText = $htmlBody
		}
		catch {
			Write-Verbose "WebBrowser unavailable (needs an STA host): $($_.Exception.Message)"
			$htmlView = $null
		}
	}

	if ($htmlView) {
		[void]$bodyPanel.Controls.Add($htmlView)
		[void]$bodyPanel.Controls.Add($textView)
		$textView.Visible = $false
		$noteLabel.Text = "Rendering PR_HTML. Attachments aren't reachable through this API - export the item for those."
	}
	else {
		[void]$bodyPanel.Controls.Add($textView)
		$btnToggle.Enabled = $false
		if ($plainBody) {
			$noteLabel.Text = "Plain-text body (PR_BODY). This item has no HTML body."
		}
		else {
			$noteLabel.Text = "This item returned neither PR_HTML nor PR_BODY - check Properties for what was retrieved."
			$textView.Text = "(no body content was returned for this item)"
		}
	}

	if ($detail.Failed.Count -gt 0) {
		$noteLabel.Text += "  Some properties couldn't be retrieved."
		$noteLabel.ForeColor = [System.Drawing.Color]::Firebrick
	}

	[void]$btnToggle.Add_Click({
			if (-not $htmlView) { return }
			if ($htmlView.Visible) {
				$htmlView.Visible = $false
				$textView.Visible = $true
				# Fall back to showing the HTML source when there's no plain-text alternative.
				if (-not $plainBody) { $textView.Text = $htmlBody }
				$btnToggle.Text = "Show HTML"
			}
			else {
				$textView.Visible = $false
				$htmlView.Visible = $true
				$btnToggle.Text = "Show plain text"
			}
		})

	# Dock=Fill child must be added before the Dock=Top siblings so it doesn't claim their space.
	[void]$msgform.Controls.Add($bodyPanel)
	[void]$msgform.Controls.Add($toolPanel)
	[void]$msgform.Controls.Add($headerPanel)

	# Activate THIS window on show. The old code called $Script:form.Activate() here, which pulled
	# focus back to the main window the moment the message opened.
	[void]$msgform.Add_Shown({ $msgform.Activate() })
	[void]$msgform.ShowDialog()
	$msgform.Dispose()
}

function Invoke-SMCShowClientHeader {
	<# Internet message headers (PR_TRANSPORT_MESSAGE_HEADERS). Not every item has them. #>
	[CmdletBinding()]
	Param ()
	$row = Get-SMCSelectedGridRow
	if (-not $row) {
		[System.Windows.Forms.MessageBox]::Show("Select a message in the list first.")
		return
	}
	$MessageID = $row[4]
	$Props = @(Get-SMCTaggedProperty -DataType String -Id "0x7D")
	try {
		$item = Invoke-SMCGetItemDetail -MailboxId $Script:lfFolderID.MailboxId -FolderId $Script:lfFolderID.id -ItemId $MessageID -PropertyList $Props
	}
	catch {
		[System.Windows.Forms.MessageBox]::Show("Couldn't read the headers for this item:`r`n`r`n$($_.Exception.Message)", "Show Header Failed")
		return
	}
	$script:msMessage = $item

	$headerText = "$($item.TransportHeaders)"
	if ([String]::IsNullOrWhiteSpace($headerText)) {
		# Internal items (appointments, contacts, anything that never traversed SMTP) simply don't
		# carry transport headers - say so rather than showing an empty box.
		$headerText = "This item has no internet message headers (PR_TRANSPORT_MESSAGE_HEADERS is not set).`r`n`r`nThat's normal for items that never came in over SMTP - calendar items, contacts, tasks, and messages created directly in the mailbox."
	}

	$hdrform = New-Object System.Windows.Forms.Form
	$hdrform.Text = "Internet headers"
	$hdrform.Size = New-Object System.Drawing.Size(820, 620)
	$hdrform.StartPosition = "CenterParent"
	$hdrform.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::None

	$txtHeader = New-Object System.Windows.Forms.TextBox
	$txtHeader.Dock = "Fill"
	$txtHeader.Multiline = $true
	$txtHeader.ScrollBars = "Both"
	$txtHeader.WordWrap = $false
	$txtHeader.ReadOnly = $true
	$txtHeader.Font = New-Object System.Drawing.Font("Consolas", 9)
	$txtHeader.Text = $headerText
	[void]$hdrform.Controls.Add($txtHeader)

	[void]$hdrform.Add_Shown({ $hdrform.Activate() })
	[void]$hdrform.ShowDialog()
	$hdrform.Dispose()
}

function Invoke-SMCSelectClientAttachment {
	[CmdletBinding()]
	Param ()
	$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ Multiselect = $true }
	[void]$FileBrowser.ShowDialog()
	$attname = ""
	foreach ($File in $FileBrowser.FileNames) {
		$script:Attachments += $File
		$attname += $File + " "
	}
	$miMessageAttachmentslableBox1.Text = $attname
}

function Invoke-SMCSendClientMessage {
	[CmdletBinding()]
	Param ()
	# Send from whichever mailbox's folder is currently selected, since more than one mailbox can
	# be open in the tree at once - falls back to the most recently added/opened mailbox if
	# nothing's selected yet.
	$sendingUpn = if ($Script:lfFolderID -and $Script:lfFolderID.MailboxUpn) { $Script:lfFolderID.MailboxUpn } else { $Script:MailboxUpn }
	Send-SMCMessageREST -MailboxUpn $sendingUpn `
		-ToRecipients @(New-SMCEmailAddress -Address $miMessageTotextlabelBox.Text) `
		-Subject $miMessageSubjecttextlabelBox.Text `
		-Body $miMessageBodytextlabelBox.Text `
		-Attachments $script:Attachments
	$script:newmsgform.Close()
}

function Invoke-SMCNewClientMessage {
	[CmdletBinding()]
	Param ()
	$script:newmsgform = New-Object System.Windows.Forms.form
	$script:newmsgform.Text = "New Message"
	$script:newmsgform.size = New-Object System.Drawing.Size(1000, 800)

	$lbl1 = New-Object System.Windows.Forms.Label; $lbl1.Location = New-Object System.Drawing.Size(20, 20); $lbl1.Size = New-Object System.Drawing.Size(80, 20); $lbl1.Text = "To"
	$script:newmsgform.Controls.Add($lbl1)
	$lbl2 = New-Object System.Windows.Forms.Label; $lbl2.Location = New-Object System.Drawing.Size(20, 65); $lbl2.Size = New-Object System.Drawing.Size(80, 20); $lbl2.Text = "Subject"
	$script:newmsgform.Controls.Add($lbl2)

	$miMessageTotextlabelBox = New-Object System.Windows.Forms.TextBox
	$miMessageTotextlabelBox.Location = New-Object System.Drawing.Size(100, 20); $miMessageTotextlabelBox.Size = New-Object System.Drawing.Size(400, 20)
	$script:newmsgform.Controls.Add($miMessageTotextlabelBox)

	$miMessageSubjecttextlabelBox = New-Object System.Windows.Forms.TextBox
	$miMessageSubjecttextlabelBox.Location = New-Object System.Drawing.Size(100, 65); $miMessageSubjecttextlabelBox.Size = New-Object System.Drawing.Size(600, 20)
	$script:newmsgform.Controls.Add($miMessageSubjecttextlabelBox)

	$miMessageBodytextlabelBox = New-Object System.Windows.Forms.RichTextBox
	$miMessageBodytextlabelBox.Location = New-Object System.Drawing.Size(100, 100); $miMessageBodytextlabelBox.Size = New-Object System.Drawing.Size(600, 350)
	$script:newmsgform.Controls.Add($miMessageBodytextlabelBox)

	$lbl3 = New-Object System.Windows.Forms.Label; $lbl3.Location = New-Object System.Drawing.Size(20, 460); $lbl3.Size = New-Object System.Drawing.Size(80, 20); $lbl3.Text = "Attachments"
	$script:newmsgform.Controls.Add($lbl3)
	$miMessageAttachmentslableBox1 = New-Object System.Windows.Forms.Label
	$miMessageAttachmentslableBox1.Location = New-Object System.Drawing.Size(100, 460); $miMessageAttachmentslableBox1.Size = New-Object System.Drawing.Size(600, 20)
	$script:newmsgform.Controls.Add($miMessageAttachmentslableBox1)

	$btnAttach = New-Object System.Windows.Forms.Button
	$btnAttach.Location = New-Object System.Drawing.Size(95, 490); $btnAttach.Size = New-Object System.Drawing.Size(150, 20); $btnAttach.Text = "Add Attachment"
	$btnAttach.Add_Click({ Invoke-SMCSelectClientAttachment })
	$script:newmsgform.Controls.Add($btnAttach)

	$btnSend = New-Object System.Windows.Forms.Button
	$btnSend.Location = New-Object System.Drawing.Size(95, 520); $btnSend.Size = New-Object System.Drawing.Size(125, 20); $btnSend.Text = "Send Message"
	$btnSend.Add_Click({ Invoke-SMCSendClientMessage })
	$script:newmsgform.Controls.Add($btnSend)

	$script:Attachments = @()
	$script:newmsgform.autoscroll = $true
	$script:newmsgform.Add_Shown({ $script:newmsgform.Activate() })
	[void]$script:newmsgform.ShowDialog()
}

function Start-SMCMailClient {
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [String]$MailboxName
	)
	Process {
		$Script:form = New-Object System.Windows.Forms.form
		# Size the form up front, before any Dock/SplitContainer children exist. Dock layout does
		# respond live to later resizes in the normal case, but SplitterDistance in particular was
		# being set against a throwaway placeholder size further down - giving the form its real
		# size first means every control below is laid out against the actual final dimensions on
		# its very first layout pass, nothing is relying on a resize cascade to correct itself.
		$Script:form.Size = New-Object System.Drawing.Size(1200, 800)
		$Script:form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::None

		$Script:mbtable = New-Object System.Data.DataTable
		$mbtable.TableName = "Folder Item"
		[void]$mbtable.Columns.Add("From")
		[void]$mbtable.Columns.Add("Subject")
		[void]$mbtable.Columns.Add("Received", [DATETIME])
		[void]$mbtable.Columns.Add("SizeKB", [INT64])
		[void]$mbtable.Columns.Add("ID")
		[void]$mbtable.Columns.Add("HasAttachments")
		[void]$mbtable.Columns.Add("MessageClass")

		# ---- Layout note ----
		# The previous version placed every control at an absolute (x,y) on the form with
		# AutoScroll = $true. That combination has a classic WinForms failure mode: once the
		# DataGridView gets focus (which it does as soon as it's bound with data), the AutoScroll
		# container scrolls to bring the *focused* control fully into view. On a form that isn't
		# maximized to its full 1200x800 design size, that scroll hides everything positioned
		# before the grid - the mailbox box, Open Mailbox button, and the item-count NumericUpDown
		# included. That's the "item count box" disappearing.
		#
		# Fix: dock the toolbar in a fixed-height Panel at the top (it never scrolls out of view),
		# dock the TreeView to the left and the DataGridView to fill the remaining space inside a
		# second Panel, and turn AutoScroll off on the form entirely - Dock/Anchor layouts don't
		# need it, and it was the thing causing the jump in the first place.

		# ---- Toolbar note ----
		# Nested FlowLayoutPanels (TopDown outer, LeftToRight rows) turned out to be the wrong
		# tool here: FlowDirection=TopDown wraps into new *columns* when it thinks it's short on
		# vertical space, and on a Dock=Top container whose AutoSize height hasn't grown yet at
		# the first layout pass, that's exactly what happened - whole rows got shoved sideways
		# instead of stacking, which is the scrambled/overlapping toolbar you saw.
		#
		# Replacing that with something with zero ambiguity: four plain Panels, each Dock=Top
		# with a fixed Height (Dock=Top stacking has no "wrap" concept at all, so this can't
		# misbehave the way Flow layout did), and within each row, controls are placed left-to-
		# right using a running X cursor computed in code - every control's X is literally
		# "previous control's X + previous control's width + gap", so overlap is not possible.
		#
		# Two more things fixed here based on the last screenshot:
		# 1) Same-edge Dock siblings stack in REVERSE of the order they're added to Controls -
		#    the LAST one added ends up closest to the edge. Row 3 (search) was added last and
		#    landed on top instead of row 1 (mailbox). Controls.Add calls below are reordered
		#    (row4, row3, row2, row1, then body) so the visual stack comes out
		#    row1/row2/row3/row4/body.
		# 2) AutoScaleMode defaults to Font, which rescales every hand-set pixel Size/Location at
		#    runtime based on the system font/DPI vs. the design-time assumption - a bad match
		#    for layout that's computed by hand rather than drawn in the designer, and the likely
		#    reason the first tree/grid row was getting clipped under the toolbar. Turning it off
		#    (done once, up top, before the form has any children) makes every pixel value below
		#    literal, regardless of the machine's DPI/font settings.
		#
		# Row 4 exists because row 2 was already ~1093px of a 1200px form - adding the copy button
		# there would have run it off the right edge on any narrower window. Folder-level
		# operations that aren't per-message live on their own row instead.

		$RowHeight = 36
		$Script:xCursor = 0
		function Add-SMCRowControl {
			param(
				[Parameter(Mandatory = $true)] $Row,
				[Parameter(Mandatory = $true)] $Control,
				[Int]$Gap = 10
			)
			$y = [int](($Row.Height - $Control.Height) / 2)
			$Control.Location = New-Object System.Drawing.Point($Script:xCursor, $y)
			[void]$Row.Controls.Add($Control)
			$Script:xCursor += $Control.Width + $Gap
		}

		$row1 = New-Object System.Windows.Forms.Panel
		$row1.Dock = "Top"; $row1.Height = $RowHeight

		$row2 = New-Object System.Windows.Forms.Panel
		$row2.Dock = "Top"; $row2.Height = $RowHeight

		$row3 = New-Object System.Windows.Forms.Panel
		$row3.Dock = "Top"; $row3.Height = $RowHeight

		$row4 = New-Object System.Windows.Forms.Panel
		$row4.Dock = "Top"; $row4.Height = $RowHeight

		# Reverse add-order: last-added sits closest to the Top edge, so adding row4 first and
		# row1 last makes row1 end up visually topmost.
		[void]$Script:form.Controls.Add($row4)
		[void]$Script:form.Controls.Add($row3)
		[void]$Script:form.Controls.Add($row2)
		[void]$Script:form.Controls.Add($row1)

		$body = New-Object System.Windows.Forms.Panel
		$body.Dock = "Fill"
		# Push the SplitContainer (and so both the TreeView and DataGridView inside it) down with
		# a clear, deliberate gap from the toolbar. Padding on a container is honored by a
		# Dock=Fill child's layout, so this moves both controls down together in one place,
		# rather than juggling another Dock=Top sibling's stacking order. Four rows of 36 = 144,
		# plus a few px of breathing room.
		$body.Padding = New-Object System.Windows.Forms.Padding(0, 148, 0, 0)
		[void]$Script:form.Controls.Add($body)

		# Row 1: mailbox + open + item count
		$Script:xCursor = 8
		$emEmailAddresslableBox = New-Object System.Windows.Forms.Label
		$emEmailAddresslableBox.Size = New-Object System.Drawing.Size(95, 20); $emEmailAddresslableBox.Text = "Mailbox (UPN)"
		Add-SMCRowControl -Row $row1 -Control $emEmailAddresslableBox

		$Script:emEmailAddressTextBox = New-Object System.Windows.Forms.TextBox
		$emEmailAddressTextBox.Size = New-Object System.Drawing.Size(250, 20)
		Add-SMCRowControl -Row $row1 -Control $emEmailAddressTextBox -Gap 20

		$exButton1 = New-Object System.Windows.Forms.Button
		$exButton1.Size = New-Object System.Drawing.Size(110, 25); $exButton1.Text = "Open Mailbox"
		$exButton1.Add_Click({ Invoke-SMCOpenMailbox -Upn $emEmailAddressTextBox.Text })
		Add-SMCRowControl -Row $row1 -Control $exButton1 -Gap 10

		$exButton10 = New-Object System.Windows.Forms.Button
		$exButton10.Size = New-Object System.Drawing.Size(110, 25); $exButton10.Text = "Add Mailbox"
		$exButton10.Add_Click({ Invoke-SMCAddMailboxButtonClick })
		Add-SMCRowControl -Row $row1 -Control $exButton10 -Gap 20

		$saNumItemsBoxLable = New-Object System.Windows.Forms.Label
		$saNumItemsBoxLable.Size = New-Object System.Drawing.Size(65, 20); $saNumItemsBoxLable.Text = "# of Items"
		Add-SMCRowControl -Row $row1 -Control $saNumItemsBoxLable

		$Script:neResultCheckNum = New-Object System.Windows.Forms.NumericUpDown
		$neResultCheckNum.Size = New-Object System.Drawing.Size(70, 20)
		$neResultCheckNum.Minimum = 1; $neResultCheckNum.Maximum = 5000; $neResultCheckNum.Value = 100
		Add-SMCRowControl -Row $row1 -Control $neResultCheckNum

		# Row 2: message actions
		$Script:xCursor = 8
		$exButton2 = New-Object System.Windows.Forms.Button
		$exButton2.Size = New-Object System.Drawing.Size(115, 25); $exButton2.Text = "Show Message"
		$exButton2.Add_Click({ Invoke-SMCShowClientMessage })
		Add-SMCRowControl -Row $row2 -Control $exButton2

		$exButton5 = New-Object System.Windows.Forms.Button
		$exButton5.Size = New-Object System.Drawing.Size(115, 25); $exButton5.Text = "Show Header"
		$exButton5.Add_Click({ Invoke-SMCShowClientHeader })
		Add-SMCRowControl -Row $row2 -Control $exButton5

		$exButton6 = New-Object System.Windows.Forms.Button
		$exButton6.Size = New-Object System.Drawing.Size(115, 25); $exButton6.Text = "New Message"
		$exButton6.Add_Click({ Invoke-SMCNewClientMessage })
		Add-SMCRowControl -Row $row2 -Control $exButton6

		$exButton8 = New-Object System.Windows.Forms.Button
		$exButton8.Size = New-Object System.Drawing.Size(130, 25); $exButton8.Text = "Export Item (.fts)"
		$exButton8.Add_Click({ Invoke-SMCExportMessage })
		Add-SMCRowControl -Row $row2 -Control $exButton8

		$exButton11 = New-Object System.Windows.Forms.Button
		$exButton11.Size = New-Object System.Drawing.Size(150, 25); $exButton11.Text = "Export Folder (.fts)..."
		$exButton11.Add_Click({ Invoke-SMCExportFolderButtonClick })
		Add-SMCRowControl -Row $row2 -Control $exButton11

		$exButton9 = New-Object System.Windows.Forms.Button
		$exButton9.Size = New-Object System.Drawing.Size(140, 25); $exButton9.Text = "Import from FTS..."
		$exButton9.Add_Click({ Invoke-SMCImportFromFts })
		Add-SMCRowControl -Row $row2 -Control $exButton9

		$exButton12 = New-Object System.Windows.Forms.Button
		$exButton12.Size = New-Object System.Drawing.Size(150, 25); $exButton12.Text = "Import Folder (.fts)..."
		$exButton12.Add_Click({ Invoke-SMCImportFolderButtonClick })
		Add-SMCRowControl -Row $row2 -Control $exButton12 -Gap 20

		$exButton7 = New-Object System.Windows.Forms.Button
		$exButton7.Size = New-Object System.Drawing.Size(90, 25); $exButton7.Text = "Update"
		$exButton7.Add_Click({ Get-SMCClientFolderItems })
		Add-SMCRowControl -Row $row2 -Control $exButton7

		# Row 3: client-side search
		$Script:xCursor = 8
		$Script:seSearchCheck = New-Object System.Windows.Forms.CheckBox
		$seSearchCheck.Size = New-Object System.Drawing.Size(20, 20)
		[void]$seSearchCheck.Add_Click({
				$snSearchPropDrop.Enabled = $seSearchCheck.Checked
				$sbSearchTextBox.Enabled = $seSearchCheck.Checked
			})
		Add-SMCRowControl -Row $row3 -Control $seSearchCheck

		$saSeachBoxLable = New-Object System.Windows.Forms.Label
		$saSeachBoxLable.Size = New-Object System.Drawing.Size(210, 20); $saSeachBoxLable.Text = "Search by Property (client-side)"
		Add-SMCRowControl -Row $row3 -Control $saSeachBoxLable -Gap 20

		$Script:snSearchPropDrop = New-Object System.Windows.Forms.ComboBox
		$snSearchPropDrop.Size = New-Object System.Drawing.Size(130, 22)
		[void]$snSearchPropDrop.Items.Add("Subject")
		[void]$snSearchPropDrop.Items.Add("From")
		[void]$snSearchPropDrop.Items.Add("Body")
		$snSearchPropDrop.SelectedIndex = 0
		$snSearchPropDrop.Enabled = $false
		Add-SMCRowControl -Row $row3 -Control $snSearchPropDrop -Gap 20

		$Script:sbSearchTextBox = New-Object System.Windows.Forms.TextBox
		$sbSearchTextBox.Size = New-Object System.Drawing.Size(220, 20)
		$sbSearchTextBox.Enabled = $false
		Add-SMCRowControl -Row $row3 -Control $sbSearchTextBox

		# Row 4: folder-to-folder operations
		$Script:xCursor = 8
		$exButton13 = New-Object System.Windows.Forms.Button
		$exButton13.Size = New-Object System.Drawing.Size(160, 25); $exButton13.Text = "Copy Folder to..."
		$exButton13.Add_Click({ Invoke-SMCCopyFolderButtonClick })
		Add-SMCRowControl -Row $row4 -Control $exButton13 -Gap 20

		$copyHintLabel = New-Object System.Windows.Forms.Label
		$copyHintLabel.Size = New-Object System.Drawing.Size(760, 20)
		$copyHintLabel.Text = "Copies every item from the selected folder into any folder in any open mailbox, 20 items at a time. Source is left untouched."
		$copyHintLabel.ForeColor = [System.Drawing.Color]::DimGray
		Add-SMCRowControl -Row $row4 -Control $copyHintLabel

		$Script:splitContainer = New-Object System.Windows.Forms.SplitContainer
		# Give it a sane starting size before setting SplitterDistance - SplitContainer throws if
		# you set SplitterDistance while its own Width/Height are still 0 (i.e. before it's been
		# sized by its parent's Dock layout). Dock=Fill below then takes over for real sizing.
		$splitContainer.Size = New-Object System.Drawing.Size(1200, 600)
		$splitContainer.Dock = "Fill"
		$splitContainer.Panel1MinSize = 150
		$splitContainer.Panel2MinSize = 300
		$splitContainer.SplitterDistance = 226
		$splitContainer.SplitterWidth = 6
		[void]$body.Controls.Add($splitContainer)

		$Script:tvTreView = New-Object System.Windows.Forms.TreeView
		$tvTreView.Dock = "Fill"
		# Handler takes the sender/event args explicitly rather than relying on the automatic
		# $this/$_ variables PowerShell populates for add_EventName scriptblocks - functionally
		# equivalent, but easier to reason about if item listing ever misfires again.
		[void]$tvTreView.add_AfterSelect({
				param($sender, $e)
				$Script:lfFolderID = $e.Node.Tag
				if (-not $Script:lfFolderID -or -not $Script:lfFolderID.id) {
					Write-Verbose "Selected node '$($e.Node.Text)' has no folder id on its Tag - not fetching items."
					return
				}
				Get-SMCClientFolderItems
			})
		[void]$splitContainer.Panel1.Controls.Add($tvTreView)

		$Script:dgDataGrid = New-Object System.Windows.Forms.DataGridView
		$dgDataGrid.Dock = "Fill"
		$dgDataGrid.AutoSizeRowsMode = "AllHeaders"
		$dgDataGrid.AllowUserToDeleteRows = $false
		$dgDataGrid.AllowUserToAddRows = $false
		# Don't let the grid grab focus on data-bind - that's what was driving the AutoScroll jump.
		$dgDataGrid.TabStop = $false
		# Double-clicking a row opens the viewer, which is what most people try first.
		[void]$dgDataGrid.Add_CellDoubleClick({
				param($sender, $e)
				if ($e.RowIndex -ge 0) { Invoke-SMCShowClientMessage }
			})
		[void]$splitContainer.Panel2.Controls.Add($dgDataGrid)

		$Script:form.Text = "Simple Exchange Mailbox Client (Import/Export API)"
		$Script:form.AutoScroll = $false
		[void]$Script:form.Add_Shown({
				$Script:form.Activate()
				# Re-apply scroll-to-top here rather than only in Invoke-SMCOpenMailbox: setting
				# TopNode before the window has ever been shown (i.e. before the TreeView has a
				# real, final size) can get silently discarded once WinForms actually lays out and
				# paints the control for the first time. Shown is the first point the control's
				# size is guaranteed final, so this is where it actually sticks.
				if ($tvTreView.Nodes.Count -gt 0) { $tvTreView.TopNode = $tvTreView.Nodes[0] }
			})
		$emEmailAddressTextBox.Text = $MailboxName
		Invoke-SMCOpenMailbox -Upn $MailboxName
		[void]$Script:form.ShowDialog()
	}
}

#endregion
