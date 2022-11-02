## Migrating EWS Powershell Scripts to the Microsoft Graph PowerShell SDK Evergreen Tips Mailbox Folders ##

**Scope**

The purpose of this document is to create an evergreen doc for tips for migrating from EWS based powershell scripts to the Microsoft Graph Powershell SDK based on real world experience when doing a migration. It won't cover the specifics around connection and authentication which are already well documented but the deeper level technical problems around doing specific activities. 

**Conventions**

For this document I'm going to use MGPSDK as shorthand for the Microsoft Graph Powershell SDK

**Getting a Folder**

When you want to get a users Mailbox folder there are a couple of methods you can use based on what type of Mailbox folder your trying to get.

**WellKnown/System/Default Folders** (People refer to them as different things)

For default mailbox folders like the Inbox, SentItems etc you can use the WellKnownFoldername which is a language independent enumeration (because depending on the language being used by a Mailbox the Inbox will be named differently).For example in the MGPSDK you could use the following to get the Inbox by its WellknownFolderName

    Get-MgUserMailFolder -UserId $MailboxName -MailFolderId Inbox 

Other WellKnownFolderName you can use for other folders are

- MsgFolderRoot (Root of the Mailbox)
- Root (Non IPM Mailbox Root)
- SentItems (SentItems folder)
- Drafts (Drafts Folder)
- RecoverableItemsDeletions (Root of the Dumpster)
- RecoverableItemsPurges (Dumpster Purges Folder)

There is a list of all the EWS WellknownFolderNames here https://learn.microsoft.com/en-us/dotnet/api/microsoft.exchange.webservices.data.wellknownfoldername?view=exchange-ews-api

**Note** If you want to access any folder in the Archive Store this is not currently possible using the Graph API 

**User Created Folders**

For Mailbox folders that the user has created you must know the FolderId of the folder before you can retrieve it (folderid's are not immutable so if a folder moves in a mailbox it will change). To Get the folderid you need to either search for that folder or have it stored or supplied from another source (eg ExoV2 cmdlets etc).

**Searching**

There are a couple of different approaches to searching for a user created folder, for example you could just use a search based on the DisplayName of the folder, the weakness of that method is that you can have a folder with the same name with different ParentFolders. My preferred method is to search using a Path that similar to a directory path eg \Inbox\subFolder1\Subfolder2. The weakness of this method is for mailboxes that have a different language this will fail because Inbox will be in the Regionalized format. So I've come up with the following to allow search by path but also allows changing the Search Root.

    function Get-MailBoxFolderFromPath {
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
        $WellKnownSearchRoot = "MsgFolderRoot"
    )

    process {
        if($FolderPath -eq '\'){
            return Get-MgUserMailFolder -UserId $MailboxName -MailFolderId msgFolderRoot 
        }
        $fldArray = $FolderPath.Split("\")
        #Loop through the Split Array and do a Search for each level of folder 
        $folderId = $WellKnownSearchRoot
        for ($lint = 1; $lint -lt $fldArray.Length; $lint++) {
            #Perform search based on the displayname of each folder level
            $FolderName = $fldArray[$lint];
            $tfTargetFolder = Get-MgUserMailFolderChildFolder -UserId $MailboxName -Filter "DisplayName eq '$FolderName'" -MailFolderId $folderId -All 
            if ($tfTargetFolder.displayname -eq $FolderName) {
                $folderId = $tfTargetFolder.Id.ToString()
            }
            else {
                throw ("Folder Not found")
            }
        }
        return $tfTargetFolder 
        }
    }

So using the above Function if you where searching for a user folder of the Mailbox Root you could use

     Get-MailBoxFolderFromPath -MailboxName mailbox@contoso.com -FolderPath "\userFolder"

If it was a Subfolder on the Inbox eg where it full path would be "\Inbox\userFolder" but where you have regional mailboxes in different languages you could use

     Get-MailBoxFolderFromPath -MailboxName mailbox@contoso.com -FolderPath "\userFolder" -WellKnownSearchRoot inbox

Using Id's from ExoV3 Powershell Cmdlets

If you are using the Exo Powershell cmdlets and need to extend what you are doing in those cmdlets with further things in the Microsoft Graph you can use the Id from ExoV3 cmdlet and convert that into a usable Id in the Graph for example

The output of Exo Get-MailboxFolderStatistics cmdlet mailbox@contoso.com may look something like this

eg

    RunspaceId: 529aa9ab-1403-49d6-94d9-ee0c7f1625fc
    Date  : 23/07/2021 12:13:27 AM
    CreationTime  : 23/07/2021 12:13:27 AM
    LastModifiedTime  : 29/05/2022 8:13:38 PM
    Name  : test123
    FolderPath: /Inbox/test123
    FolderId  : LgAAAAC+HN09lgYnSJDz3kt9375JAQB1EEf9GOowTZ1AsUKLrCDQAAYPugioAAAB
    ParentFolderId: LgAAAAC+HN09lgYnSJDz3kt9375JAQB1EEf9GOowTZ1AsUKLrCDQAAAAAAEMAAAB

The Id format use in return is what used to be called as the OWAId in EWS or in the Exo cmdlets is called the FolderStoreObjectId which is a padded version of the PidtagEntryId property stored as Base64. You can extra the PidtagEntryId by dropping the first and last bytes so if you see

"LgAAAAC+HN09lgYnSJDz3kt9375JAQB1EEf9GOowTZ1AsUKLrCDQAAAAAAEIAAAB"

This is the Base64 value which if you converted it to hex would look like 2E00000000BE1CDD3D9606274890F3DE4B7DDFBE490100751047FD18EA304D9D40B1428BAC20D0000000000108000001

If you drop the first and last byte “2E” and “01” you have the PR_EntryId of the folder

00000000BE1CDD3D9606274890F3DE4B7DDFBE490100751047FD18EA304D9D40B1428BAC20D00000000001080000

You can then use a function like the following to convert it to GraphId

    function Invoke-TranslateId {
    [CmdletBinding()] 
    param( 
        [Parameter(Position = 1, Mandatory = $false)]
        [String]
        $Base64FolderIdToTranslate = "",
        [Parameter(Position = 1, Mandatory = $false)]
        [String]
        $HexEntryId = ""

    )  
    Process {
        if (![String]::IsNullOrEmpty($Base64FolderIdToTranslate)) {
            $HexEntryId = [System.BitConverter]::ToString([Convert]::FromBase64String($Base64FolderIdToTranslate)).Replace("-", "").Substring(2)  
            $HexEntryId = $HexEntryId.SubString(0, ($HexEntryId.Length - 2))
        }
        $FolderIdBytes = [byte[]]::new($HexEntryId.Length / 2)
        For ($i = 0; $i -lt $HexEntryId.Length; $i += 2) {
            $FolderIdBytes[$i / 2] = [convert]::ToByte($HexEntryId.Substring($i, 2), 16)
        }

        $FolderIdToConvert = [System.Web.HttpServerUtility]::UrlTokenEncode($FolderIdBytes)

        $ConvertRequest = @"
	{
  	"inputIds" : [
   	 "$FolderIdToConvert"
  	],
  	"sourceIdType": "entryId",
  	"targetIdType": "restId"
	}
	"@

        $ConvertResult = Invoke-MgGraphRequest -Method POST -Uri https://graph.microsoft.com/v1.0/me/translateExchangeIds -Body $ConvertRequest
        return $ConvertResult.Value.targetId
    }
	}

eg to get a Folder from Get-MailboxStatistics output

    $folder = Get-MailboxFolderStatistics mailbox@contso.com | where-object {$_.FolderPath -eq '/Inbox/test'}
    $restId = Invoke-TranslateId -Base64FolderIdToTranslate $folder.FolderId
    Get-MgUserMailFolder -UserId mailbox@contso.com -MailFolderId $restid

**Enumerating all Mail Folders in a Folder Tree**

In EWS there was the ability when using a GetFolder request to specify the Folder Traversal as either Shallow or Deep. A Shallow traversal would retrieve the folders with a child depth of 1 and Deep would return all the folders in a Folder Tree (or where you started the search). This meant with one operation you could get the entire folder tree, in the Graph this isn't possible as all retrievals are Shallow (a couple of exceptions are Delta ops, or if you use Expand on Child-folders but this only does 1 level in this case). So to get all the Mail-folders in a Tree it requires that you traverse each of the childFolders that have children of their own.

For getting all the folders in a Folder Tree I have the following script [https://github.com/gscales/Powershell-Scripts/blob/master/Graph101/GraphSDK/EnumerateMailboxFolders.ps1](https://github.com/gscales/Powershell-Scripts/blob/master/Graph101/GraphSDK/EnumerateMailboxFolders.ps1)

To Get all the Child Folders in the Inbox and include the Inbox in the returned collection use

    $FolderCollection = Invoke-EnumerateMailBoxFolders -MailboxName mailbox@contso.com -WellKnownFolder inbox -returnSearchRoot

To Get all folders in a Mailbox (don't include the root in the results)

    $FolderCollection = Invoke-EnumerateMailBoxFolders -MailboxName mailbox@contso.com -WellKnownFolder msgfolderRoot

**Example script** 

Create a csv report of all the Folders in the Inbox that have a zero item count

    Invoke-EnumerateMailBoxFolders -WellKnownFolder inbox -MailboxName mailbox@contso.com | Where-Object {$_.TotalItemCount -eq 0} | select-object FolderPath,DisplayName,TotalItemCount | export-csv -NoTypeInformation -Path "c:\temp\emptyFolders.csv"