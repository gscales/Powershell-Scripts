# Converting between FolderId formats used in the Exchange Online PowerShell Module and the Graph PowerShell SDK #

One of the harder concepts to get your head around when dealing with Mailbox folders programmatically is the various FolderId formats that are used across different API's and Endpoints in Exchange Online. Here's a list (not complete) of FolderId’s you might see and where they are used

**PR_EntryId**(PidtagEntryId) [https://docs.microsoft.com/en-us/office/client-developer/outlook/mapi/pidtagentryid-canonical-property](https://docs.microsoft.com/en-us/office/client-developer/outlook/mapi/pidtagentryid-canonical-property). Used primarily in MAPI applications and can be used to access a Folder or Item directly. Used in Get-RecoverableItems, Get-PublicFolder

**FolderStoreObjectId** from Get-MailboxFolder and Get-MailboxFolderStatistics this is a padded version of the PidtagEntryId property stored as Base64 this used to be called the owaId in EWS. You can extra the PidtagEntryId by dropping the first and last bytes so if you see

"LgAAAAC+HN09lgYnSJDz3kt9375JAQB1EEf9GOowTZ1AsUKLrCDQAAAAAAEIAAAB"

This is the Base64 value which if you converted it to hex would look like
2E00000000BE1CDD3D9606274890F3DE4B7DDFBE490100751047FD18EA304D9D40B1428BAC20D0000000000108000001

If you drop the first and last byte “2E” and “01” you have the PR_EntryId of the folder

00000000BE1CDD3D9606274890F3DE4B7DDFBE490100751047FD18EA304D9D40B1428BAC20D00000000001080000

**LastParentFolderID** (LapFid) from the Get-RecoverableItems cmdlet To understand what that property Id value represents you need to first understand a little bit more about the Folder EntryId format (PR_EntryId) that exchange uses which is documented in this Exchange Protocol Document https://msdn.microsoft.com/en-us/library/ee217297(v=exchg.80).aspx . A different visual representation of this with the individual components highlighted would look something like this

![](https://4.bp.blogspot.com/-dSlqjHm-FRs/W8guVCIJVGI/AAAAAAAACIU/4NmQ5CZshFA736OIM6-f5vU3OiOJFzjCACLcBGAs/s1600/highlighted.JPG)
 
Hopefully from this you can see that the LAPFID is comprised of the DatabaseGUID and GlobalCounter constituents of the FolderEntryId. To turn this into a hex FolderEntryId (Pr_EntryId) you need to Add the other parts of the Id back in (eg in the Get-RecoverableItems you can do this using the EntryId) eg

    $hexEntyId = $RecoverableItem.EntryID.Substring(0,40) + "0100" + $RecoverableItem.LastParentFolderID + "0000"

So in the above example we are taking the Flags & ProvierUid from the Item (which will be the same as long as the Item is in the same Mailbox as the folder) added the folderType and last padding to reconstitute the PR_EntryId of the folder

**EWSId** (used for any EWS operations and in Mail Addin's) the actual format of this one isn’t documented but its basically made up of the PR_EntryId + the flags that a necessary to access the Mailbox Object (globally). Its Base64 encoded. To convert between the different formats in EWS there is the convertId operation [https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/convertid-operation](https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/convertid-operation) which allows conversion between Hex,Owa,PublicFolder formats.

**Graph (or Rest)id** is the same as the EWSId but it’s Base64URL safe
 
**Graph ImmutableId** [https://docs.microsoft.com/en-us/graph/outlook-immutable-id](https://docs.microsoft.com/en-us/graph/outlook-immutable-id) which is a identifier for an Item/Folder that doesn’t change when the Item is moved between folders which all the above Id’s do (where applicable). This is the one you want to use if you ever are storing Item/FolderId in a external data source.

To convert Id in the Graph there is the Translateexchangeids [https://docs.microsoft.com/en-us/graph/api/user-translateexchangeids?view=graph-rest-1.0&tabs=http](https://docs.microsoft.com/en-us/graph/api/user-translateexchangeids?view=graph-rest-1.0&tabs=http) its more limited then the EWS operation as it doesn’t support translating the owaId or PublicFolderIds (but the Graph can’t access Public folders).

**Mids and Fids** if your using raw ROP’s to access the Exchange Store you need to know about these [https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/22f193f5-f642-4d98-bb2f-5f186397b3ce](https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/22f193f5-f642-4d98-bb2f-5f186397b3ce) for the FolderIds the Fid is I8 in the Exchange Store. 
  
## Other now defunct Id’s ##

**Outlook v1 & v2 RestId’s** are the same as the Graph Id’s
DAV:href This is the old WebDAV itemId this was removed in Exchange 2007 with the demise of the Steaming Store.

## Examples ##

The following examples use this PowerShell Function to covert a Base64String or HexEntryId to the GraphId (Base64SafeURL) and then access the underlying folder using the Graph PowerShell SDK. Note the Invoke-MgGraphRequest  is used with the Powersehll SDK because the Invoke-MgTranslateUserExchangeId doesn't work like a number of the Mail cmdlets because of the way the cmdlets get built from the Open API definitions. 

<script src="https://gist.github.com/gscales/c672746262141e3713ef904b595e8fe3.js"></script>    

Find a usercreated folder called test using Get-MailFolder and then get the folder in the Graph SDK

    $testFolder = get-mailboxfolder -Identity ":\test"
    $testGraphFolderId = Invoke-TranslateId -Base64FolderIdToTranslate $testFolder.FolderStoreObjectId
    Get-MgUserMailFolder -MailFolderId $testGraphFolderId -UserId gscales@userdomain.com

Get a Folder using the Lapfid from Get-RecoverableItems and return that in the Graph SDK

    $lastItem = Get-RecoverableItems -ResultSize 1
    $lastItemsLastActiveFolderId = $lastItem.EntryID.Substring(0,40) + "0100" + $lastItem.LastParentFolderID + "0000"
    $lastItemsLastActiveFolderGraphrId = Invoke-TranslateId -HexEntryId $lastItemsLastActiveFolderId
    Get-MgUserMailFolder -MailFolderId $lastItemsLastActiveFolderGraphrId -UserId gscales@userdomain.com





