<strong>Get-Contact</strong>&nbsp; <br />
<br />
This can be used to get a Contact from the Mailbox's default&nbsp;contacts folder, other contacts subfolder or the Global Address List eg to get a contact from the default contact folder by searching using the Email Address (This will return a EWS Managed API Contact object).<br />
<br />
Example 1 <br />
<br />
Get-Contact -MailboxName <a href="mailto:mailbox@domain.com">mailbox@domain.com</a> -EmailAddress <a href="mailto:contact@email.com">contact@email.com</a><br />
<br />
This will search the default contacts folder using the ResolveName operation in EWS, it also caters for contacts that where added from the Global Address List in Outlook. When you add a contact from the GAL the email address that is stored in the Mailbox's contacts Folder uses the EX Address. So in this case when you go to resolve or search on&nbsp;the SMTP address it won't find the contact that has been added from the GAL with this address. To cater for this the&nbsp;GAL is also searched for the EmailAddress you enter in (using ResolveName), if a GAL entry is returned (that has a matching EmailAddress)&nbsp;then the EX Address is obtained using Autodiscover and the UserDN property and then another Resolve is done against the Contacts Folder using the EX address.<br />
<br />
Because ResolveName allows you to resolve against more then just the Email address I've added a -Partial Switch so you can also do partial match searches. Eg to return all the contacts that contain a particular word (note this could be across all the properties that are searched) you can use<br />
<br />
Get-Contact -MailboxName <a href="mailto:mailbox@domain.com">mailbox@domain.com</a> -EmailAddress glen -Partial<br />
<br />
By default only the Primary Email of&nbsp;a contact&nbsp;is checked when you using ResolveName if you want it to search the multivalued Proxyaddressses property you need to use something like the following<br />
<br />
Get-Contact -MailboxName&nbsp; <a href="mailto:mailbox@domain.com">mailbox@domain.com</a>&nbsp;-EmailAddress <a href="mailto:info@domain.com">smtp:info@domain.com</a>&nbsp;-Partial<br />
<br />
Or to search via the SIP address you can use<br />
<br />
Get-Contact -MailboxName&nbsp; <a href="mailto:mailbox@domain.com">mailbox@domain.com</a>&nbsp;-EmailAddress <a href="mailto:info@domain.com">sip:info@domain.com</a>&nbsp;-Partial<br />
<br />
(using the Partial switch is required in this case because the EmailAddress your search on won't match the PrimaryAddress of the contact so in this case also you can get partial matches back).<br />
<br />
There is also a&nbsp;-SearchGal switch for this cmdlet which means only the GAL is searched eg<br />
<br />
Get-Contact -MailboxName <a href="mailto:mailbox@domain.com">mailbox@domain.com</a> -EmailAddress <a href="mailto:gscales@domain.com">gscales@domain.com</a> -SearchGal<br />
<br />
In this case the contact object returned will be read only (although you can save it into a contacts folder which I've used in another cmdlet).<br />
<br />
Finally if your contacts aren't located in the default contacts folder you can&nbsp;use the folder parameter to enter in the path to folder that you want&nbsp;to search&nbsp;eg<br />
<br />
Get-Contact -MailboxName <a href="mailto:mailbox@domain.com">mailbox@domain.com</a> -EmailAddress <a href="mailto:gscales@domain.com">gscales@domain.com</a> -Folder "\Contacts\SubFolder"<br />
<br><br><strong>Get-Contacts <br><br></strong>This can be used to get all the 
contacts from a contacts folder in a Mailbox
<p>Example 1 To get a Contact from a Mailbox's default contacts folder<br>
Get-Contacts -MailboxName mailbox@domain.com <br><br>Example 2 To get all the 
Contacts from subfolder of the Mailbox's default contacts folder<br>Get-Contacts 
-MailboxName mailbox@domain.com -Folder \Contact\test</p>
&nbsp;<strong><br>Create-Contact</strong><br />
<br />
This can be used to create a contact, I've added parameters for all the most common properties you would set in a contact but I haven't added any Extended properties (if you need to set this you can either add it in yourself or after you create the contact use Get-Contact and update the contact object).<br />
<br />
Example 1&nbsp;to create a contact in the default contacts folder <br />
<br />
Create-Contact -Mailboxname <a href="mailto:mailbox@domain.com">mailbox@domain.com</a> -EmailAddress <a href="mailto:contactEmai@domain.com">contactEmai@domain.com</a> -FirstName John -LastName Doe -DisplayName "John Doe"<br />
<br />
to create a contact and add a contact picture<br />
<br />
Create-Contact -Mailboxname <a href="mailto:mailbox@domain.com">mailbox@domain.com</a> -EmailAddress <a href="mailto:contactEmai@domain.com">contactEmai@domain.com</a> -FirstName John -LastName Doe -DisplayName "John Doe" -photo 'c:\photo\Jdoe.jpg'<br />
<br />
to create a contact in&nbsp;a user created subfolder <br />
<br />
&nbsp;Create-Contact -Mailboxname <a href="mailto:mailbox@domain.com">mailbox@domain.com</a> -EmailAddress <a href="mailto:contactEmai@domain.com">contactEmai@domain.com</a> -FirstName John -LastName Doe -DisplayName "John Doe" -Folder "\MyCustomContacts"<br />
<br />
This cmdlet uses the EmailAddress as unique key so it wont let you create a contact with that email address if one already exists.<br />
<br />
<strong>Update-Contact</strong><br />
<strong></strong><br />
This Cmdlet can be used to update an existing contact the Primary email address is used as a unique key so this is the one property you can't update. It will take the Partial switch like the other cmdlet but will always prompt before updating in this case.<br />
<br />
Example1 update the phone number of an existing contact<br />
<br />
Update-Contact&nbsp; -Mailboxname <a href="mailto:mailbox@domain.com">mailbox@domain.com</a> -EmailAddress <a href="mailto:contactEmai@domain.com">contactEmai@domain.com</a>&nbsp;-MobilePhone 023213421 <br />
<br />
Example 2 update the phone number of a contact in a users subfolder<br />
<br />
Update-Contact&nbsp; -Mailboxname <a href="mailto:mailbox@domain.com">mailbox@domain.com</a> -EmailAddress <a href="mailto:contactEmai@domain.com">contactEmai@domain.com</a>&nbsp;-MobilePhone 023213421 -Folder "\MyCustomContacts"<br />
<strong></strong><br />
<strong>Delete-Contact</strong><br />
<strong></strong><br />
This Cmdlet can be used to delete a contact from a contact folders<br />
<br />
eg to delete a contact from the default contacts folder<br />
<br />
Delete-Contact -MailboxName <a href="mailto:mailbox@domain.com">mailbox@domain.com</a> -EmailAddress <a href="mailto:email@domain.com">email@domain.com</a> <br />
<br />
to delete a contact from&nbsp;a non&nbsp;user subfolder<br />
<br />
Delete-Contact -MailboxName <a href="mailto:mailbox@domain.com">mailbox@domain.com</a> -EmailAddress <a href="mailto:email@domain.com">email@domain.com</a> -Folder \Contacts\Subfolder<br />
<br />
<strong>Export-Contact</strong><br />
<strong></strong><br />
This cmdlet can be used to export a contact to a VCF file, this takes advantage of EWS ability to provide the contact as a VCF stream via the MimeContent property.<br />
<br />
To export a Contact to a vcf use<br />
<br />
Export-Contact -MailboxName <a href="mailto:mailbox@domain.com">mailbox@domain.com</a> -EmailAddress <a href="mailto:address@domain.com">address@domain.com</a> -FileName c:\export\filename.vcf<br />
<br />
If the file already exists it will handle creating a unique filename<br />
<br />
To export from a contacts subfolder use<br />
<br />
Export-Contact -MailboxName <a href="mailto:mailbox@domain.com">mailbox@domain.com</a> -EmailAddress <a href="mailto:address@domain.com">address@domain.com</a> -FileName c:\export\filename.vcf -folder \contacts\subfolder<br />
<br />
<strong>Export-GALContact</strong><br />
<strong></strong><br />
This cmdlet exports a Global Address List entry to a VCF file, unlike the Export-Contact cmdlet which can take advantage of the MimeStream provided by the Exchange Store with GAL Contact you don't have this available. The script creates aVCF file manually using the properties returned from ResolveName. By default the GAL photo is included with the file but I have included a -IncludePhoto switch which will use the GetUserPhoto operation which is only available on 2013 and greater. <br />
<br />
Example 1 to save a GAL Entry to a vcf <br />
<br />
Export-GalContact -MailboxName <a href="mailto:user@domain.com">user@domain.com</a> -EmailAddress <a href="mailto:email@domain.com">email@domain.com</a> -FileName c:\export\export.vcf<br />
<br />
Example 2 to save a GAL Entry to vcf including the users photo<br />
<br />
Export-GalContact -MailboxName <a href="mailto:user@domain.com">user@domain.com</a> -EmailAddress <a href="mailto:email@domain.com">email@domain.com</a> -FileName c:\export\export.vcf -IncludePhoto<br />
<br />
<strong>Copy-Contacts.GalToMailbox</strong><br />
<strong></strong><br />
This Cmdlet copies a contact from the Global Address list to a local contacts folder. The EmailAddress in used as a unique key so the same contact won't be copied into a local contacts folder if it already exists. By default the GAL photo is included with the file but I have included a -IncludePhoto switch which will use the GetUserPhoto operation which is only available on 2013 and greater.<br />
<br />
Example 1 to Copy a Gal contacts to&nbsp;local Contacts folder<br />
<br />
Copy-Contacts.GalToMailbox -MailboxName&nbsp;<a href="mailto:mailbox@domain.com">mailbox@domain.com</a> -EmailAddress&nbsp;<a href="mailto:email@domain.com">email@domain.com</a>&nbsp;<br />
<br />
Example 2&nbsp;Copy a GAL contact to a Contacts subfolder<br />
<br />
Copy-Contacts.GalToMailbox -MailboxName&nbsp;<a href="mailto:mailbox@domain.com">mailbox@domain.com</a> -EmailAddress&nbsp;<a href="mailto:email@domain.com">email@domain.com</a>&nbsp;&nbsp;-Folder \Contacts\UnderContacts<br />
<br />
<p><strong>Get-ContactGroup </strong></p>
<p>This Cmdlet can be used to get a ContactGroup from a Mailbox</p>
<p>Example 1 To Get a Contact Group in the default contacts folder <br><br>
Get-ContactGroup -Mailboxname mailbox@domain.com -GroupName GroupName <br><br>
Example 2 To Get a Contact Group in a subfolder of default contacts folder <br>
<br>Get-ContactGroup -Mailboxname mailbox@domain.com -GroupName GroupName 
-Folder \Contacts\Folder1 </p>
<p><strong>Create-ContactGroup</strong></p>
<p>Example 1 To create a Contact Group in the default contacts folder <br><br>
Create-ContactGroup -Mailboxname mailbox@domain.com -GroupName GroupName 
-Members (&quot;member1@domain.com&quot;,&quot;member2@domain.com&quot;)<br><br>Example 2 To create 
a Contact Group in a subfolder of default contacts folder <br><br>
Create-ContactGroup -Mailboxname mailbox@domain.com -GroupName GroupName -Folder 
\Contacts\Folder1 -Members (&quot;member1@domain.com&quot;,&quot;member2@domain.com&quot;)</p>

