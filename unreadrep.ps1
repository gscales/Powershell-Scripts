Import-Module .\unReadModule.ps1 -Force
$Credentials = get-credential
$report = @()
$Mailboxes = Get-Mailbox -ResultSize Unlimited
foreach($Mailbox in $Mailboxes){
	Write-host ("Processing Mailbox " + $Mailbox.PrimarySMTPAddress.ToString())
	$report += Get-UnReadMessageCount -MailboxName $Mailbox.PrimarySMTPAddress.ToString() -Credentials $Credentials -Months 6 -useImpersonation -url https://outlook.office365.com/EWS/Exchange.asmx
}
$report | Export-Csv -path c:\Temp\Unreadreport.csv -NoTypeInformation 