[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 

$form1 = new-object System.Windows.Forms.form 
$form1.Text = "Exchange 2010 Remote PS Select"
$form1.size = new-object System.Drawing.Size(400,200) 
$ServerName = ""
$conuri = ""
# Add Server DropLable
$snServerNamelableBox1 = new-object System.Windows.Forms.Label
$snServerNamelableBox1.Location = new-object System.Drawing.Size(10,20) 
$snServerNamelableBox1.size = new-object System.Drawing.Size(120,20) 
$snServerNamelableBox1.Text = "Remote ServerName"
$form1.Controls.Add($snServerNamelableBox1) 

$exButton4 = new-object System.Windows.Forms.Button
$exButton4.Location = new-object System.Drawing.Size(10,80)
$exButton4.Size = new-object System.Drawing.Size(125,20)
$exButton4.Text = "Connect"
$exButton4.Add_Click({
                          $conuri = "http://" + $ServerName + "/PowerShell/"
                          $form1.close()
                    })
$form1.Controls.Add($exButton4)


# Add Server Drop Down
$snServerNameDrop1 = new-object System.Windows.Forms.ComboBox
$snServerNameDrop1.Location = new-object System.Drawing.Size(150,20)
$snServerNameDrop1.Size = new-object System.Drawing.Size(160,30)
$snServerNameDrop1.Add_SelectedValueChanged({$ServerName = $snServerNameDrop1.SelectedItem.ToString()})

$form1.Controls.Add($snServerNameDrop1)

$root = [ADSI]'LDAP://RootDSE'
$cfConfigRootpath = "LDAP://" + $root.ConfigurationNamingContext.tostring()
$configRoot = [ADSI]$cfConfigRootpath
$searcher = new-object System.DirectoryServices.DirectorySearcher($configRoot)
$searcher.Filter = '(objectCategory=msExchExchangeServer)'
[VOID]$searcher.PropertiesToLoad.Add("cn")
$searchres = $searcher.FindAll()
foreach ($res in $searchres){
    $srvOjb = $res.Properties
    $snServerNameDrop1.Items.Add($srvOjb.cn[0])
}

$form1.autoscroll = $true
$form1.Add_Shown({$form1.Activate()})
$form1.ShowDialog()

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $conuri -Authentication Kerberos 
Import-PSSession $Session 


$HistoryDir = "c:\mbsizehistory"

if (!(Test-Path -path $HistoryDir))
{
	New-Item $HistoryDir -type directory
}

function CovertBitValue($String){
	$numItempattern = '(?=\().*(?=bytes)'
	$matchedItemsNumber = [regex]::matches($String, $numItempattern) 
	$Mb = [INT64]$matchedItemsNumber[0].Value.Replace("(","").Replace(",","")
	return [math]::round($Mb/1048576,0)
}

function getMailboxSizes(){
$datetime = get-date
$fname = $script:HistoryDir + "\"
$mbcombCollection = @()
$mscombCollection = @()
$mstoresquotas.clear()
$msTable.clear()
$usrquotas.clear()
if ($snServerNameDrop.SelectedItem.ToString() -eq "ALL Servers"){
	$dbsetting = get-mailboxdatabase 
}
else{
	$dbsetting = get-mailboxdatabase -server $snServerNameDrop.SelectedItem.ToString()
} 
$dbsetting | ForEach-Object{
	$_.identity
	$_.ProhibitSendReceiveQuota
	if ($qtTypeDrop.SelectedItem -eq $null){
		if ($_.ProhibitSendReceiveQuota.IsUnlimited -ne $true){
			$mstoresquotas.add($_.identity,$_.ProhibitSendReceiveQuota)
		}
	}
	else{
		$soStoreObject = $_
		switch ($qtTypeDrop.SelectedItem.ToString()){
			"Warning" {
					if ($soStoreObject.IssueWarningQuota.IsUnlimited -ne $true){
						$mstoresquotas.add($soStoreObject.identity,$soStoreObject.IssueWarningQuota)
					}
				}
			"Proibit Send" {
					if ($soStoreObject.ProhibitSendQuota.IsUnlimited -ne $true){
						$mstoresquotas.add($soStoreObject.identity,$soStoreObject.ProhibitSendQuota)
					}
				}
			"Proibit Send/Recieve" {
				if ($soStoreObject.ProhibitSendReceiveQuota.IsUnlimited -ne $true){
					$mstoresquotas.add($soStoreObject.identity,$soStoreObject.ProhibitSendReceiveQuota)
				}
			}
		}
	
	}

	
}

$usrquotas = @{ }
if ($snServerNameDrop.SelectedItem.ToString() -eq "ALL Servers"){
	$mailboxes = get-mailbox -ResultSize Unlimited
}
else{
	$mailboxes = get-mailbox -server $snServerNameDrop.SelectedItem.ToString() -ResultSize Unlimited
} 
$mailboxes | foreach-object{
	if ($qtTypeDrop.SelectedItem -eq $null){
		if($_.ProhibitSendReceiveQuota.IsUnlimited -ne $true){
			$usrquotas.add($_.ExchangeGuid,$_.ProhibitSendReceiveQuota)
		}
	}
	else {
		$uoUserobject = $_
		switch ($qtTypeDrop.SelectedItem.ToString()){
			"Warning" {
				if($uoUserobject.IssueWarningQuota.isUnlimited -eq $false){
					$usrquotas.add($uoUserobject.ExchangeGuid,$uoUserobject.IssueWarningQuota)
				}
				}
			"Proibit Send" {
				if($uoUserobject.ProhibitSendQuota.isUnlimited -eq $false){
					$usrquotas.add($uoUserobject.ExchangeGuid,$uoUserobject.ProhibitSendQuota)
				}
				}
			"Proibit Send/Recieve" {
				if($uoUserobject.ProhibitSendReceiveQuota.isUnlimited -eq $false){
					$usrquotas.add($uoUserobject.ExchangeGuid,$uoUserobject.ProhibitSendReceiveQuota)
				}
				}
		}
	}
}

$mbServers = get-mailboxserver

if ($mtTypeDrop.SelectedItem -ne $null){
	if ($mtTypeDrop.SelectedItem.ToString() -eq "Disconnected"){
		$fname ="disconnected"
		if ($snServerNameDrop.SelectedItem.ToString() -eq "ALL Servers"){
			$mbServers | foreach-object{
				$mscombCollection += get-mailboxstatistics -server $_.Name  | Where {$_.DisconnectDate -ne $null}
			}
		}
		else{
			$mscombCollection += get-mailboxstatistics -server $snServerNameDrop.SelectedItem.ToString()  | Where {$_.DisconnectDate -ne $null}
		} 
		$mscombCollection | ForEach-Object{
		$quQuota = "0"
		if ($usrquotas.ContainsKey($_.MailboxGUID)){
			if ($usrquotas[$_.MailboxGUID].Value -ne $null){
				if (CovertBitValue($usrquotas[$_.MailboxGUID].Value) -gt 0){
					$quQuota = "{0:P0}" -f (CovertBitValue($_.TotalItemSize)/CovertBitValue($usrquotas[$_.MailboxGUID]))}
				else{
					$quQuota = "100"
				}
			}
		}
		else{
			if ($mstoresquotas.ContainsKey($_.database)){
				if ($mstoresquotas[$_.database].Value -ne $null){
					$quQuota = "{0:P0}" -f (CovertBitValue($_.TotalItemSize)/CovertBitValue($mstoresquotas[$_.database]))
			}}
		}
		$icount = 0
		$tisize = 0
		$disize = 0
		if ($_.DisplayName -ne $null){$dname = $_.DisplayName}
		if ($_.ItemCount -ne $null){$icount = $_.ItemCount}
		if (CovertBitValue($_.TotalItemSize) -ne $null){$tisize = CovertBitValue($_.TotalItemSize)}
		if (CovertBitValue($_.TotalDeletedItemSize) -ne $null){$disize = CovertBitValue($_.TotalDeletedItemSize)}  
		$msTable.Rows.add($dname,$icount,$tisize,$disize,$quQuota.replace("%","").replace(",",""))
		$mbcomb = "" | select Date,ServerName,DisplayName,ItemCount,TotalItemSize,TotalDeletedItemSize
		$mbcomb.Date = $datetime.ToString("yyyyMMdd")
		$mbcomb.ServerName = $snServerNameDrop.SelectedItem.ToString()
		$mbcomb.DisplayName = $dname
		$mbcomb.ItemCount = $icount
		$mbcomb.TotalItemSize = $tisize
		$mbcomb.TotalDeletedItemSize = $disize
		$mbcombCollection += $mbcomb
		}
	}
	else{	
		if ($snServerNameDrop.SelectedItem.ToString() -eq "ALL Servers"){
			$fname = $fname  + $datetime.ToString("yyyyMMdd")  + "-ALLSERVERS.csv"
			$mbServers | foreach-object{
				$mscombCollection += get-mailboxstatistics -server $_.Name | Where {$_.DisconnectDate -eq $null}
			}
		}
		else{
			$fname = $fname  + $datetime.ToString("yyyyMMdd")  + "-" + $snServerNameDrop.SelectedItem.ToString() + ".csv"
			$mscombCollection += get-mailboxstatistics -server $snServerNameDrop.SelectedItem.ToString()  | Where {$_.DisconnectDate -eq $null}
		} 
		$mscombCollection  | ForEach-Object{
		$quQuota = "0"
		if ($usrquotas.ContainsKey($_.MailboxGUID)){
			if ($usrquotas[$_.MailboxGUID].Value -ne $null){
				if (CovertBitValue($usrquotas[$_.MailboxGUID]) -gt 0){
					$quQuota = "{0:P0}" -f (CovertBitValue($_.TotalItemSize)/CovertBitValue($usrquotas[$_.MailboxGUID]))}
				else{
					$quQuota = "100"
				}
			}
		}
		else{
			if ($mstoresquotas.ContainsKey($_.database)){
				if ($mstoresquotas[$_.database].Value -ne $null){
				$quQuota = "{0:P0}" -f (CovertBitValue($_.TotalItemSize)/CovertBitValue($mstoresquotas[$_.database]))}}
		}
		$icount = 0
		$tisize = 0
		$disize = 0
		if ($_.DisplayName -ne $null){$dname = $_.DisplayName}
		if ($_.ItemCount -ne $null){$icount = $_.ItemCount}
		if (CovertBitValue($_.TotalItemSize) -ne $null){$tisize = CovertBitValue($_.TotalItemSize)}
		if (CovertBitValue($_.TotalDeletedItemSize) -ne $null){$disize = CovertBitValue($_.TotalDeletedItemSize)}    
		$msTable.Rows.add($dname,$icount,$tisize,$disize,$quQuota.replace("%","").replace(",",""))
		$mbcomb = "" | select Date,ServerName,DisplayName,ItemCount,TotalItemSize,TotalDeletedItemSize
		$mbcomb.Date = $datetime.ToString("yyyyMMdd")
		$mbcomb.ServerName = $snServerNameDrop.SelectedItem.ToString()
		$mbcomb.DisplayName = $dname
		$mbcomb.ItemCount = $icount
		$mbcomb.TotalItemSize = $tisize
		$mbcomb.TotalDeletedItemSize = $disize
		$mbcombCollection += $mbcomb
		}

	}
}
else{
		
		if ($snServerNameDrop.SelectedItem.ToString() -eq "ALL Servers"){
			$fname = $fname  + $datetime.ToString("yyyyMMdd")  + "-ALLSERVERS.csv"
			$mbServers | foreach-object{
				$mscombCollection += get-mailboxstatistics -server $_.Name
			}
		}
		else{
			$fname = $fname  + $datetime.ToString("yyyyMMdd")  + "-" + $snServerNameDrop.SelectedItem.ToString() + ".csv"
			$mscombCollection += get-mailboxstatistics -server $snServerNameDrop.SelectedItem.ToString()  
		} 
		$mscombCollection | ForEach-Object{
		$quQuota = "0"
		if ($usrquotas.ContainsKey($_.MailboxGUID)){
			if ($usrquotas[$_.MailboxGUID].Value -ne $null){
				if (CovertBitValue($usrquotas[$_.MailboxGUID]) -gt 0){
					$quQuota = "{0:P0}" -f (CovertBitValue($_.TotalItemSize)/CovertBitValue($usrquotas[$_.MailboxGUID]))}
				else{
					$quQuota = "100"
				}
			}
		}
		else{
		if ($mstoresquotas.ContainsKey($_.database)){
				if ($mstoresquotas[$_.database].Value -ne $null){
				$quQuota = "{0:P0}" -f (CovertBitValue($_.TotalItemSize)/CovertBitValue($mstoresquotas[$_.database]))}}
		}
	        $icount = 0
		$tisize = 0
		$disize = 0
		if ($_.DisplayName -ne $null){$dname = $_.DisplayName}
		if ($_.ItemCount -ne $null){$icount = $_.ItemCount}
		if (CovertBitValue($_.TotalItemSize) -ne $null){$tisize = CovertBitValue($_.TotalItemSize)}
		if (CovertBitValue($_.TotalDeletedItemSize) -ne $null){$disize = CovertBitValue($_.TotalDeletedItemSize)}    
		$msTable.Rows.add($dname,$icount,$tisize,$disize,$quQuota.replace("%","").replace(",",""))
		$mbcomb = "" | select Date,ServerName,DisplayName,ItemCount,TotalItemSize,TotalDeletedItemSize
		$mbcomb.Date = $datetime.ToString("yyyyMMdd")
		$mbcomb.ServerName = $snServerNameDrop.SelectedItem.ToString()
		$mbcomb.DisplayName = $dname
		$mbcomb.ItemCount = $icount
		$mbcomb.TotalItemSize = $tisize
		$mbcomb.TotalDeletedItemSize = $disize
		$mbcombCollection += $mbcomb
	}

} 
write-host $fstring 

$dgDataGrid.DataSource = $msTable
if ($fname -ne "disconnected") {
	$mbcombCollection | export-csv –encoding "unicode" -noTypeInformation $fname 
}

}

function ShowGrowth(){


$gtTable.clear()
$datetime = get-date
$arArrayList = New-Object System.Collections.ArrayList
dir $script:HistoryDir\*.csv | foreach-object{ 
	$fname = $_.name
	$nmArray = $_.name.split("-")
	if ($nmArray[1].Replace(".csv","") -eq $snServerNameDrop.SelectedItem.ToString()) {
		[VOID]$arArrayList.Add($nmArray[0])
	}
}
$arArrayList.Sort()
$spoint = $arArrayList[$arArrayList.Count-1]
$oneday = $spoint
$sevenday = $spoint
$onemonth = $spoint
$oneyear = $spoint
foreach ($file in $arArrayList){
	if ($file -gt ($datetime.Adddays(-2).ToString("yyyyMMdd")) -band $file -lt $oneday) {$oneday = $file} 
	if ($file -gt ($datetime.Adddays(-7).ToString("yyyyMMdd")) -band $file -lt $sevenday) {$sevenday = $file} 
	if ($file -gt ($datetime.Adddays(-31).ToString("yyyyMMdd")) -band $file -lt $onemonth) {$onemonth = $file} 
	if ($file -gt ($datetime.Adddays(-256).ToString("yyyyMMdd")) -band $file -lt $oneyear) {$oneyear = $file} 
}
write-host $oneday
write-host $sevenday
write-host $onemonth
write-host $oneyear

$onedaystats = @{ }
$sevendaystats = @{ }
$onemonthsdaystats = @{ }
$oneyearstats = @{ }

Import-Csv ("$script:HistoryDir\" + $oneday + "-" + $snServerNameDrop.SelectedItem.ToString() + ".csv") | %{ 
	$onedaystats.add($_.DisplayName,$_.TotalItemSize)	
}
Import-Csv ("$script:HistoryDir\" + $sevenday + "-" + $snServerNameDrop.SelectedItem.ToString() + ".csv") | %{ 
	$sevendaystats.add($_.DisplayName,$_.TotalItemSize)	
}
Import-Csv ("$script:HistoryDir\" + $onemonth + "-" + $snServerNameDrop.SelectedItem.ToString() + ".csv") | %{ 
	$onemonthsdaystats.add($_.DisplayName,$_.TotalItemSize)	
}
Import-Csv ("$script:HistoryDir\" + $oneyear + "-" + $snServerNameDrop.SelectedItem.ToString() + ".csv") | %{ 
	$oneyearstats.add($_.DisplayName,$_.TotalItemSize)	
}

foreach($row in $msTable.Rows){
	if ($onedaystats.ContainsKey($row[0].ToString())){
		$ondaysizegrowth = $row[2] - $onedaystats[$row[0].ToString()]
	}
	else{$ondaysizegrowth = 0}
	if ($sevendaystats.ContainsKey($row[0].ToString())){
		$sevendaysizegrowth = $row[2] - $sevendaystats[$row[0].ToString()]}
	else{$sevendaysizegrowth = 0}
	if ($onemonthsdaystats.ContainsKey($row[0].ToString())){
		$onemonthsizegrowth = $row[2] - $onemonthsdaystats[$row[0].ToString()]}
	else{$onemonthsizegrowth = 0}
	if ($oneyearstats.ContainsKey($row[0].ToString())){
		$oneyearsizegrowth = $row[2] - $oneyearstats[$row[0].ToString()]}
	else{$oneyearsizegrowth = 0}
	$gtTable.rows.add($row[0].ToString(),$row[2],$ondaysizegrowth,$sevendaysizegrowth,$onemonthsizegrowth,$oneyearsizegrowth)
	
}
$dgDataGrid.DataSource = $gtTable

}

function GetFolderSizes(){
$fsTable.clear()
$snServername = $snServerNameDrop.SelectedItem.ToString()
write-host $dgDataGrid.CurrentCell.RowIndex
get-user $msTable.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0] -RecipientType "UserMailbox" | foreach-object{
	$siSIDToSearch = $_
}
write-host $siSIDToSearch.WindowsEmailAddress.ToString()
Get-MailboxFolderStatistics $siSIDToSearch.WindowsEmailAddress.ToString() | ForEach-Object{
	$ficount = 0
	$fisize = 0
	$fsisize = 0
	$fscount = 0
	$fname = $_.Name
	$fsisize =  CovertBitValue($_.FolderSize)
	if ($_.ItemsInFolder -ne $null){$ficount = $_.ItemsInFolder}
	if ($_.ItemsInFolderAndSubfolders -ne $null){$fscount = $_.ItemsInFolderAndSubfolders} 
	$fsisize =  CovertBitValue($_.FolderAndSubfolderSize)      
	$fsTable.Rows.add($fname,$ficount,$fsisize,$fscount,$fsisize)
}
$dgDataGrid1.DataSource = $fsTable
}

function ExportMBcsv{

$exFileName = new-object System.Windows.Forms.saveFileDialog
$exFileName.DefaultExt = "csv"
$exFileName.Filter = "csv files (*.csv)|*.csv"
$exFileName.InitialDirectory = "c:\temp"
$exFileName.ShowHelp = $true
$exFileName.ShowDialog()
if ($exFileName.FileName -ne ""){
	$logfile = new-object IO.StreamWriter($exFileName.FileName,$true)
	$logfile.WriteLine("UserName,# Items,MB Size(MB),DelItems(MB),QuotaUsed")
	foreach($row in $msTable.Rows){
		$logfile.WriteLine("`"" + $row[0].ToString() + "`"," + $row[1].ToString() + "," + $row[2].ToString() + "," + $row[3].ToString() + "," + $row[4].ToString()) 
	}
	$logfile.Close()
}
}

function ExportFScsv{

$exFileName = new-object System.Windows.Forms.saveFileDialog
$exFileName.DefaultExt = "csv"
$exFileName.Filter = "csv files (*.csv)|*.csv"
$exFileName.InitialDirectory = "c:\temp"
$exFileName.ShowHelp = $true
$exFileName.ShowDialog()
if ($exFileName.FileName -ne ""){
	$logfile = new-object IO.StreamWriter($exFileName.FileName,$true)
	$logfile.WriteLine("DisplayName,# Items,Folder Size(MB),# Items + Sub,Folder Size + Sub(MB)")
	foreach($row in $fsTable.Rows){
		$logfile.WriteLine("`"" + $row[0].ToString() + "`"," + $row[3].ToString() + "," + $row[2].ToString() + "," + $row[3].ToString() + "," + $row[4].ToString()) 
	}
	$logfile.Close()
}
}

$usrquotas = @{ }
$mstoresquotas = @{ }
$form = new-object System.Windows.Forms.form 
$global:LastFolder = ""
# Add DataTable

$Dataset = New-Object System.Data.DataSet
$fsTable = New-Object System.Data.DataTable
$fsTable.TableName = "Folder Sizes"
$fsTable.Columns.Add("DisplayName")
$fsTable.Columns.Add("# Items",[int64])
$fsTable.Columns.Add("Folder Size(MB)",[int64])
$fsTable.Columns.Add("# Items + Sub",[int64])
$fsTable.Columns.Add("Folder Size + Sub(MB)",[int64])
$Dataset.tables.add($fsTable)

$msTable = New-Object System.Data.DataTable
$msTable.TableName = "Mailbox Sizes"
$msTable.Columns.Add("UserName")
$msTable.Columns.Add("# Items",[int64])
$msTable.Columns.Add("MB Size(MB)",[int64])
$msTable.Columns.Add("DelItems(MB)",[int64])
$msTable.Columns.Add("Quota Used",[int64])
$Dataset.tables.add($msTable)


$gtTable = New-Object System.Data.DataTable
$gtTable.TableName = "Mailbox Growth"
$gtTable.Columns.Add("UserName")
$gtTable.Columns.Add("Mailbox Size",[int64])
$gtTable.Columns.Add("1 Day",[int64])
$gtTable.Columns.Add("7 Days",[int64])
$gtTable.Columns.Add("31 Days",[int64])
$gtTable.Columns.Add("1 Year",[int64])
$Dataset.tables.add($gtTable)

# Add Server DropLable
$snServerNamelableBox = new-object System.Windows.Forms.Label
$snServerNamelableBox.Location = new-object System.Drawing.Size(10,20) 
$snServerNamelableBox.size = new-object System.Drawing.Size(80,20) 
$snServerNamelableBox.Text = "ServerName"
$form.Controls.Add($snServerNamelableBox) 

# Add Server Drop Down
$snServerNameDrop = new-object System.Windows.Forms.ComboBox
$snServerNameDrop.Location = new-object System.Drawing.Size(90,20)
$snServerNameDrop.Size = new-object System.Drawing.Size(100,30)
$snServerNameDrop.Items.Add("ALL Servers")
get-mailboxserver | ForEach-Object{$snServerNameDrop.Items.Add($_.Name)}
$snServerNameDrop.Add_SelectedValueChanged({getMailboxSizes})  
$form.Controls.Add($snServerNameDrop)

# Add Mailbox Type DropLable
$mtTypeDroplableBox = new-object System.Windows.Forms.Label
$mtTypeDroplableBox.Location = new-object System.Drawing.Size(200,20) 
$mtTypeDroplableBox.size = new-object System.Drawing.Size(70,20) 
$mtTypeDroplableBox.Text = "MailboxType"
$form.Controls.Add($mtTypeDroplableBox) 

# Add Mailbox Type Drop Down
$mtTypeDrop = new-object System.Windows.Forms.ComboBox
$mtTypeDrop.Location = new-object System.Drawing.Size(270,20)
$mtTypeDrop.Size = new-object System.Drawing.Size(100,30)
$mtTypeDrop.Items.Add("Disconnected")
$mtTypeDrop.Items.Add("Connected")
$mtTypeDrop.Add_SelectedValueChanged({if ($snServerNameDrop.SelectedItem -ne $null){getMailboxSizes}})  
$form.Controls.Add($mtTypeDrop)

# Add Quota Type DropLable
$qtTypeDroplableBox = new-object System.Windows.Forms.Label
$qtTypeDroplableBox.Location = new-object System.Drawing.Size(375,20) 
$qtTypeDroplableBox.size = new-object System.Drawing.Size(70,20) 
$qtTypeDroplableBox.Text = "Quota Type"
$form.Controls.Add($qtTypeDroplableBox) 

# Add Quota Type Drop Down
$qtTypeDrop = new-object System.Windows.Forms.ComboBox
$qtTypeDrop.Location = new-object System.Drawing.Size(455,20)
$qtTypeDrop.Size = new-object System.Drawing.Size(130,30)
$qtTypeDrop.Items.Add("Warning")
$qtTypeDrop.Items.Add("Proibit Send")
$qtTypeDrop.Items.Add("Proibit Send/Recieve")
$qtTypeDrop.Add_SelectedValueChanged({if ($snServerNameDrop.SelectedItem -ne $null){getMailboxSizes}})  
$form.Controls.Add($qtTypeDrop)

# Add Export MB Button

$exButton1 = new-object System.Windows.Forms.Button
$exButton1.Location = new-object System.Drawing.Size(10,560)
$exButton1.Size = new-object System.Drawing.Size(125,20)
$exButton1.Text = "Export Mailbox Grid"
$exButton1.Add_Click({ExportMBcsv})
$form.Controls.Add($exButton1)

# Add Export FG Button

$exButton2 = new-object System.Windows.Forms.Button
$exButton2.Location = new-object System.Drawing.Size(550,560)
$exButton2.Size = new-object System.Drawing.Size(135,20)
$exButton2.Text = "Export FolderSize Grid"
$exButton2.Add_Click({ExportFScsv})
$form.Controls.Add($exButton2)

# Add DataGrid View

$dgDataGrid = new-object System.windows.forms.DataGridView
$dgDataGrid.Location = new-object System.Drawing.Size(10,50) 
$dgDataGrid.size = new-object System.Drawing.Size(530,500)
$dgDataGrid.AutoSizeRowsMode = "AllHeaders"
$form.Controls.Add($dgDataGrid)

$dgDataGrid1 = new-object System.windows.forms.DataGridView
$dgDataGrid1.Location = new-object System.Drawing.Size(550,50) 
$dgDataGrid1.size = new-object System.Drawing.Size(450,500)
$dgDataGrid1.AutoSizeRowsMode = "AllHeaders"
$form.Controls.Add($dgDataGrid1)

# folder Size Button

$fsizeButton = new-object System.Windows.Forms.Button
$fsizeButton.Location = new-object System.Drawing.Size(600,19)
$fsizeButton.Size = new-object System.Drawing.Size(120,23)
$fsizeButton.Text = "Get Folder Size"
$fsizeButton.visible = $True
$fsizeButton.Add_Click({GetFolderSizes})
$form.Controls.Add($fsizeButton)

# Show Mailbox Size Growth History 

$mgrowButton = new-object System.Windows.Forms.Button
$mgrowButton.Location = new-object System.Drawing.Size(730,19)
$mgrowButton.Size = new-object System.Drawing.Size(120,23)
$mgrowButton.Text = "Show Growth History"
$mgrowButton.visible = $True
$mgrowButton.Add_Click({if($mgrowButton.Text -eq "Show Growth History"){$mgrowButton.Text = "Show Mailbox Size"
				$dgDataGrid.Location = new-object System.Drawing.Size(10,50) 
				$dgDataGrid.size = new-object System.Drawing.Size(650,500)
				$dgDataGrid1.Location = new-object System.Drawing.Size(670,50) 
				$dgDataGrid1.size = new-object System.Drawing.Size(250,500)
				ShowGrowth}
		       else{$mgrowButton.Text = "Show Growth History"
				$dgDataGrid.Location = new-object System.Drawing.Size(10,50) 
				$dgDataGrid.size = new-object System.Drawing.Size(530,500)
				$dgDataGrid1.Location = new-object System.Drawing.Size(550,50) 
				$dgDataGrid1.size = new-object System.Drawing.Size(450,500)
				getMailboxSizes}			
})
$form.Controls.Add($mgrowButton)



$form.Text = "Exchange 2010 Mailbox Size Form"
$form.size = new-object System.Drawing.Size(1000,620) 
$form.autoscroll = $true
$form.Add_Shown({$form.Activate()})
$form.ShowDialog()

