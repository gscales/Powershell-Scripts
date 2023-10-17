[VOID][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[VOID][System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
$Script:form = new-object System.Windows.Forms.form
$Script:Treeinfo = @{ }
function Invoke-SMCOpenMailbox {
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [String]$MailboxName,
		[Parameter(Position = 1, Mandatory = $false)] [String]$FolderName = "MsgFolderRoot"
	)
	
	Process {        
		$tvTreView.Nodes.Clear()
		$Script:Treeinfo.Clear()
		$rootFolder = Get-MgUserMailFolder -UserId $MailboxName -MailFolderId $FolderName 
		$Folders = Invoke-EnumerateMailBoxFolders -MailboxName $MailboxName -WellKnownFolder $FolderName   
		$Script:Treeinfo = @{ }       
		$TNRoot = new-object System.Windows.Forms.TreeNode("Root")
		$TNRoot.Name = "Mailbox"
		$TNRoot.Text = "Mailbox - " + $emEmailAddressTextBox.Text
		$exProgress = 0
		foreach ($ffFolder in $Folders) {
			#Process folder here
			$ParentFolderId = $ffFolder.parentFolderId
			$folderName = $ffFolder.displayName
			$TNChild = new-object System.Windows.Forms.TreeNode($ffFolder.Name)
			$TNChild.Name = $folderName
			$TNChild.Text = $folderName
			$TNChild.tag = $ffFolder			
			if ($ParentFolderId -eq $rootFolder.Id) {
				[void]$TNRoot.Nodes.Add($TNChild)
				$Script:Treeinfo.Add($ffFolder.Id.ToString(), $TNChild)
			}
			else {
				$pfFolder = $Script:Treeinfo[$ParentFolderId]
				[void]$pfFolder.Nodes.Add($TNChild)
				if ($Script:Treeinfo.ContainsKey($ffFolder.Id) -eq $false) {
					$Script:Treeinfo.Add($ffFolder.Id, $TNChild)
				}
			}
		}
		$Script:clickedFolder = $null
		[void]$tvTreView.Nodes.Add($TNRoot)
		Write-Progress -Activity "Executing Request" -Completed
	}
}

function Get-TaggedProperty {
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[String]
		$DataType,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[String]
		$Id,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[String]
		$Value
	)
	Begin {
		$Property = "" | Select-Object Id, DataType, PropertyType, Value
		$Property.Id = $Id
		$Property.DataType = $DataType
		$Property.PropertyType = "Tagged"
		if (![String]::IsNullOrEmpty($Value)) {
			$Property.Value = $Value
		}
		return, $Property
	}
}

function Get-ExtendedPropList {
	[CmdletBinding()]
	param (
		[Parameter(Position = 1, Mandatory = $false)]
		[PSCustomObject]
		$PropertyList
	)
	Begin {
		$rtString = "";
		$PropName = "Id"
		foreach ($Prop in $PropertyList) {
			if ($Prop.PropertyType -eq "Tagged") {
				if ($rtString -eq "") {
					$rtString = "($PropName eq '" + $Prop.DataType + " " + $Prop.Id + "')"
				}
				else {
					$rtString += " or ($PropName eq '" + $Prop.DataType + " " + $Prop.Id + "')"
				}
			}
			else {
				if ($Prop.Type -eq "String") {
					if ($rtString -eq "") {
						$rtString = "($PropName eq '" + $Prop.DataType + " {" + $Prop.Guid + "} Name " + $Prop.Id + "')"
					}
					else {
						$rtString += " or ($PropName eq '" + $Prop.DataType + " {" + $Prop.Guid + "} Name " + $Prop.Id + "')"
					}
				}
				else {
					if ($rtString -eq "") {
						$rtString = "($PropName eq '" + $Prop.DataType + " {" + $Prop.Guid + "} Id " + $Prop.Id + "')"
					}
					else {
						$rtString += " or ($PropName eq '" + $Prop.DataType + " {" + $Prop.Guid + "} Id " + $Prop.Id + "')"
					}
				}
			}
			
		}
		return "SingleValueExtendedProperties(`$filter=" + $rtString + ")" 	
	}
}

function Invoke-EnumerateMailBoxFolders {
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$FolderPath,

		[Parameter(Position = 1, Mandatory = $false)]
		[string]
		$WellKnownFolder,
		
		[Parameter(Position = 2, Mandatory = $true)]
		[String]
		$MailboxName,

		[Parameter(Position = 3, Mandatory = $false)]
		[switch]
		$returnSearchRoot
	)

	process {
		$Script:Mailboxfolders = @()
		if ($FolderPath) {
			$searchRootFolder = Get-MailBoxFolderFromPath -MailboxName $MailboxName -FolderPath $FolderPath
			Add-Member -InputObject $searchRootFolder -NotePropertyName "FolderPath" -NotePropertyValue $FolderPath
		}
		if ($WellKnownFolder) {
			$searchRootFolder = Get-MgUserMailFolder -UserId $MailboxName -MailFolderId $WellKnownFolder             
		}
		if ($returnSearchRoot) {               
			$Script:Mailboxfolders += $searchRootFolder

		}      
		if ($searchRootFolder) {
			Invoke-EnumerateChildMailFolders -MailboxName $MailboxName -folderId $searchRootFolder.id
		}      
		return $Script:Mailboxfolders 
	}
}

function Invoke-EnumerateChildMailFolders {
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$folderId,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[String]
		$MailboxName
	)

	process {
		$childFolders = Get-MgUserMailFolderChildFolder -UserId $MailboxName -MailFolderId $folderId -All -ExpandProperty "singleValueExtendedProperties(`$filter=id eq 'String 0x66b5')"
		Write-Verbose ("Returned " + $childFolders.Count)
		foreach ($childfolder in $childFolders) {
			#Expand-ExtendedProperties -Item $childfolder
			$Script:Mailboxfolders += $childfolder
			if ($childfolder.ChildFolderCount -gt 0) {
				Invoke-EnumerateChildMailFolders -MailboxName $MailboxName -folderId $childfolder.id
			}
		}
	}
}

function Start-SMCMailClient {
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[String]
		$MailboxName,
		[Parameter(Position = 1, Mandatory = $false)] [String]$FolderName = "MsgFolderRoot"
	)
	Process {
		$Script:form = new-object System.Windows.Forms.form
		$Script:Treeinfo = @{ }
		$mbtable = New-Object System.Data.DataTable
		$mbtable.TableName = "Folder Item"
		[Void]$mbtable.Columns.Add("From")
		[Void]$mbtable.Columns.Add("Subject")
		[Void]$mbtable.Columns.Add("Recieved", [DATETIME])
		[Void]$mbtable.Columns.Add("Size", [INT64])
		[Void]$mbtable.Columns.Add("ID")
		[Void]$mbtable.Columns.Add("hasAttachments")
		
		# Add Email Address 
		$emEmailAddressTextBox = new-object System.Windows.Forms.TextBox
		$emEmailAddressTextBox.Location = new-object System.Drawing.Size(130, 20)
		$emEmailAddressTextBox.size = new-object System.Drawing.Size(300, 20)
		$emEmailAddressTextBox.Enabled = $true
		$emEmailAddressTextBox.text = ""
		[Void]$Script:form.controls.Add($emEmailAddressTextBox)
		
		# Add  Email Address  Lable
		$emEmailAddresslableBox = new-object System.Windows.Forms.Label
		$emEmailAddresslableBox.Location = new-object System.Drawing.Size(10, 20)
		$emEmailAddresslableBox.size = new-object System.Drawing.Size(120, 20)
		$emEmailAddresslableBox.Text = "Email Address"
		[Void]$Script:form.controls.Add($emEmailAddresslableBox)
			
		$exButton1 = new-object System.Windows.Forms.Button
		$exButton1.Location = new-object System.Drawing.Size(10, 50)
		$exButton1.Size = new-object System.Drawing.Size(125, 20)
		$exButton1.Text = "Open Mailbox"
		$exButton1.Add_Click({ Open-ClientMailbox })
		[Void]$Script:form.Controls.Add($exButton1)
		
		# Add Numeric Results
		$neResultCheckNum = new-object System.Windows.Forms.numericUpDown
		$neResultCheckNum.Location = new-object System.Drawing.Size(250, 50)
		$neResultCheckNum.Size = new-object System.Drawing.Size(70, 30)
		$neResultCheckNum.Enabled = $true
		$neResultCheckNum.Value = 100
		$neResultCheckNum.Maximum = 10000000000
		[Void]$Script:form.Controls.Add($neResultCheckNum)
		
		$exButton2 = new-object System.Windows.Forms.Button
		$exButton2.Location = new-object System.Drawing.Size(330, 50)
		$exButton2.Size = new-object System.Drawing.Size(125, 25)
		$exButton2.Text = "Show Message"
		$exButton2.Add_Click({ Invoke-SMCShowClientMessage })
		[Void]$Script:form.Controls.Add($exButton2)
		
		$exButton5 = new-object System.Windows.Forms.Button
		$exButton5.Location = new-object System.Drawing.Size(455, 50)
		$exButton5.Size = new-object System.Drawing.Size(125, 25)
		$exButton5.Text = "Show Header"
		$exButton5.Add_Click({ Invoke-SMCShowClientHeader })
		[Void]$Script:form.Controls.Add($exButton5)
		
		$exButton6 = new-object System.Windows.Forms.Button
		$exButton6.Location = new-object System.Drawing.Size(330, 75)
		$exButton6.Size = new-object System.Drawing.Size(125, 25)
		$exButton6.Text = "New Message"
		$exButton6.Add_Click({ Invoke-SMCNewClientMessage })
		[Void]$Script:form.Controls.Add($exButton6)
		
		$exButton7 = new-object System.Windows.Forms.Button
		$exButton7.Location = new-object System.Drawing.Size(960, 85)
		$exButton7.Size = new-object System.Drawing.Size(90, 25)
		$exButton7.Text = "Update"
		$exButton7.Add_Click({ Get-SMCClientFolderItems })
		[Void]$Script:form.Controls.Add($exButton7)

		$exButton8 = new-object System.Windows.Forms.Button
		$exButton8.Location = new-object System.Drawing.Size(455, 75)
		$exButton8.Size = new-object System.Drawing.Size(125, 25)
		$exButton8.Text = "Export Message"
		$exButton8.Add_Click({ Invoke-SMCExportMessage })
		[Void]$Script:form.Controls.Add($exButton8)
		
		
		# Add Search Lable
		$saSeachBoxLable = new-object System.Windows.Forms.Label
		$saSeachBoxLable.Location = new-object System.Drawing.Size(600, 55)
		$saSeachBoxLable.Size = new-object System.Drawing.Size(170, 20)
		$saSeachBoxLable.Text = "Search by Property"
		[Void]$Script:form.controls.Add($saSeachBoxLable)
		
		$saNumItemsBoxLable = new-object System.Windows.Forms.Label
		$saNumItemsBoxLable.Location = new-object System.Drawing.Size(160, 55)
		$saNumItemsBoxLable.Size = new-object System.Drawing.Size(170, 20)
		$saNumItemsBoxLable.Text = "Number of Items"
		[Void]$Script:form.controls.Add($saNumItemsBoxLable)
		
		$seSearchCheck = new-object System.Windows.Forms.CheckBox
		$seSearchCheck.Location = new-object System.Drawing.Size(585, 50)
		$seSearchCheck.Size = new-object System.Drawing.Size(30, 25)
		[Void]$seSearchCheck.Add_Click({
				if ($seSearchCheck.Checked -eq $false) {
					$sbSearchTextBox.Enabled = $false
					$snSearchPropDrop.Enabled = $false
				}
				else {
					$sbSearchTextBox.Enabled = $true
					$snSearchPropDrop.Enabled = $true
				}
			})
		[Void]$Script:form.controls.Add($seSearchCheck)
		
		#Add Search box
		$snSearchPropDrop = new-object System.Windows.Forms.ComboBox
		$snSearchPropDrop.Location = new-object System.Drawing.Size(585, 85)
		$snSearchPropDrop.Size = new-object System.Drawing.Size(150, 30)
		$snSearchPropDrop.Items.Add("Subject")
		$snSearchPropDrop.Items.Add("Body")
		$snSearchPropDrop.Items.Add("From")
		$snSearchPropDrop.Enabled = $false
		[Void]$Script:form.Controls.Add($snSearchPropDrop)
		
		# Add Search TextBox
		$sbSearchTextBox = new-object System.Windows.Forms.TextBox
		$sbSearchTextBox.Location = new-object System.Drawing.Size(750, 85)
		$sbSearchTextBox.size = new-object System.Drawing.Size(200, 20)
		$sbSearchTextBox.Enabled = $false
		[Void]$Script:form.controls.Add($sbSearchTextBox)
		
		$tvTreView = new-object System.Windows.Forms.TreeView
		$tvTreView.Location = new-object System.Drawing.Size(10, 75)
		$tvTreView.size = new-object System.Drawing.Size(216, 400)
		$tvTreView.Anchor = "Top,left,Bottom"
		[Void]$tvTreView.add_AfterSelect({
				$Script:lfFolderID = $this.SelectedNode.tag
				Get-SMCClientFolderItems
				
			})
		[Void]$Script:form.Controls.Add($tvTreView)
		
		# Add DataGrid View
		$dgDataGrid = new-object System.windows.forms.DataGridView
		$dgDataGrid.Location = new-object System.Drawing.Size(250, 120)
		$dgDataGrid.size = new-object System.Drawing.Size(800, 600)
		$dgDataGrid.AutoSizeRowsMode = "AllHeaders"
		$dgDataGrid.AllowUserToDeleteRows = $false
		$dgDataGrid.AllowUserToAddRows = $false
		[Void]$Script:form.Controls.Add($dgDataGrid)
		
		$Script:form.Text = "Simple Exchange Mailbox Client"
		$Script:form.size = new-object System.Drawing.Size(1200, 800)
		$Script:form.autoscroll = $true
		[Void]$Script:form.Add_Shown({ $Script:form.Activate() })
		$emEmailAddressTextBox.Text = $MailboxName
		Invoke-SMCOpenMailbox -MailboxName $MailboxName -FolderName $FolderName 
		[Void]$Script:form.ShowDialog()
	}
}

function Get-SMCClientFolderItems {
	[CmdletBinding()]
	Param (
		
	)
	$props = @()
	$props += (Get-TaggedProperty -id 0x0E08 -DataType Long)
	$mbtable.Clear()
	$folder = $Script:lfFolderID
	if ($seSearchCheck.Checked) {
		$sfilter = ""
		switch ($snSearchPropDrop.SelectedItem.ToString()) {
			"Subject" {
				$sfilter = "Subject:'" + $sbSearchTextBox.Text.ToString() + "'"				
			}
			"Body" {
				$sfilter = "`"Body:'" + $sbSearchTextBox.Text.ToString() + "'`""				
			}
			"From" {
				$sfilter = "`"From:'" + $sbSearchTextBox.Text.ToString() + "'`""				
			}
		}
		if (!$sfilter -eq "") {
			$Items = Get-MgUserMailFolderMessage -UserId $emEmailAddressTextBox.Text -MailFolderId $folder.id -Top $neResultCheckNum.Value -Search $sfilter -ExpandProperty (Get-ExtendedPropList -PropertyList $props)
		}
	}
	else {
		$Items = Get-MgUserMailFolderMessage -UserId $emEmailAddressTextBox.Text -MailFolderId $folder.id -Top $neResultCheckNum.Value -ExpandProperty (Get-ExtendedPropList -PropertyList $props)
	}
	foreach ($mail in $Items) {
		if ($mail.sender.emailAddress.name -ne $null) { $fnFromName = $mail.sender.emailAddress.name }
		else { $fnFromName = "N/A" }
		if ($mail.Subject -ne $null) { $sbSubject = $mail.Subject.ToString() }
		else { $sbSubject = "N/A" }
		$messageSize = [math]::round($mail.SingleValueExtendedProperties[0].value / 1Kb, 0)
		$mbtable.rows.add($fnFromName, $sbSubject, $mail.receivedDateTime, $messageSize, $mail.id, $mail.hasAttachments)
	}
	$dgDataGrid.DataSource = $mbtable
}

function Invoke-SMCExportMessage {
	[CmdletBinding()]
	Param (
		$MessageID
	)	
	$MessageID = $mbtable.DefaultView[$dgDataGrid.CurrentCell.RowIndex][4]
	$saveFileDialog = [System.Windows.Forms.SaveFileDialog]@{
		CheckPathExists  = $true
		CreatePrompt     = $true
		OverwritePrompt  = $true
		InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
		FileName         = $mbtable.DefaultView[$dgDataGrid.CurrentCell.RowIndex][1]
		Title            = 'Choose directory to save the exported message'
		Filter           = "Email Message documents (.eml)|*.eml"
	}
	if ($saveFileDialog.ShowDialog() -eq 'Ok') {
		Get-MgUserMessageContent -UserId  $emEmailAddressTextBox.Text -MessageId $MessageID -OutFile $saveFileDialog.FileName 
	}
	
}
function Invoke-SMCShowClientMessage {
	[CmdletBinding()]
	Param (
		$MessageID
	)	
	$MessageID = $mbtable.DefaultView[$dgDataGrid.CurrentCell.RowIndex][4]
	$script:msMessage = Get-MgUserMailFolderMessage -UserId $emEmailAddressTextBox.Text -MailFolderId AllItems -MessageId $MessageID -ExpandProperty Attachments
	write-host $MessageID
	$msgform = new-object System.Windows.Forms.form
	$msgform.Text = $script:msMessage.Subject
	$msgform.size = new-object System.Drawing.Size(1000, 800)
	
	
	# Add Message From Lable
	$miMessageTolableBox = new-object System.Windows.Forms.Label
	$miMessageTolableBox.Location = new-object System.Drawing.Size(20, 20)
	$miMessageTolableBox.size = new-object System.Drawing.Size(80, 20)
	$miMessageTolableBox.Text = "To"
	$msgform.controls.Add($miMessageTolableBox)
	
	# Add MessageID Lable
	$miMessageSentlableBox = new-object System.Windows.Forms.Label
	$miMessageSentlableBox.Location = new-object System.Drawing.Size(20, 40)
	$miMessageSentlableBox.size = new-object System.Drawing.Size(80, 20)
	$miMessageSentlableBox.Text = "From"
	$msgform.controls.Add($miMessageSentlableBox)
	
	# Add Message Subject Lable
	$miMessageSubjectlableBox = new-object System.Windows.Forms.Label
	$miMessageSubjectlableBox.Location = new-object System.Drawing.Size(20, 60)
	$miMessageSubjectlableBox.size = new-object System.Drawing.Size(80, 20)
	$miMessageSubjectlableBox.Text = "Subject"
	$msgform.controls.Add($miMessageSubjectlableBox)
	
	# Add Message To
	$miMessageTotextlabelBox = new-object System.Windows.Forms.Label
	$miMessageTotextlabelBox.Location = new-object System.Drawing.Size(100, 20)
	$miMessageTotextlabelBox.size = new-object System.Drawing.Size(400, 20)
	$msgform.controls.Add($miMessageTotextlabelBox)
	$ToRecips = "";
	foreach ($torcp in $msMessage.toRecipients) {
		$ToRecips += $torcp.emailAddress.address.ToString() + ";"
	}
	$miMessageTotextlabelBox.Text = $ToRecips
	
	# Add Message From
	$miMessageSenttextlabelBox = new-object System.Windows.Forms.Label
	$miMessageSenttextlabelBox.Location = new-object System.Drawing.Size(100, 40)
	$miMessageSenttextlabelBox.size = new-object System.Drawing.Size(600, 20)
	$msgform.controls.Add($miMessageSenttextlabelBox)
	$miMessageSenttextlabelBox.Text = $msMessage.sender.emailAddress.name.ToString() + " (" + $msMessage.sender.emailAddress.address.ToString() + ")"
	
	# Add Message Subject 
	$miMessageSubjecttextlabelBox = new-object System.Windows.Forms.Label
	$miMessageSubjecttextlabelBox.Location = new-object System.Drawing.Size(100, 60)
	$miMessageSubjecttextlabelBox.size = new-object System.Drawing.Size(600, 20)
	$msgform.controls.Add($miMessageSubjecttextlabelBox)
	$miMessageSubjecttextlabelBox.Text = $msMessage.Subject.ToString()
	
	# Add Message body 
	$miMessageBodytextlabelBox = new-object System.Windows.Forms.WebBrowser
	$miMessageBodytextlabelBox.Location = new-object System.Drawing.Size(100, 80)
	$miMessageBodytextlabelBox.size = new-object System.Drawing.Size(900, 550)
	$miMessageBodytextlabelBox.AutoSize = $true
	$miMessageBodytextlabelBox.DocumentText = $msMessage.Body.Content
	$msgform.controls.Add($miMessageBodytextlabelBox)
	
	# Add Message Attachments Lable
	$miMessageAttachmentslableBox = new-object System.Windows.Forms.Label
	$miMessageAttachmentslableBox.Location = new-object System.Drawing.Size(20, 645)
	$miMessageAttachmentslableBox.size = new-object System.Drawing.Size(80, 20)
	$miMessageAttachmentslableBox.Text = "Attachments"
	$msgform.controls.Add($miMessageAttachmentslableBox)
	
	$miMessageAttachmentslableBox1 = new-object System.Windows.Forms.Label
	$miMessageAttachmentslableBox1.Location = new-object System.Drawing.Size(100, 645)
	$miMessageAttachmentslableBox1.size = new-object System.Drawing.Size(600, 20)
	$miMessageAttachmentslableBox1.Text = ""
	$msgform.Controls.Add($miMessageAttachmentslableBox1)
	
	
	$exButton4 = new-object System.Windows.Forms.Button
	$exButton4.Location = new-object System.Drawing.Size(10, 665)
	$exButton4.Size = new-object System.Drawing.Size(150, 20)
	$exButton4.Text = "Download Attachments"
	$exButton4.Enabled = $false
	$exButton4.Add_Click({ Invoke-SMCSaveAttachment })
	$msgform.Controls.Add($exButton4)
	
	$attname = ""
	if ($script:msMessage.hasattachments) {
		write-host "Attachment"
		$exButton4.Enabled = $true		
		foreach ($attach in $script:msMessage.Attachments) {
			$attname = $attname + $attach.Name.ToString() + "; "
		}
	}
	$miMessageAttachmentslableBox1.Text = $attname
	# Add Download Button
	
	$msgform.autoscroll = $true
	$msgform.Add_Shown({ $Script:form.Activate() })
	$msgform.ShowDialog()	
}

function Invoke-SMCSaveAttachment {
	[CmdletBinding()]
	Param (
		
	)
	
	$dlfolder = new-object -ComObject shell.application
	$dlfolderpath = $dlfolder.BrowseForFolder(0, "Download attachments to", 0)
	foreach ($attachment in $script:msMessage.Attachments) {
		$fiFile = new-object System.IO.FileStream(($dlfolderpath.Self.Path + "\" + $attachment.Name.ToString()), [System.IO.FileMode]::Create)
		$attachBytes = [System.Convert]::FromBase64String($attachment.AdditionalProperties.contentBytes)
		$fiFile.Write($attachBytes, 0, $attachBytes.Length)
		$fiFile.Close()
		write-host ("Downloaded Attachment : " + (($dlfolderpath.Self.Path + "\" + $attachment.Name.ToString())))
	}
}

function Invoke-SMCShowClientHeader {
	[CmdletBinding()]
	Param (
		
	)
	$Props = @()
	$Props += (Get-TaggedProperty -Id "0x007D" -DataType "String")
	$MessageID = $mbtable.DefaultView[$dgDataGrid.CurrentCell.RowIndex][4]
	$script:msMessage = Get-MgUserMailFolderMessage -UserId $emEmailAddressTextBox.Text -MailFolderId AllItems -MessageId $MessageID -ExpandProperty (Get-ExtendedPropList -PropertyList $Props)
	write-host $MessageID
	$hdrform = new-object System.Windows.Forms.form
	$hdrform.Text = $script:msMessage.Subject
	$hdrform.size = new-object System.Drawing.Size(800, 600)
	
	# Add Message header
	$miMessageHeadertextlabelBox = new-object System.Windows.Forms.RichTextBox
	$miMessageHeadertextlabelBox.Location = new-object System.Drawing.Size(10, 10)
	$miMessageHeadertextlabelBox.size = new-object System.Drawing.Size(800, 600)
	$miMessageHeadertextlabelBox.text = $script:msMessage.singleValueExtendedProperties[0].Value
	$hdrform.controls.Add($miMessageHeadertextlabelBox)
	$hdrform.autoscroll = $true
	$hdrform.Add_Shown({ $Script:form.Activate() })
	$hdrform.ShowDialog()
}

function Invoke-SMCNewClientMessage {
	[CmdletBinding()]
	Param (
		$Reply
	)
	
	$script:newmsgform = new-object System.Windows.Forms.form
	$script:newmsgform.Text = "New Message"
	$script:newmsgform.size = new-object System.Drawing.Size(1000, 800)
	
	# Add Message To Lable
	$miMessageTolableBox = new-object System.Windows.Forms.Label
	$miMessageTolableBox.Location = new-object System.Drawing.Size(20, 20)
	$miMessageTolableBox.size = new-object System.Drawing.Size(80, 20)
	$miMessageTolableBox.Text = "To"
	$script:newmsgform.controls.Add($miMessageTolableBox)
	
	# Add Message Subject Lable
	$miMessageSubjectlableBox = new-object System.Windows.Forms.Label
	$miMessageSubjectlableBox.Location = new-object System.Drawing.Size(20, 65)
	$miMessageSubjectlableBox.size = new-object System.Drawing.Size(80, 20)
	$miMessageSubjectlableBox.Text = "Subject"
	$script:newmsgform.controls.Add($miMessageSubjectlableBox)
	
	# Add Message To
	$miMessageTotextlabelBox = new-object System.Windows.Forms.TextBox
	$miMessageTotextlabelBox.Location = new-object System.Drawing.Size(100, 20)
	$miMessageTotextlabelBox.size = new-object System.Drawing.Size(400, 20)
	$script:newmsgform.controls.Add($miMessageTotextlabelBox)
	
	# Add Message Subject 
	$miMessageSubjecttextlabelBox = new-object System.Windows.Forms.TextBox
	$miMessageSubjecttextlabelBox.Location = new-object System.Drawing.Size(100, 65)
	$miMessageSubjecttextlabelBox.size = new-object System.Drawing.Size(600, 20)
	$script:newmsgform.controls.Add($miMessageSubjecttextlabelBox)
	
	
	# Add Message body 
	$miMessageBodytextlabelBox = new-object System.Windows.Forms.RichTextBox
	$miMessageBodytextlabelBox.Location = new-object System.Drawing.Size(100, 100)
	$miMessageBodytextlabelBox.size = new-object System.Drawing.Size(600, 350)
	$script:newmsgform.controls.Add($miMessageBodytextlabelBox)
	
	# Add Message Attachments Lable
	$miMessageAttachmentslableBox = new-object System.Windows.Forms.Label
	$miMessageAttachmentslableBox.Location = new-object System.Drawing.Size(20, 460)
	$miMessageAttachmentslableBox.size = new-object System.Drawing.Size(80, 20)
	$miMessageAttachmentslableBox.Text = "Attachments"
	$script:newmsgform.controls.Add($miMessageAttachmentslableBox)
	
	$miMessageAttachmentslableBox1 = new-object System.Windows.Forms.Label
	$miMessageAttachmentslableBox1.Location = new-object System.Drawing.Size(100, 460)
	$miMessageAttachmentslableBox1.size = new-object System.Drawing.Size(600, 20)
	$miMessageAttachmentslableBox1.Text = ""
	$script:newmsgform.Controls.Add($miMessageAttachmentslableBox1)
	
	$exButton7 = new-object System.Windows.Forms.Button
	$exButton7.Location = new-object System.Drawing.Size(95, 520)
	$exButton7.Size = new-object System.Drawing.Size(125, 20)
	$exButton7.Text = "Send Message"
	$exButton7.Add_Click({ Invoke-SMCSendClientMessage })
	$script:newmsgform.Controls.Add($exButton7)
	
	$exButton4 = new-object System.Windows.Forms.Button
	$exButton4.Location = new-object System.Drawing.Size(95, 490)
	$exButton4.Size = new-object System.Drawing.Size(150, 20)
	$exButton4.Text = "Add Attachment"
	$exButton4.Enabled = $true
	$exButton4.Add_Click({ Invoke-SMCSelectClientAttachment })
	
	$script:Attachments = @()
	
	$script:newmsgform.Controls.Add($exButton4)
	$script:newmsgform.autoscroll = $true
	$script:newmsgform.Add_Shown({ $Script:form.Activate() })
	$script:newmsgform.ShowDialog()	
}

function Invoke-SMCSendClientMessage {
	[CmdletBinding()]
	Param (
		
	)
	Send-SMCMessageREST -MailboxName $emEmailAddressTextBox.Text -ToRecipients @(New-SMCEmailAddress -Address $miMessageTotextlabelBox.Text) -Subject $miMessageSubjecttextlabelBox.Text -Body $miMessageBodytextlabelBox.Text -Attachments $script:Attachments
	$script:newmsgform.close()
	
}

function New-SMCEmailAddress {
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$Name,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$Address
	)
	Begin {
		$EmailAddress = "" | Select-Object Name, Address
		if ([String]::IsNullOrEmpty($Name)) {
			$EmailAddress.Name = $Address
		}
		else {
			$EmailAddress.Name = $Name
		}
		$EmailAddress.Address = $Address
		return, $EmailAddress
	}
}


function Send-SMCMessageREST {
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[string]
		$FolderPath,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[PSCustomObject]
		$Folder,
		
		[Parameter(Position = 4, Mandatory = $true)]
		[String]
		$Subject,
		
		[Parameter(Position = 5, Mandatory = $false)]
		[String]
		$Body,
		
		[Parameter(Position = 7, Mandatory = $false)]
		[psobject]
		$SenderEmailAddress,
		
		[Parameter(Position = 8, Mandatory = $false)]
		[psobject]
		$Attachments,
		
		[Parameter(Position = 9, Mandatory = $false)]
		[psobject]
		$ReferanceAttachments,
		
		[Parameter(Position = 10, Mandatory = $false)]
		[psobject]
		$ToRecipients,
		
		[Parameter(Position = 11, Mandatory = $false)]
		[psobject]
		$CCRecipients,
		
		[Parameter(Position = 12, Mandatory = $false)]
		[psobject]
		$BCCRecipients,
		
		[Parameter(Position = 13, Mandatory = $false)]
		[psobject]
		$ExPropList,
		
		[Parameter(Position = 14, Mandatory = $false)]
		[psobject]
		$StandardPropList,
		
		[Parameter(Position = 15, Mandatory = $false)]
		[string]
		$ItemClass,
		
		[Parameter(Position = 16, Mandatory = $false)]
		[switch]
		$SaveToSentItems,
		
		[Parameter(Position = 17, Mandatory = $false)]
		[switch]
		$ShowRequest,
		
		[Parameter(Position = 18, Mandatory = $false)]
		[switch]
		$RequestReadRecipient,
		
		[Parameter(Position = 19, Mandatory = $false)]
		[switch]
		$RequestDeliveryRecipient,
		
		[Parameter(Position = 20, Mandatory = $false)]
		[psobject]
		$ReplyTo
	)
	Begin {
		
		if (![String]::IsNullOrEmpty($ItemClass)) {
			$ItemClassProp = Get-EXRTaggedProperty -DataType "String" -Id "0x001A" -Value $ItemClass
			if ($ExPropList -eq $null) {
				$ExPropList = @()
			}
			$ExPropList += $ItemClassProp
		}
		$SaveToSentFolder = "false"
		if ($SaveToSentItems.IsPresent) {
			$SaveToSentFolder = "true"
		}
		$NewMessage = Get-MessageJSONFormat -Subject $Subject -Body $Body -SenderEmailAddress $SenderEmailAddress -Attachments $Attachments -ReferanceAttachments $ReferanceAttachments -ToRecipients $ToRecipients -SentDate $SentDate -ExPropList $ExPropList -CcRecipients $CCRecipients -bccRecipients $BCCRecipients -StandardPropList $StandardPropList -SaveToSentItems $SaveToSentFolder -SendMail -ReplyTo $ReplyTo -RequestReadRecipient $RequestReadRecipient.IsPresent -RequestDeliveryRecipient $RequestDeliveryRecipient.IsPresent
		if ($ShowRequest.IsPresent) {
			write-host $NewMessage
		}
		Send-MgUserMail -UserId $MailboxName -BodyParameter $NewMessage 
	
	}
}

function Get-MessageJSONFormat {
	[CmdletBinding()]
	param (
		[Parameter(Position = 1, Mandatory = $false)]
		[String]
		$Subject,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[String]
		$Body,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[psobject]
		$SenderEmailAddress,
		
		[Parameter(Position = 5, Mandatory = $false)]
		[psobject]
		$Attachments,
		
		[Parameter(Position = 5, Mandatory = $false)]
		[psobject]
		$ReferanceAttachments,
		
		[Parameter(Position = 6, Mandatory = $false)]
		[psobject]
		$ToRecipients,
		
		[Parameter(Position = 7, Mandatory = $false)]
		[psobject]
		$CcRecipients,
		
		[Parameter(Position = 7, Mandatory = $false)]
		[psobject]
		$bccRecipients,
		
		[Parameter(Position = 8, Mandatory = $false)]
		[psobject]
		$SentDate,
		
		[Parameter(Position = 9, Mandatory = $false)]
		[psobject]
		$StandardPropList,
		
		[Parameter(Position = 10, Mandatory = $false)]
		[psobject]
		$ExPropList,
		
		[Parameter(Position = 11, Mandatory = $false)]
		[switch]
		$ShowRequest,
		
		[Parameter(Position = 12, Mandatory = $false)]
		[String]
		$SaveToSentItems,
		
		[Parameter(Position = 13, Mandatory = $false)]
		[switch]
		$SendMail,
		
		[Parameter(Position = 14, Mandatory = $false)]
		[psobject]
		$ReplyTo,
		
		[Parameter(Position = 17, Mandatory = $false)]
		[bool]
		$RequestReadRecipient,
		
		[Parameter(Position = 18, Mandatory = $false)]
		[bool]
		$RequestDeliveryRecipient
	)
	Begin {
		$NewMessage = "{" + "`r`n"
		if ($SendMail.IsPresent) {
			$NewMessage += "  `"Message`" : {" + "`r`n"
		}
		if (![String]::IsNullOrEmpty($Subject)) {
			$NewMessage += "`"Subject`": `"" + $Subject + "`"" + "`r`n"
		}
		if ($SenderEmailAddress -ne $null) {
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"Sender`":{" + "`r`n"
			$NewMessage += " `"EmailAddress`":{" + "`r`n"
			$NewMessage += "  `"Name`":`"" + $SenderEmailAddress.Name + "`"," + "`r`n"
			$NewMessage += "  `"Address`":`"" + $SenderEmailAddress.Address + "`"" + "`r`n"
			$NewMessage += "}}" + "`r`n"
		}
		if (![String]::IsNullOrEmpty($Body)) {
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"Body`": {" + "`r`n"
			$NewMessage += "`"ContentType`": `"HTML`"," + "`r`n"
			$NewMessage += "`"Content`": `"" + $Body + "`"" + "`r`n"
			$NewMessage += "}" + "`r`n"
		}
		
		$toRcpcnt = 0;
		if ($ToRecipients -ne $null) {
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"ToRecipients`": [ " + "`r`n"
			foreach ($EmailAddress in $ToRecipients) {
				if ($toRcpcnt -gt 0) {
					$NewMessage += "      ,{ " + "`r`n"
				}
				else {
					$NewMessage += "      { " + "`r`n"
				}
				$NewMessage += " `"EmailAddress`":{" + "`r`n"
				$NewMessage += "  `"Name`":`"" + $EmailAddress.Name + "`"," + "`r`n"
				$NewMessage += "  `"Address`":`"" + $EmailAddress.Address + "`"" + "`r`n"
				$NewMessage += "}}" + "`r`n"
				$toRcpcnt++
			}
			$NewMessage += "  ]" + "`r`n"
		}
		$ccRcpcnt = 0
		if ($CcRecipients -ne $null) {
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"CcRecipients`": [ " + "`r`n"
			foreach ($EmailAddress in $CcRecipients) {
				if ($ccRcpcnt -gt 0) {
					$NewMessage += "      ,{ " + "`r`n"
				}
				else {
					$NewMessage += "      { " + "`r`n"
				}
				$NewMessage += " `"EmailAddress`":{" + "`r`n"
				$NewMessage += "  `"Name`":`"" + $EmailAddress.Name + "`"," + "`r`n"
				$NewMessage += "  `"Address`":`"" + $EmailAddress.Address + "`"" + "`r`n"
				$NewMessage += "}}" + "`r`n"
				$ccRcpcnt++
			}
			$NewMessage += "  ]" + "`r`n"
		}
		$bccRcpcnt = 0
		if ($bccRecipients -ne $null) {
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"BccRecipients`": [ " + "`r`n"
			foreach ($EmailAddress in $bccRecipients) {
				if ($bccRcpcnt -gt 0) {
					$NewMessage += "      ,{ " + "`r`n"
				}
				else {
					$NewMessage += "      { " + "`r`n"
				}
				$NewMessage += " `"EmailAddress`":{" + "`r`n"
				$NewMessage += "  `"Name`":`"" + $EmailAddress.Name + "`"," + "`r`n"
				$NewMessage += "  `"Address`":`"" + $EmailAddress.Address + "`"" + "`r`n"
				$NewMessage += "}}" + "`r`n"
				$bccRcpcnt++
			}
			$NewMessage += "  ]" + "`r`n"
		}
		$ReplyTocnt = 0
		if ($ReplyTo -ne $null) {
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"ReplyTo`": [ " + "`r`n"
			foreach ($EmailAddress in $ReplyTo) {
				if ($ReplyTocnt -gt 0) {
					$NewMessage += "      ,{ " + "`r`n"
				}
				else {
					$NewMessage += "      { " + "`r`n"
				}
				$NewMessage += " `"EmailAddress`":{" + "`r`n"
				$NewMessage += "  `"Name`":`"" + $EmailAddress.Name + "`"," + "`r`n"
				$NewMessage += "  `"Address`":`"" + $EmailAddress.Address + "`"" + "`r`n"
				$NewMessage += "}}" + "`r`n"
				$ReplyTocnt++
			}
			$NewMessage += "  ]" + "`r`n"
		}
		if ($RequestDeliveryRecipient) {
			$NewMessage += ",`"IsDeliveryReceiptRequested`": true`r`n"
		}
		if ($RequestReadRecipient) {
			$NewMessage += ",`"IsReadReceiptRequested`": true `r`n"
		}
		if ($StandardPropList -ne $null) {
			foreach ($StandardProp in $StandardPropList) {
				if ($NewMessage.Length -gt 5) { $NewMessage += "," }
				switch ($StandardProp.PropertyType) {
					"Single" {
						if ($StandardProp.QuoteValue) {
							$NewMessage += "`"" + $StandardProp.Name + "`": `"" + $StandardProp.Value + "`"" + "`r`n"
						}
						else {
							$NewMessage += "`"" + $StandardProp.Name + "`": " + $StandardProp.Value + "`r`n"
						}
						
						
					}
					"Object" {
						if ($StandardProp.isArray) {
							$NewMessage += "`"" + $StandardProp.PropertyName + "`": [ {" + "`r`n"
						}
						else {
							$NewMessage += "`"" + $StandardProp.PropertyName + "`": {" + "`r`n"
						}
						$acCount = 0
						foreach ($PropKeyValue in $StandardProp.PropertyList) {
							if ($acCount -gt 0) {
								$NewMessage += ","
							}
							$NewMessage += "`"" + $PropKeyValue.Name + "`": `"" + $PropKeyValue.Value + "`"" + "`r`n"
							$acCount++
						}
						if ($StandardProp.isArray) {
							$NewMessage += "}]" + "`r`n"
						}
						else {
							$NewMessage += "}" + "`r`n"
						}
						
					}
					"ObjectCollection" {
						if ($StandardProp.isArray) {
							$NewMessage += "`"" + $StandardProp.PropertyName + "`": [" + "`r`n"
						}
						else {
							$NewMessage += "`"" + $StandardProp.PropertyName + "`": {" + "`r`n"
						}
						foreach ($EnclosedStandardProp in $StandardProp.PropertyList) {
							$NewMessage += "`"" + $EnclosedStandardProp.PropertyName + "`": {" + "`r`n"
							foreach ($PropKeyValue in $EnclosedStandardProp.PropertyList) {
								$NewMessage += "`"" + $PropKeyValue.Name + "`": `"" + $PropKeyValue.Value + "`"," + "`r`n"
							}
							$NewMessage += "}" + "`r`n"
						}
						if ($StandardProp.isArray) {
							$NewMessage += "]" + "`r`n"
						}
						else {
							$NewMessage += "}" + "`r`n"
						}
					}
					
				}
			}
		}
		$atcnt = 0
		$processAttachments = $false
		if ($Attachments -ne $null) { $processAttachments = $true }
		if ($ReferanceAttachments -ne $null) { $processAttachments = $true }
		if ($processAttachments) {
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "  `"Attachments`": [ " + "`r`n"
			if ($Attachments -ne $null) {
				foreach ($Attachment in $Attachments) {
					if ($atcnt -gt 0) {
						$NewMessage += "   ,{" + "`r`n"
					}
					else {
						$NewMessage += "    {" + "`r`n"
					}
					if ($Attachment.name) {
						$NewMessage += "     `"@odata.type`": `"#microsoft.graph.fileAttachment`"," + "`r`n"
						$NewMessage += "     `"Name`": `"" + $Attachment.name + "`"," + "`r`n"
						$NewMessage += "     `"ContentBytes`": `" " + $Attachment.contentBytes + "`"" + "`r`n"
					}
					else {
						$Item = Get-Item $Attachment

						$NewMessage += "     `"@odata.type`": `"#microsoft.graph.fileAttachment`"," + "`r`n"
						$NewMessage += "     `"Name`": `"" + $Item.Name + "`"," + "`r`n"
						$NewMessage += "     `"ContentBytes`": `" " + [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($Attachment)) + "`"" + "`r`n"

					}
					$NewMessage += "    } " + "`r`n"
					$atcnt++
					
				}
			}
			$atcnt = 0
			if ($ReferanceAttachments -ne $null) {
				foreach ($Attachment in $ReferanceAttachments) {
					if ($atcnt -gt 0) {
						$NewMessage += "   ,{" + "`r`n"
					}
					else {
						$NewMessage += "    {" + "`r`n"
					}
					$NewMessage += "     `"@odata.type`": `"#microsoft.graph.referenceAttachment`"," + "`r`n"
					$NewMessage += "     `"Name`": `"" + $Attachment.Name + "`"," + "`r`n"
					$NewMessage += "     `"SourceUrl`": `"" + $Attachment.SourceUrl + "`"," + "`r`n"
					$NewMessage += "     `"ProviderType`": `"" + $Attachment.ProviderType + "`"," + "`r`n"
					$NewMessage += "     `"Permission`": `"" + $Attachment.Permission + "`"," + "`r`n"
					$NewMessage += "     `"IsFolder`": `"" + $Attachment.IsFolder + "`"" + "`r`n"
					$NewMessage += "    } " + "`r`n"
					$atcnt++
				}
			}
			$NewMessage += "  ]" + "`r`n"
		}
		
		if ($ExPropList -ne $null) {
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"SingleValueExtendedProperties`": [" + "`r`n"
			$propCount = 0
			foreach ($Property in $ExPropList) {
				if ($propCount -eq 0) {
					$NewMessage += "{" + "`r`n"
				}
				else {
					$NewMessage += ",{" + "`r`n"
				}
				if ($Property.PropertyType -eq "Tagged") {
					$NewMessage += "`"Id`":`"" + $Property.DataType + " " + $Property.Id + "`", " + "`r`n"
				}
				else {
					if ($Property.Type -eq "String") {
						$NewMessage += "`"Id`":`"" + $Property.DataType + " " + $Property.Guid + " Name " + $Property.Id + "`", " + "`r`n"
					}
					else {
						$NewMessage += "`"Id`":`"" + $Property.DataType + " " + $Property.Guid + " Id " + $Property.Id + "`", " + "`r`n"
					}
				}
				if ($Property.Value -eq "null") {
					$NewMessage += "`"Value`":null" + "`r`n"
				}
				else {
					$NewMessage += "`"Value`":`"" + $Property.Value + "`"" + "`r`n"
				}				
				$NewMessage += " } " + "`r`n"
				$propCount++
			}
			$NewMessage += "]" + "`r`n"
		}
		if (![String]::IsNullOrEmpty($SaveToSentItems)) {
			$NewMessage += "}   ,`"SaveToSentItems`": `"" + $SaveToSentItems.ToLower() + "`"" + "`r`n"
		}
		$NewMessage += "}"
		if ($ShowRequest.IsPresent) {
			Write-Host $NewMessage
		}
		return, $NewMessage
	}
}

function Invoke-SMCSelectClientAttachment {
	[CmdletBinding()]
	Param (
		
	)
	
	$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
		Multiselect = $true
	}
	
	[void]$FileBrowser.ShowDialog()
	foreach ($File in $FileBrowser.FileNames) {
		$script:Attachments += $File
		$attname += $File + " "
	}
	$miMessageAttachmentslableBox1.Text = $attname
	
}

