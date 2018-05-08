
function Get-EWSDigestEmailBody
{
	[CmdletBinding()] 
    param (
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$MessageList,
		[Parameter(Position = 2, Mandatory = $false)]
		[switch]
		$weblink,
		[Parameter(Position = 3, Mandatory = $false)]
		[switch]
		$Detail,
		[Parameter(Position = 4, Mandatory = $false)]
		[String]
		$InfoField1Name,
		[Parameter(Position = 5, Mandatory = $false)]
		[String]
		$InfoField2Name,
		[Parameter(Position = 6, Mandatory = $false)]
		[String]
		$InfoField3Name,
		[Parameter(Position = 7, Mandatory = $false)]
		[String]
		$InfoField4Name,
		[Parameter(Position = 8, Mandatory = $false)]
		[String]
        $InfoField5Name
	)
	
 	process
	{
        $PR_ENTRYID = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0FFF,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)  
		if($Detail.IsPresent){
		$rpReport = ""
		foreach ($message in $MessageList){
			$PR_ENTRYIDValue = $null
            if($message.TryGetProperty($PR_ENTRYID,[ref]$PR_ENTRYIDValue)){
                $Oulookid = [System.BitConverter]::ToString($PR_ENTRYIDValue).Replace("-","")
            }	
			$fromstring = $message.From.Address
			if ($fromstring.length -gt 30){$fromstring = $fromstring.Substring(0,30)}
			$HeaderLine  = [DateTime]::Parse($message.DateTimeReceived).ToString("G") + " : " + $fromstring + " : " + $message.Subject
			$BodyLine = $message.Preview
			if($weblink.IsPresent){
				$BodyLine += "`r`n</br></br><a href=`"" + $message.WebClientReadFormQueryString + "`">MoreInfo</a href>"
			}else{
				$BodyLine += "`r`n</br></br><a href=`"outlook:" + $Oulookid + "`">MoreInfo</a href>"
			}
			
			$InfoField1Value = $message.$InfoField1Name
			$InfoField2Value = $message.$InfoField2Name
			$InfoField3Value = $message.$InfoField3Name
			$InfoField4Value = $message.$InfoField4Name
			$InfoField5Value = $message.$InfoField5Name
$nextTable = @"
<div style=" text-align: left; text-indent: 0px; padding: 0px 0px 0px 0px; margin: 0px 0px 0px 0px;">
<table width="100%" border="1" cellpadding="0" cellspacing="0" style="border-width: 0px; background-color: #ffffff;">
<tr valign="top">
<td colspan=5 style="border-width : 1px; border-color : #000000 #000000 #000000 #000000; border-style: solid;">
<p style=" text-align: center; text-indent: 0px; padding: 0px 0px 0px 0px; margin: 0px 0px 0px 0px; background-color: #3366ff;">
<span style=" font-size: 10pt; alignment-adjust: central; font-family: 'Arial', 'Helvetica', sans-serif; font-style: normal; font-weight: bold; color: #ffffff;  text-decoration: none;">
$HeaderLine</span></p>
</td>
</tr>
<tr valign="top">
<td width="20%" style="border-width : 1px; border-color : #000000 #000000 #000000 #000000; text-align: center; font-weight: bold; border-style: solid;">$InfoField1Name<br />
</td>
<td width="20%" style="border-width : 1px; border-color : #000000 #000000 #000000 #000000; text-align: center; font-weight: bold; border-style: solid;">$InfoField2Name<br />
</td>
<td width="20%" style="border-width : 1px; border-color : #000000 #000000 #000000 #000000; text-align: center; font-weight: bold; border-style: solid;">$InfoField3Name<br />
</td>
<td width="20%" style="border-width : 1px; border-color : #000000 #000000 #000000 #000000; text-align: center; font-weight: bold; border-style: solid;">$InfoField4Name<br />
</td>
<td width="20%" style="border-width : 1px; border-color : #000000 #000000 #000000 #000000; text-align: center; font-weight: bold; border-style: solid;">$InfoField5Name<br />
</td>
</tr>
<tr valign="top">
<td width="20%" style="border-width : 1px; border-color : #000000 #000000 #000000 #000000; text-align: center; border-style: solid;">$InfoField1Value<br />
</td>
<td width="20%" style="border-width : 1px; border-color : #000000 #000000 #000000 #000000; text-align: center; border-style: solid;">$InfoField2Value<br />
</td>
<td width="20%" style="border-width : 1px; border-color : #000000 #000000 #000000 #000000; text-align: center; border-style: solid;">$InfoField3Value<br />
</td>
<td width="20%" style="border-width : 1px; border-color : #000000 #000000 #000000 #000000; text-align: center; border-style: solid;">$InfoField4Value<br />
</td>
<td width="20%" style="border-width : 1px; border-color : #000000 #000000 #000000 #000000; text-align: center; border-style: solid;">$InfoField5Value<br />
</td>
</tr>
<td colspan=5 style="border-width : 1px; border-color : #000000 #000000 #000000 #000000; border-style: solid;">
<p style=" text-align: left; text-indent: 0px; padding: 0px 0px 0px 0px; margin: 0px 0px 0px 0px; >
<span style=" font-size: 10pt; alignment-adjust: central; font-family: 'Arial', 'Helvetica', sans-serif; font-style: normal; font-weight: bold; color: #ffffff;  text-decoration: none;">$BodyLine</span></p>
</td>
</table>
</div>
</br>
"@
			 $rpReport += $nextTable
		}
	}
	else{
		$rpReport = $rpReport + "<table><tr bgcolor=`"#95aedc`">" +"`r`n"
		$rpReport = $rpReport + "<td align=`"center`" style=`"width:15%;`" ><b>Recieved</b></td>" +"`r`n"
		$rpReport = $rpReport + "<td align=`"center`" style=`"width:20%;`" ><b>From</b></td>" +"`r`n"
		$rpReport = $rpReport + "<td align=`"center`" style=`"width:60%;`" ><b>Subject</b></td>" +"`r`n"
		$rpReport = $rpReport + "<td align=`"center`" style=`"width:5%;`" ><b>Size</b></td>" +"`r`n"
		$rpReport = $rpReport + "</tr>" + "`r`n"
		foreach ($message in $MessageList){
            $fromstring = $message.From.Address
            $PR_ENTRYIDValue = $null
            if($message.TryGetProperty($PR_ENTRYID,[ref]$PR_ENTRYIDValue)){
                $Oulookid = [System.BitConverter]::ToString($PR_ENTRYIDValue).Replace("-","")
            }			
			if ($fromstring.length -gt 30){$fromstring = $fromstring.Substring(0,30)}
			$rpReport = $rpReport + "  <tr>"  + "`r`n"
			$rpReport = $rpReport + "<td>" + [DateTime]::Parse($message.DateTimeReceived).ToString("G") + "</td>"  + "`r`n"
			$rpReport = $rpReport + "<td>" +  $fromstring + "</td>"  + "`r`n"
			if($weblink.IsPresent){
				$rpReport = $rpReport + "<td><a href=`"" + $message.WebClientReadFormQueryString + "`">" + $message.Subject + "</td>"  + "`r`n"
			}
			else{
				$rpReport = $rpReport + "<td><a href=`"outlook:" + $Oulookid + "`">" + $message.Subject + "</td>"  + "`r`n"
			}			
			$rpReport = $rpReport + "<td>" +  ($message.Size/1024).ToString(0.00) + "</td>"  + "`r`n"
			$rpReport = $rpReport + "</tr>"  + "`r`n"
		}
		$rpReport = $rpReport + "</table>"  + "  " 
		
	}
	return $rpReport
    }
}