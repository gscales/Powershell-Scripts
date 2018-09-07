

$mbHash = @{ }
$mbHash.Add("gscales@datarumble.com", "Glen Scales")
$mbHash.Add("jcool@datarumble.com", "Jcool")
$mbHash.Add("mec@datarumble.com", "mec")

$tmValHash = @{ }
$tidx = 0
for ($vsStartTime = [DateTime]::Parse([DateTime]::Now.ToString("yyyy-MM-dd 0:00")); $vsStartTime -lt [DateTime]::Parse([DateTime]::Now.ToString("yyyy-MM-dd 0:00")).AddDays(1); $vsStartTime = $vsStartTime.AddMinutes(30)) {
    $tmValHash.add($vsStartTime.ToString("HH:mm"), $tidx)	
    $tidx++
}



$usrIdx = 0
$frow = $true

$displayStartTime = [DateTime]::Parse([DateTime]::Now.ToString("yyyy-MM-dd 08:30"))
$displayEndTime = [DateTime]::Parse([DateTime]::Now.ToString("yyyy-MM-dd 17:30"))

$avResults = Get-EXRSchedule -Mailboxes $mbHash.Keys -Start $displayStartTime -EndTime $displayEndTime  -availabilityViewInterval 30

foreach ($fbResult in $avResults) {
    if ($frow -eq $true) {
        $fbBoard = $fbBoard + "<table><tr bgcolor=`"#95aedc`">" + "`r`n"
        $fbBoard = $fbBoard + "<td align=`"center`" style=`"width=200;`" ><b>User</b></td>" + "`r`n"
        for ($stime = $displayStartTime; $stime -lt $displayEndTime; $stime = $stime.AddMinutes(30)) {
            $fbBoard = $fbBoard + "<td align=`"center`" style=`"width=50;`" ><b>" + $stime.ToString("HH:mm") + "</b></td>" + "`r`n"
        }
        $fbBoard = $fbBoard + "</tr>" + "`r`n"
        $frow = $false
    }
    $aphash = @{}
    foreach($item in $fbResult.scheduleItems){
        $key = ([DateTime]::Parse($item.start.dateTime).ToString("HH:mm"))
        if(!$aphash.ContainsKey($key)){
            $aphash.Add($key,$item.Subject)
        }
    }
    $aphash
    Write-Host ""
    if($mbHash.ContainsKey($fbResult.scheduleId)){
        $fbBoard = $fbBoard + "<td bgcolor=`"#CFECEC`"><b>" + $mbHash[$fbResult.scheduleId] + "</b></td>" + "`r`n"
    }else{
        $fbBoard = $fbBoard + "<td bgcolor=`"#CFECEC`"><b>" + $fbResult.scheduleId + "</b></td>" + "`r`n"
    }    
    $CurrentTime = $displayStartTime  
    $fbResult.availabilityView.ToCharArray() | ForEach-Object {
        $title = "title="
        switch ($_) {
            0 {$bgColour = "bgcolor=`"#41A317`""}
            1 {$bgColour = "bgcolor=`"#52F3FF`""}
            2 {$bgColour = "bgcolor=`"#153E7E`""}
            3 {$bgColour = "bgcolor=`"#4E387E`""}
            4 {$bgColour = "bgcolor=`"#98AFC7`""}
        }
        $key = $CurrentTime.ToString("HH:mm");
        if($aphash.ContainsKey($key)){
            $title = "title=" + $aphash[$key]
        }
        if ($title -ne "title=") {
            $fbBoard = $fbBoard + "<td " + $bgColour + " " + $title + "></td>" + "`r`n"
        }
        else {
            $fbBoard = $fbBoard + "<td " + $bgColour + "></td>" + "`r`n"
        }
        $CurrentTime = $CurrentTime.AddMinutes(30)
    }

    $fbBoard = $fbBoard + "</tr>" + "`r`n"
    $usrIdx++
}
$fbBoard = $fbBoard + "</table>" + "  " 
$fbBoard | out-file "c:\temp\fbboard.htm"


