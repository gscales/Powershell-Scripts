function Get-MailboxQuotaIf {
    [CmdletBinding()]
    param (   
        [Parameter(Position = 1, Mandatory = $true)]
        [Int32]
        $QuotaIfVal,
        [Parameter(Position = 2, Mandatory = $false)]
        [String]
        $OutputFile
    )
    Process {
        Connect-MgGraph -Scopes "Reports.Read.All" | Out-Null
        $TempFileName = [Guid]::NewGuid().ToString() + ".csv"
	    Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='D7')" -OutputFilePath ./$TempFileName
        $MbReport = Import-csv ./$TempFileName
        $rptCollection = @()  
        foreach($Mailbox in $MbReport){
            $rptObj = "" | Select MailboxName,TotalSize,QuotaIfPercent,PercentGraph  
            $rptObj.MailboxName = $Mailbox.'Display Name'
            [Int64]$rptObj.TotalSize = $Mailbox.'Storage Used (Byte)'/1MB  
          
            $rptObj.QuotaIfPercent = 0    
            if($rptObj.TotalSize -gt 0){  
                $rptObj.QuotaIfPercent = [Math]::round((($rptObj.TotalSize/$QuotaIfVal) * 100))   
            }  
            $PercentGraph = ""  
            for($intval=0;$intval -lt 100;$intval+=4){
                if($rptObj.QuotaIfPercent -gt $intval){
                    $PercentGraph += [char]0x2593
                }
                else{		
                    $PercentGraph += [char]0x2591
                }
            }
            $rptObj.PercentGraph = $PercentGraph 
            $rptCollection +=$rptObj   
        }           
        Remove-Item -Path ./$TempFileName
        return, $rptCollection
    }
}   