function Invoke-HuntRcvdHeaders {
    [CmdletBinding()]
    param (
        [Parameter()]
        [String]
        $MailboxName,
        [Parameter()]
        [Int]
        $MessageCount=10,
        [Parameter()]
        [String]
        $Filter,
        [Parameter()]
        [string]
        $FolderId,
        [Parameter()]
        [string]
        $FolderPath
    )
    Process {
        $VersionHeaderRegex = ' id ([\s\S]*?) ' 
        $versionList = Get-ExchangeVersions
        if($FolderPath){
            $Folder = Get-MailBoxFolderFromPath -MailboxName $MailboxName -FolderPath $FolderPath 
            $FolderId = $Folder.Id
        }
        if($FolderId){
           $Messages = Get-MgUserMailFolderMessage -UserId $MailboxName -MailFolderId $FolderId -Top $MessageCount -Filter $Filter -Property InternetMessageHeaders,Subject,Id
        }else{
           $Messages = Get-MgUserMessage -UserId $MailboxName -Top $MessageCount -Filter $Filter -Property InternetMessageHeaders,Subject,Id
        }
        $Messages | ForEach-Object{
            Write-Verbose ("Process Message " + $_.Subject)
            $rcvHeaders = $_.InternetMessageHeaders | Where-Object -FilterScript {$_.Name -eq "Received"} 
            $hopCount =  $rcvHeaders.Count
            foreach($header in $rcvHeaders){
                $versionMatches = $header.Value | Select-String -Pattern $VersionHeaderRegex -AllMatches
                if($versionMatches){        
                    $hopCount--            
                    $version = $versionMatches.Matches[0].Groups[1].Value.Replace(";","")                    
                    if(!($version -contains " for")){
                        if($version -match "15.20"){
                            $exVersion = "Exchange Online"
                        }else{
                            $exVersion = $versionList -match $version
                        }                        
                        Write-Verbose ("Hop Count $hopCount : $version : $exVersion")
                    }
                }
            }
        }
    }
}

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

function Get-ExchangeVersions(){
$version = @"
Exchange Server 2019 CU12 Mar23SU,"March 14, 2023",15.2.1118.26,15.02.1118.026
Exchange Server 2019 CU12 Feb23SU,"February 14, 2023",15.2.1118.25,15.02.1118.025
Exchange Server 2019 CU12 Jan23SU,"January 10, 2023",15.2.1118.21,15.02.1118.021
Exchange Server 2019 CU12 Nov22SU,"November 8, 2022",15.2.1118.20,15.02.1118.020
Exchange Server 2019 CU12 Oct22SU,"October 11, 2022",15.2.1118.15,15.02.1118.015
Exchange Server 2019 CU12 Aug22SU,"August 9, 2022",15.2.1118.12,15.02.1118.012
Exchange Server 2019 CU12 May22SU,"May 10, 2022",15.2.1118.9,15.02.1118.009
Exchange Server 2019 CU12 (2022H1),"April 20, 2022",15.2.1118.7,15.02.1118.007
Exchange Server 2019 CU11 Mar23SU,"March 14, 2023",15.2.986.42,15.02.0986.042
Exchange Server 2019 CU11 Feb23SU,"February 14, 2023",15.2.986.41,15.02.0986.041
Exchange Server 2019 CU11 Jan23SU,"January 10, 2023",15.2.986.37,15.02.0986.037
Exchange Server 2019 CU11 Nov22SU,"November 8, 2022",15.2.986.36,15.02.0986.036
Exchange Server 2019 CU11 Oct22SU,"October 11, 2022",15.2.986.30,15.02.0986.030
Exchange Server 2019 CU11 Aug22SU,"August 9, 2022",15.2.986.29,15.02.0986.029
Exchange Server 2019 CU11 May22SU,"May 10, 2022",15.2.986.26,15.02.0986.026
Exchange Server 2019 CU11 Mar22SU,"March 8, 2022",15.2.986.22,15.02.0986.022
Exchange Server 2019 CU11 Jan22SU,"January 11, 2022",15.2.986.15,15.02.0986.015
Exchange Server 2019 CU11 Nov21SU,"November 9, 2021",15.2.986.14,15.02.0986.014
Exchange Server 2019 CU11 Oct21SU,"October 12, 2021",15.2.986.9,15.02.0986.009
Exchange Server 2019 CU11,"September 28, 2021",15.2.986.5,15.02.0986.005
Exchange Server 2019 CU10 Mar22SU,"March 8, 2022",15.2.922.27,15.02.0922.027
Exchange Server 2019 CU10 Jan22SU,"January 11, 2022",15.2.922.20,15.02.0922.020
Exchange Server 2019 CU10 Nov21SU,"November 9, 2021",15.2.922.19,15.02.0922.019
Exchange Server 2019 CU10 Oct21SU,"October 12, 2021",15.2.922.14,15.02.0922.014
Exchange Server 2019 CU10 Jul21SU,"July 13, 2021",15.2.922.13,15.02.0922.013
Exchange Server 2019 CU10,"June 29, 2021",15.2.922.7,15.02.0922.007
Exchange Server 2019 CU9 Jul21SU,"July 13, 2021",15.2.858.15,15.02.0858.015
Exchange Server 2019 CU9 May21SU,"May 11, 2021",15.2.858.12,15.02.0858.012
Exchange Server 2019 CU9 Apr21SU,"April 13, 2021",15.2.858.10,15.02.0858.010
Exchange Server 2019 CU9,"March 16, 2021",15.2.858.5,15.02.0858.005
Exchange Server 2019 CU8 May21SU,"May 11, 2021",15.2.792.15,15.02.0792.015
Exchange Server 2019 CU8 Apr21SU,"April 13, 2021",15.2.792.13,15.02.0792.013
Exchange Server 2019 CU8 Mar21SU,"March 2, 2021",15.2.792.10,15.02.0792.010
Exchange Server 2019 CU8,"December 15, 2020",15.2.792.3,15.02.0792.003
Exchange Server 2019 CU7 Mar21SU,"March 2, 2021",15.2.721.13,15.02.0721.013
Exchange Server 2019 CU7,"September 15, 2020",15.2.721.2,15.02.0721.002
Exchange Server 2019 CU6 Mar21SU,"March 2, 2021",15.2.659.12,15.02.0659.012
Exchange Server 2019 CU6,"June 16, 2020",15.2.659.4,15.02.0659.004
Exchange Server 2019 CU5 Mar21SU,"March 2, 2021",15.2.595.8,15.02.0595.008
Exchange Server 2019 CU5,"March 17, 2020",15.2.595.3,15.02.0595.003
Exchange Server 2019 CU4 Mar21SU,"March 2, 2021",15.2.529.13,15.02.0529.013
Exchange Server 2019 CU4,"December 17, 2019",15.2.529.5,15.02.0529.005
Exchange Server 2019 CU3 Mar21SU,"March 2, 2021",15.2.464.15,15.02.0464.015
Exchange Server 2019 CU3,"September 17, 2019",15.2.464.5,15.02.0464.005
Exchange Server 2019 CU2 Mar21SU,"March 2, 2021",15.2.397.11,15.02.0397.011
Exchange Server 2019 CU2,"June 18, 2019",15.2.397.3,15.02.0397.003
Exchange Server 2019 CU1 Mar21SU,"March 2, 2021",15.2.330.11,15.02.0330.011
Exchange Server 2019 CU1,"February 12, 2019",15.2.330.5,15.02.0330.005
Exchange Server 2019 RTM Mar21SU,"March 2, 2021",15.2.221.18,15.02.0221.018
Exchange Server 2019 RTM,"October 22, 2018",15.2.221.12,15.02.0221.012
Exchange Server 2019 Preview,"July 24, 2018",15.2.196.0,15.02.0196.000
Exchange Server 2016 CU23 Mar23SU,"March 14, 2023",15.1.2507.23,15.01.2507.023
Exchange Server 2016 CU23 Feb23SU,"February 14, 2023",15.1.2507.21,15.01.2507.021
Exchange Server 2016 CU23 Jan23SU,"January 10, 2023",15.1.2507.17,15.01.2507.017
Exchange Server 2016 CU23 Nov22SU,"November 8, 2022",15.1.2507.16,15.01.2507.016
Exchange Server 2016 CU23 Oct22SU,"October 11, 2022",15.1.2507.13,15.01.2507.013
Exchange Server 2016 CU23 Aug22SU,"August 9, 2022",15.1.2507.12,15.01.2507.012
Exchange Server 2016 CU23 May22SU,"May 10, 2022",15.1.2507.9,15.01.2507.009
Exchange Server 2016 CU23 (2022H1),"April 20, 2022",15.1.2507.6,15.01.2507.006
Exchange Server 2016 CU22 Nov22SU,"November 8, 2022",15.1.2375.37,15.01.2375.037
Exchange Server 2016 CU22 Oct22SU,"October 11, 2022",15.1.2375.32,15.01.2375.032
Exchange Server 2016 CU22 Aug22SU,"August 9, 2022",15.1.2375.31,15.01.2375.031
Exchange Server 2016 CU22 May22SU,"May 10, 2022",15.1.2375.28,15.01.2375.028
Exchange Server 2016 CU22 Mar22SU,"March 8, 2022",15.1.2375.24,15.01.2375.024
Exchange Server 2016 CU22 Jan22SU,"January 11, 2022",15.1.2375.18,15.01.2375.018
Exchange Server 2016 CU22 Nov21SU,"November 9, 2021",15.1.2375.17,15.01.2375.017
Exchange Server 2016 CU22 Oct21SU,"October 12, 2021",15.1.2375.12,15.01.2375.012
Exchange Server 2016 CU22,"September 28, 2021",15.1.2375.7,15.01.2375.007
Exchange Server 2016 CU21 Mar22SU,"March 8, 2022",15.1.2308.27,15.01.2308.027
Exchange Server 2016 CU21 Jan22SU,"January 11, 2022",15.1.2308.21,15.01.2308.021
Exchange Server 2016 CU21 Nov21SU,"November 9, 2021",15.1.2308.20,15.01.2308.020
Exchange Server 2016 CU21 Oct21SU,"October 12, 2021",15.1.2308.15,15.01.2308.015
Exchange Server 2016 CU21 Jul21SU,"July 13, 2021",15.1.2308.14,15.01.2308.014
Exchange Server 2016 CU21,"June 29, 2021",15.1.2308.8,15.01.2308.008
Exchange Server 2016 CU20 Jul21SU,"July 13, 2021",15.1.2242.12,15.01.2242.012
Exchange Server 2016 CU20 May21SU,"May 11, 2021",15.1.2242.10,15.01.2242.010
Exchange Server 2016 CU20 Apr21SU,"April 13, 2021",15.1.2242.8,15.01.2242.008
Exchange Server 2016 CU20,"March 16, 2021",15.1.2242.4,15.01.2242.004
Exchange Server 2016 CU19 May21SU,"May 11, 2021",15.1.2176.14,15.01.2176.014
Exchange Server 2016 CU19 Apr21SU,"April 13, 2021",15.1.2176.12,15.01.2176.012
Exchange Server 2016 CU19 Mar21SU,"March 2, 2021",15.1.2176.9,15.01.2176.009
Exchange Server 2016 CU19,"December 15, 2020",15.1.2176.2,15.01.2176.002
Exchange Server 2016 CU18 Mar21SU,"March 2, 2021",15.1.2106.13,15.01.2106.013
Exchange Server 2016 CU18,"September 15, 2020",15.1.2106.2,15.01.2106.002
Exchange Server 2016 CU17 Mar21SU,"March 2, 2021",15.1.2044.13,15.01.2044.013
Exchange Server 2016 CU17,"June 16, 2020",15.1.2044.4,15.01.2044.004
Exchange Server 2016 CU16 Mar21SU,"March 2, 2021",15.1.1979.8,15.01.1979.008
Exchange Server 2016 CU16,"March 17, 2020",15.1.1979.3,15.01.1979.003
Exchange Server 2016 CU15 Mar21SU,"March 2, 2021",15.1.1913.12,15.01.1913.012
Exchange Server 2016 CU15,"December 17, 2019",15.1.1913.5,15.01.1913.005
Exchange Server 2016 CU14 Mar21SU,"March 2, 2021",15.1.1847.12,15.01.1847.012
Exchange Server 2016 CU14,"September 17, 2019",15.1.1847.3,15.01.1847.003
Exchange Server 2016 CU13 Mar21SU,"March 2, 2021",15.1.1779.8,15.01.1779.008
Exchange Server 2016 CU13,"June 18, 2019",15.1.1779.2,15.01.1779.002
Exchange Server 2016 CU12 Mar21SU,"March 2, 2021",15.1.1713.10,15.01.1713.010
Exchange Server 2016 CU12,"February 12, 2019",15.1.1713.5,15.01.1713.005
Exchange Server 2016 CU11 Mar21SU,"March 2, 2021",15.1.1591.18,15.01.1591.018
Exchange Server 2016 CU11,"October 16, 2018",15.1.1591.10,15.01.1591.010
Exchange Server 2016 CU10 Mar21SU,"March 2, 2021",15.1.1531.12,15.01.1531.012
Exchange Server 2016 CU10,"June 19, 2018",15.1.1531.3,15.01.1531.003
Exchange Server 2016 CU9 Mar21SU,"March 2, 2021",15.1.1466.16,15.01.1466.016
Exchange Server 2016 CU9,"March 20, 2018",15.1.1466.3,15.01.1466.003
Exchange Server 2016 CU8 Mar21SU,"March 2, 2021",15.1.1415.10,15.01.1415.010
Exchange Server 2016 CU8,"December 19, 2017",15.1.1415.2,15.01.1415.002
Exchange Server 2016 CU7,"September 19, 2017",15.1.1261.35,15.01.1261.035
Exchange Server 2016 CU6,"June 27, 2017",15.1.1034.26,15.01.1034.026
Exchange Server 2016 CU5,"March 21, 2017",15.1.845.34,15.01.0845.034
Exchange Server 2016 CU4,"December 13, 2016",15.1.669.32,15.01.0669.032
Exchange Server 2016 CU3,"September 20, 2016",15.1.544.27,15.01.0544.027
Exchange Server 2016 CU2,"June 21, 2016",15.1.466.34,15.01.0466.034
Exchange Server 2016 CU1,"March 15, 2016",15.1.396.30,15.01.0396.030
Exchange Server 2016 RTM,"October 1, 2015",15.1.225.42,15.01.0225.042
Exchange Server 2016 Preview,"July 22, 2015",15.1.225.16,15.01.0225.016
Exchange Server 2013 CU23 Mar23SU,"March 14, 2023",15.0.1497.48,15.00.1497.048
Exchange Server 2013 CU23 Feb23SU,"February 14, 2023",15.0.1497.47,15.00.1497.047
Exchange Server 2013 CU23 Jan23SU,"January 10, 2023",15.0.1497.45,15.00.1497.045
Exchange Server 2013 CU23 Nov22SU,"November 8, 2022",15.0.1497.44,15.00.1497.044
Exchange Server 2013 CU23 Oct22SU,"October 11, 2022",15.0.1497.42,15.00.1497.042
Exchange Server 2013 CU23 Aug22SU,"August 9, 2022",15.0.1497.40,15.00.1497.040
Exchange Server 2013 CU23 May22SU,"May 10, 2022",15.0.1497.36,15.00.1497.036
Exchange Server 2013 CU23 Mar22SU,"March 8, 2022",15.0.1497.33,15.00.1497.033
Exchange Server 2013 CU23 Jan22SU,"January 11, 2022",15.0.1497.28,15.00.1497.028
Exchange Server 2013 CU23 Nov21SU,"November 9, 2021",15.0.1497.26,15.00.1497.026
Exchange Server 2013 CU23 Oct21SU,"October 12, 2021",15.0.1497.24,15.00.1497.024
Exchange Server 2013 CU23 Jul21SU,"July 13, 2021",15.0.1497.23,15.00.1497.023
Exchange Server 2013 CU23 May21SU,"May 11, 2021",15.0.1497.18,15.00.1497.018
Exchange Server 2013 CU23 Apr21SU,"April 13, 2021",15.0.1497.15,15.00.1497.015
Exchange Server 2013 CU23 Mar21SU,"March 2, 2021",15.0.1497.12,15.00.1497.012
Exchange Server 2013 CU23,"June 18, 2019",15.0.1497.2,15.00.1497.002
Exchange Server 2013 CU22 Mar21SU,"March 2, 2021",15.0.1473.6,15.00.1473.006
Exchange Server 2013 CU22,"February 12, 2019",15.0.1473.3,15.00.1473.003
Exchange Server 2013 CU21 Mar21SU,"March 2, 2021",15.0.1395.12,15.00.1395.012
Exchange Server 2013 CU21,"June 19, 2018",15.0.1395.4,15.00.1395.004
Exchange Server 2013 CU20,"March 20, 2018",15.0.1367.3,15.00.1367.003
Exchange Server 2013 CU19,"December 19, 2017",15.0.1365.1,15.00.1365.001
Exchange Server 2013 CU18,"September 19, 2017",15.0.1347.2,15.00.1347.002
Exchange Server 2013 CU17,"June 27, 2017",15.0.1320.4,15.00.1320.004
Exchange Server 2013 CU16,"March 21, 2017",15.0.1293.2,15.00.1293.002
Exchange Server 2013 CU15,"December 13, 2016",15.0.1263.5,15.00.1263.005
Exchange Server 2013 CU14,"September 20, 2016",15.0.1236.3,15.00.1236.003
Exchange Server 2013 CU13,"June 21, 2016",15.0.1210.3,15.00.1210.003
Exchange Server 2013 CU12,"March 15, 2016",15.0.1178.4,15.00.1178.004
Exchange Server 2013 CU11,"December 15, 2015",15.0.1156.6,15.00.1156.006
Exchange Server 2013 CU10,"September 15, 2015",15.0.1130.7,15.00.1130.007
Exchange Server 2013 CU9,"June 17, 2015",15.0.1104.5,15.00.1104.005
Exchange Server 2013 CU8,"March 17, 2015",15.0.1076.9,15.00.1076.009
Exchange Server 2013 CU7,"December 9, 2014",15.0.1044.25,15.00.1044.025
Exchange Server 2013 CU6,"August 26, 2014",15.0.995.29,15.00.0995.029
Exchange Server 2013 CU5,"May 27, 2014",15.0.913.22,15.00.0913.022
Exchange Server 2013 SP1 Mar21SU,"March 2, 2021",15.0.847.64,15.00.0847.064
Exchange Server 2013 SP1,"February 25, 2014",15.0.847.32,15.00.0847.032
Exchange Server 2013 CU3,"November 25, 2013",15.0.775.38,15.00.0775.038
Exchange Server 2013 CU2,"July 9, 2013",15.0.712.24,15.00.0712.024
Exchange Server 2013 CU1,"April 2, 2013",15.0.620.29,15.00.0620.029
Exchange Server 2013 RTM,"December 3, 2012",15.0.516.32,15.00.0516.032
Update Rollup 32 for Exchange Server 2010 SP3,"March 2, 2021",14.3.513.0,14.03.0513.000
Update Rollup 31 for Exchange Server 2010 SP3,"December 1, 2020",14.3.509.0,14.03.0509.000
Update Rollup 30 for Exchange Server 2010 SP3,"February 11, 2020",14.3.496.0,14.03.0496.000
Update Rollup 29 for Exchange Server 2010 SP3,"July 9, 2019",14.3.468.0,14.03.0468.000
Update Rollup 28 for Exchange Server 2010 SP3,"June 7, 2019",14.3.461.1,14.03.0461.001
Update Rollup 27 for Exchange Server 2010 SP3,"April 9, 2019",14.3.452.0,14.03.0452.000
Update Rollup 26 for Exchange Server 2010 SP3,"February 12, 2019",14.3.442.0,14.03.0442.000
Update Rollup 25 for Exchange Server 2010 SP3,"January 8, 2019",14.3.435.0,14.03.0435.000
Update Rollup 24 for Exchange Server 2010 SP3,"September 5, 2018",14.3.419.0,14.03.0419.000
Update Rollup 23 for Exchange Server 2010 SP3,"August 13, 2018",14.3.417.1,14.03.0417.001
Update Rollup 22 for Exchange Server 2010 SP3,"June 19, 2018",14.3.411.0,14.03.0411.000
Update Rollup 21 for Exchange Server 2010 SP3,"May 7, 2018",14.3.399.2,14.03.0399.002
Update Rollup 20 for Exchange Server 2010 SP3,"March 5, 2018",14.3.389.1,14.03.0389.001
Update Rollup 19 for Exchange Server 2010 SP3,"December 19, 2017",14.3.382.0,14.03.0382.000
Update Rollup 18 for Exchange Server 2010 SP3,"July 11, 2017",14.3.361.1,14.03.0361.001
Update Rollup 17 for Exchange Server 2010 SP3,"March 21, 2017",14.3.352.0,14.03.0352.000
Update Rollup 16 for Exchange Server 2010 SP3,"December 13, 2016",14.3.336.0,14.03.0336.000
Update Rollup 15 for Exchange Server 2010 SP3,"September 20, 2016",14.3.319.2,14.03.0319.002
Update Rollup 14 for Exchange Server 2010 SP3,"June 21, 2016",14.3.301.0,14.03.0301.000
Update Rollup 13 for Exchange Server 2010 SP3,"March 15, 2016",14.3.294.0,14.03.0294.000
Update Rollup 12 for Exchange Server 2010 SP3,"December 15, 2015",14.3.279.2,14.03.0279.002
Update Rollup 11 for Exchange Server 2010 SP3,"September 15, 2015",14.3.266.2,14.03.0266.002
Update Rollup 10 for Exchange Server 2010 SP3,"June 17, 2015",14.3.248.2,14.03.0248.002
Update Rollup 9 for Exchange Server 2010 SP3,"March 17, 2015",14.3.235.1,14.03.0235.001
Update Rollup 8 v2 for Exchange Server 2010 SP3,"December 12, 2014",14.3.224.2,14.03.0224.002
Update Rollup 8 v1 for Exchange Server 2010 SP3 (recalled),"December 9, 2014",14.3.224.1,14.03.0224.001
Update Rollup 7 for Exchange Server 2010 SP3,"August 26, 2014",14.3.210.2,14.03.0210.002
Update Rollup 6 for Exchange Server 2010 SP3,"May 27, 2014",14.3.195.1,14.03.0195.001
Update Rollup 5 for Exchange Server 2010 SP3,"February 24, 2014",14.3.181.6,14.03.0181.006
Update Rollup 4 for Exchange Server 2010 SP3,"December 9, 2013",14.3.174.1,14.03.0174.001
Update Rollup 3 for Exchange Server 2010 SP3,"November 25, 2013",14.3.169.1,14.03.0169.001
Update Rollup 2 for Exchange Server 2010 SP3,"August 8, 2013",14.3.158.1,14.03.0158.001
Update Rollup 1 for Exchange Server 2010 SP3,"May 29, 2013",14.3.146.0,14.03.0146.000
Exchange Server 2010 SP3,"February 12, 2013",14.3.123.4,14.03.0123.004
Update Rollup 8 for Exchange Server 2010 SP2,"December 9, 2013",14.2.390.3,14.02.0390.003
Update Rollup 7 for Exchange Server 2010 SP2,"August 3, 2013",14.2.375.0,14.02.0375.000
Update Rollup 6 Exchange Server 2010 SP2,"February 12, 2013",14.2.342.3,14.02.0342.003
Update Rollup 5 v2 for Exchange Server 2010 SP2,"December 10, 2012",14.2.328.10,14.02.0328.010
Update Rollup 5 for Exchange Server 2010 SP2,"November 13, 2012",14.3.328.5,14.03.0328.005
Update Rollup 4 v2 for Exchange Server 2010 SP2,"October 9, 2012",14.2.318.4,14.02.0318.004
Update Rollup 4 for Exchange Server 2010 SP2,"August 13, 2012",14.2.318.2,14.02.0318.002
Update Rollup 3 for Exchange Server 2010 SP2,"May 29, 2012",14.2.309.2,14.02.0309.002
Update Rollup 2 for Exchange Server 2010 SP2,"April 16, 2012",14.2.298.4,14.02.0298.004
Update Rollup 1 for Exchange Server 2010 SP2,"February 13, 2012",14.2.283.3,14.02.0283.003
Exchange Server 2010 SP2,"December 4, 2011",14.2.247.5,14.02.0247.005
Update Rollup 8 for Exchange Server 2010 SP1,"December 10, 2012",14.1.438.0,14.01.0438.000
Update Rollup 7 v3 for Exchange Server 2010 SP1,"November 13, 2012",14.1.421.3,14.01.0421.003
Update Rollup 7 v2 for Exchange Server 2010 SP1,"October 10, 2012",14.1.421.2,14.01.0421.002
Update Rollup 7 for Exchange Server 2010 SP1,"August 8, 2012",14.1.421.0,14.01.0421.000
Update Rollup 6 for Exchange Server 2010 SP1,"October 27, 2011",14.1.355.2,14.01.0355.002
Update Rollup 5 for Exchange Server 2010 SP1,"August 23, 2011",14.1.339.1,14.01.0339.001
Update Rollup 4 for Exchange Server 2010 SP1,"July 27, 2011",14.1.323.6,14.01.0323.006
Update Rollup 3 for Exchange Server 2010 SP1,"April 6, 2011",14.1.289.7,14.01.0289.007
Update Rollup 2 for Exchange Server 2010 SP1,"December 9, 2010",14.1.270.1,14.01.0270.001
Update Rollup 1 for Exchange Server 2010 SP1,"October 4, 2010",14.1.255.2,14.01.0255.002
Exchange Server 2010 SP1,"August 23, 2010",14.1.218.15,14.01.0218.015
Update Rollup 5 for Exchange Server 2010,"December 13, 2010",14.0.726.0,14.00.0726.000
Update Rollup 4 for Exchange Server 2010,"June 10, 2010",14.0.702.1,14.00.0702.001
Update Rollup 3 for Exchange Server 2010,"April 13, 2010",14.0.694.0,14.00.0694.000
Update Rollup 2 for Exchange Server 2010,"March 4, 2010",14.0.689.0,14.00.0689.000
Update Rollup 1 for Exchange Server 2010,"December 9, 2009",14.0.682.1,14.00.0682.001
Exchange Server 2010 RTM,"November 9, 2009",14.0.639.21,14.00.0639.021
Update Rollup 23 for Exchange Server 2007 SP3,"March 21, 2017",8.3.517.0,8.03.0517.000
Update Rollup 22 for Exchange Server 2007 SP3,"December 13, 2016",8.3.502.0,8.03.0502.000
Update Rollup 21 for Exchange Server 2007 SP3,"September 20, 2016",8.3.485.1,8.03.0485.001
Update Rollup 20 for Exchange Server 2007 SP3,"June 21, 2016",8.3.468.0,8.03.0468.000
Update Rollup 19 forExchange Server 2007 SP3,"March 15, 2016",8.3.459.0,8.03.0459.000
Update Rollup 18 forExchange Server 2007 SP3,"December, 2015",8.3.445.0,8.03.0445.000
Update Rollup 17 forExchange Server 2007 SP3,"June 17, 2015",8.3.417.1,8.03.0417.001
Update Rollup 16 for Exchange Server 2007 SP3,"March 17, 2015",8.3.406.0,8.03.0406.000
Update Rollup 15 for Exchange Server 2007 SP3,"December 9, 2014",8.3.389.2,8.03.0389.002
Update Rollup 14 for Exchange Server 2007 SP3,"August 26, 2014",8.3.379.2,8.03.0379.002
Update Rollup 13 for Exchange Server 2007 SP3,"February 24, 2014",8.3.348.2,8.03.0348.002
Update Rollup 12 for Exchange Server 2007 SP3,"December 9, 2013",8.3.342.4,8.03.0342.004
Update Rollup 11 for Exchange Server 2007 SP3,"August 13, 2013",8.3.327.1,8.03.0327.001
Update Rollup 10 for Exchange Server 2007 SP3,"February 11, 2013",8.3.298.3,8.03.0298.003
Update Rollup 9 for Exchange Server 2007 SP3,"December 10, 2012",8.3.297.2,8.03.0297.002
Update Rollup 8-v3 for Exchange Server 2007 SP3,"November 13, 2012",8.3.279.6,8.03.0279.006
Update Rollup 8-v2 for Exchange Server 2007 SP3,"October 9, 2012",8.3.279.5,8.03.0279.005
Update Rollup 8 for Exchange Server 2007 SP3,"August 13, 2012",8.3.279.3,8.03.0279.003
Update Rollup 7 for Exchange Server 2007 SP3,"April 16, 2012",8.3.264.0,8.03.0264.000
Update Rollup 6 for Exchange Server 2007 SP3,"January 26, 2012",8.3.245.2,8.03.0245.002
Update Rollup 5 for Exchange Server 2007 SP3,"September 21, 2011",8.3.213.1,8.03.0213.001
Update Rollup 4 for Exchange Server 2007 SP3,"May 28, 2011",8.3.192.1,8.03.0192.001
Update Rollup 3-v2 for Exchange Server 2007 SP3,"March 30, 2011",8.3.159.2,8.03.0159.002
Update Rollup 2 for Exchange Server 2007 SP3,"December 10, 2010",8.3.137.3,8.03.0137.003
Update Rollup 1 for Exchange Server 2007 SP3,"September 9, 2010",8.3.106.2,8.03.0106.002
Exchange Server 2007 SP3,"June 7, 2010",8.3.83.6,8.03.0083.006
Update Rollup 5 for Exchange Server 2007 SP2,"December 7, 2010",8.2.305.3,8.02.0305.003
Update Rollup 4 for Exchange Server 2007 SP2,"April 9, 2010",8.2.254.0,8.02.0254.000
Update Rollup 3 for Exchange Server 2007 SP2,"March 17, 2010",8.2.247.2,8.02.0247.002
Update Rollup 2 for Exchange Server 2007 SP2,"January 22, 2010",8.2.234.1,8.02.0234.001
Update Rollup 1 for Exchange Server 2007 SP2,"November 19, 2009",8.2.217.3,8.02.0217.003
Exchange Server 2007 SP2,"August 24, 2009",8.2.176.2,8.02.0176.002
Update Rollup 10 for Exchange Server 2007 SP1,"April 13, 2010",8.1.436.0,8.01.0436.000
Update Rollup 9 for Exchange Server 2007 SP1,"July 16, 2009",8.1.393.1,8.01.0393.001
Update Rollup 8 for Exchange Server 2007 SP1,"May 19, 2009",8.1.375.2,8.01.0375.002
Update Rollup 7 for Exchange Server 2007 SP1,"March 18, 2009",8.1.359.2,8.01.0359.002
Update Rollup 6 for Exchange Server 2007 SP1,"February 10, 2009",8.1.340.1,8.01.0340.001
Update Rollup 5 for Exchange Server 2007 SP1,"November 20, 2008",8.1.336.1,8.01.0336.01
Update Rollup 4 for Exchange Server 2007 SP1,"October 7, 2008",8.1.311.3,8.01.0311.003
Update Rollup 3 for Exchange Server 2007 SP1,"July 8, 2008",8.1.291.2,8.01.0291.002
Update Rollup 2 for Exchange Server 2007 SP1,"May 9, 2008",8.1.278.2,8.01.0278.002
Update Rollup 1 for Exchange Server 2007 SP1,"February 28, 2008",8.1.263.1,8.01.0263.001
Exchange Server 2007 SP1,"November 29, 2007",8.1.240.6,8.01.0240.006
Update Rollup 7 for Exchange Server 2007,"July 8, 2008",8.0.813.0,8.00.0813.000
Update Rollup 6 for Exchange Server 2007,"February 21, 2008",8.0.783.2,8.00.0783.002
Update Rollup 5 for Exchange Server 2007,"October 25, 2007",8.0.754.0,8.00.0754.000
Update Rollup 4 for Exchange Server 2007,"August 23, 2007",8.0.744.0,8.00.0744.000
Update Rollup 3 for Exchange Server 2007,"June 28, 2007",8.0.730.1,8.00.0730.001
Update Rollup 2 for Exchange Server 2007,"May 8, 2007",8.0.711.2,8.00.0711.002
Update Rollup 1 for Exchange Server 2007,"April 17, 2007",8.0.708.3,8.00.0708.003
Exchange Server 2007 RTM,"March 8, 2007",8.0.685.25,8.00.0685.025
Exchange Server 2003 post-SP2,Aug-08,6.5.7654.4,
Exchange Server 2003 post-SP2,Mar-08,6.5.7653.33,
Exchange Server 2003 SP2,"October 19, 2005",6.5.7683,
Exchange Server 2003 SP1,"May25, 2004",6.5.7226,
Exchange Server 2003,"September 28, 2003",6.5.6944,
Exchange 2000 Server post-SP3,Aug-08,6.0.6620.7,
Exchange 2000 Server post-SP3,Mar-08,6.0.6620.5,
Exchange 2000 Server post-SP3,Aug-04,6.0.6603,
Exchange 2000 Server post-SP3,Apr-04,6.0.6556,
Exchange 2000 Server post-SP3,Sep-03,6.0.6487,
Exchange 2000 Server SP3,"July 18, 2002",6.0.6249,
Exchange 2000 Server SP2,"November 29, 2001",6.0.5762,
Exchange 2000 Server SP1,"June 21, 2001",6.0.4712,
Exchange 2000 Server,"November 29, 2000",6.0.4417,
Exchange Server version 5.5 SP4,"November 1, 2000",5.5.2653,
Exchange Server version 5.5 SP3,"September 9, 1999",5.5.2650,
Exchange Server version 5.5 SP2,"December 23, 1998",5.5.2448,
Exchange Server version 5.5 SP1,"August 5, 1998",5.5.2232,
Exchange Server version 5.5,"February 3, 1998",5.5.1960,
Exchange Server 5.0 SP2,"February 19, 1998",5.0.1460,
Exchange Server 5.0 SP1,"June 18, 1997",5.0.1458,
Exchange Server 5.0,"May 23, 1997",5.0.1457,
Exchange Server 4.0 SP5,"May 5, 1998",4.0.996,
Exchange Server 4.0 SP4,"March 28, 1997",4.0.995,
Exchange Server 4.0 SP3,"October 29, 1996",4.0.994,
Exchange Server 4.0 SP2,"July 19, 1996",4.0.993,
Exchange Server 4.0 SP1,"May 1, 1996",4.0.838,
Exchange Server 4.0 Standard Edition,"June 11, 1996",4.0.837,
"@ 
return convertfrom-csv -InputObject $version -Header ExchangeVersion,Date,ShortVersion,LongVersion  
}