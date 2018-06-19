function New-AdaptiveCard { 
    param( 
        [Parameter(Position = 0, Mandatory = $false)] [psObject]$Columns,
        [Parameter(Position = 1, Mandatory = $false)] [switch]$DontHideBody,
        [Parameter(Position = 2, Mandatory = $false)] [Int]$ColorSwitchColumnNumber,
        [Parameter(Position = 3, Mandatory = $false)] [psObject]$ColorSwitchHashTable,
        [Parameter(Position = 4, Mandatory = $false)] [String]$ColorSwitchDefault,
        [Parameter(Position = 5, Mandatory = $false)] [String]$originator
        
    )  
    Process {
        $AdaptiveCard = @{}
        $AdaptiveCard.Add("version","1.0")
        $AdaptiveCard.Add("type","AdaptiveCard")
        if(!$DontHideBody.IsPresent){
            $AdaptiveCard.Add("hideOriginalBody","true")
        } 
        if([String]::IsNullOrEmpty($originator)){
            $AdaptiveCard.Add("originator",$originator)
        }       
        $Body = @()
        $BodyHash = @{}
        $BodyHash.Add("type","ColumnSet")
        $ColumnSet = @()
        if($Columns -ne $null){
            $RowNumber = 0    
            foreach($ColumnRow in $Columns){   
                $ColumnNumber = 0                      
                $ColumnRow.PSObject.Properties | ForEach-Object{                   
                    $CellNumber = 0
                    if($RowNumber -eq 0){
                        $ColumnItems = @()
                        $columnHeader = @{}
                        $columnHeader.Add("type", "TextBlock")
                        $columnHeader.Add("text", $_.Name)
                        $columnHeader.Add("size", "Large")
                        $columnHeader.Add("weight", "Bolder")
                        $columnHeader.Add("color", "Accent")
                        $ColumnItems += $columnHeader
                        $columnHash = @{}
                        $columnHash.Add("width","auto")            
                        $columnHash.Add("type","Column")
                        $columnHash.Add("items",$ColumnItems)    
                        $ColumnSet += $columnHash
                    }
                    $columnCell = @{}
                    $columnCell.Add("type", "TextBlock")
                    if($CellNumber -eq 0){
                        $columnCell.Add("weight", "Bolder")
                    }
                    if($ColumnNumber -eq $ColorSwitchColumnNumber){
                        if($ColorSwitchHashTable.Contains($_.Value)){
                            $columnCell.Add("color", $ColorSwitchHashTable[$_.Value])
                        }else{
                            $columnCell.Add("color", $ColorSwitchDefault)
                        }
                    }
                    $columnCell.Add("text", $_.Value)                    
                    $ColumnSet[$ColumnNumber]["items"] += $columnCell
                    $ColumnNumber++
                }
                $RowNumber++
            }
            
        }
        $BodyHash.Add("columns",$ColumnSet)
        $Body += $BodyHash
        $AdaptiveCard.Add("body",$Body)
        $convertedBody = ConvertTo-Json $AdaptiveCard -Depth 9
        return $convertedBody
    }
}