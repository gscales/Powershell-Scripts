function New-MessageCard { 
    param( 
        [Parameter(Position = 0, Mandatory = $false)] [psObject]$Facts,
        [Parameter(Position = 1, Mandatory = $false)] [String]$Title,
        [Parameter(Position = 2, Mandatory = $false)] [String]$Summary
    )  
    Process {
        $MessageCard = @{}
        $MessageCard.Add("@type", "MessageCard")
        $MessageCard.Add("@context", "http://schema.org/extensions")
        $MessageCard.Add("summary", $Summary)
        $MessageCard.Add("themeColor", "0078D7")
        $MessageCard.Add("title", $Title)
        $Sections = @()
        $SectionsHash = @{}
        if($Facts -ne $null){
            $factsCollection = @()
            foreach($fact in $Facts){
                $factEntry = @{}
                $val =0
                $fact.PSObject.Properties | ForEach-Object{
                    if($val -eq 0){
                        $factEntry.Add("name",$_.Value)
                        $val = 1
                    }else{
                        $factEntry.Add("value",$_.Value)
                        $val = 0
                    }
                }
                $factsCollection += $factEntry
                
            }   
            $SectionsHash.Add("facts",$factsCollection)
        }
        $Sections += $SectionsHash 
        $MessageCard.Add("sections", $Sections)
        $convertedBody = ConvertTo-Json $MessageCard -Depth 9
        return $convertedBody
    }
}