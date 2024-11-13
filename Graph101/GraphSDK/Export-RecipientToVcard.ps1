####################### 
<# 
.SYNOPSIS 
 This is an Extension cmldet Exo cmdlets that exports a Recpient from Get-ExoRecipient to a VCF file
 
.DESCRIPTION 
   Exports a Contact in an Exchange Online a VCF file 
  
  Requires the Exo Cmdlets
 
  Requires Powershell Graph SDK connected with Scope "ProfilePhoto.Read.All"
 eg connect-mggraph -Scopes "ProfilePhoto.Read.All
  
.EXAMPLE
 
	Example 1 Export a Recpient to a VCF file 
	Export-RecipientToVcard -FileName 'c:\temp\cc.vcf' -Identity user@contso.com  

        Example 2 Export a Recpient to a VCF file with the Contact Photo (if availble) 
	Export-RecipientToVcard -FileName 'c:\temp\cc.vcf' -Identity user@contso.com  -IncludePhoto

#> 
########################
function Export-RecipientToVcard {

	   [CmdletBinding()] 
    param( 
        [Parameter(Position = 1, Mandatory = $false)] [string]$Identity,
        [Parameter(Position = 2, Mandatory = $false)] [switch]$IncludePhoto,
        [Parameter(Position = 3, Mandatory = $true)]
        [string]
        $FileName
    )  
    Begin {
        $Recpient = Get-EXORecipient -Identity $Identity -Properties FirstName,LastName,DisplayName,Department,Company,Title,Office,emailAddresses,Phone,City,CountryOrRegion,PostalCode,PrimarySmtpAddress
        Write-Verbose $Recpient
        $directoryName = [System.IO.Path]::GetDirectoryName($FileName)
        $FileDisplayName = [System.IO.Path]::GetFileNameWithoutExtension($FileName);
        $FileExtension = [System.IO.Path]::GetExtension($FileName);
        if (Test-Path $FileName) {
            $fileUnqiue = $false
            $i = 0;
            do {	
                $i++
                if (-not (Test-Path $FileName)) {
                    $fileUnqiue = $true      
                }
                else {
                    $FileName = [System.IO.Path]::Combine($directoryName, $FileDisplayName + "(" + $i + ")" + $FileExtension);
                }			
            }while ($fileUnqiue -ne $true)   
        }     
        Set-content -path $FileName "BEGIN:VCARD"
        add-content -path $FileName "VERSION:2.1"
        $firstName = ""
        if ($Recpient.FirstName) {
            $firstName = $Recpient.FirstName
        }       
        $lastName = ""
        if ($Recpient.LastName) {
            $lastName = $Recpient.LastName
        }
        add-content -path $FileName ("N:" + $lastName + ";" + $firstName)
        add-content -path $FileName ("FN:" + $Recpient.DisplayName)
        $department = "";
        if ($Recpient.Department) {
            $department = $Recpient.Department
        }		
        $CompanyName = "";
        if ($Recpient.Company) {
            $CompanyName = $Recpient.Company
        }
        add-content -path $FileName ("ORG:" + $companyName + ";" + $department)
        if ($Recpient.Title) {
            add-content -path $FileName ("TITLE:" + $Recpient.Title)
        }
        if ($Recpient.Phone) {
            add-content -path $FileName ("TEL;WORK;VOICE:" + $Recpient.Phone)
        }
        if (![String]::IsNullOrEmpty($Recpient.City)) {
            if ($Recpient.CountryOrRegion) {
                $Country = $Recpient.CountryOrRegion.Replace("`n", "")
            }
            if ($Recpient.City) {
                $City = $Recpient.City.Replace("`n", "")
            }
            if ($Recpient.PostalCode) {
                $PCode = $Recpient.PostalCode.Replace("`n", "")
            }
            $addr = "ADR;WORK;PREF:;;;" + $City + ";;" + $PCode + ";" + $Country
            add-content -path $FileName $addr
        }
        $ImAddress = ($Recpient.EmailAddresses | where {$_.StartsWith("SIP:")})
        if($ImAddress){
            add-content -path $FileName $ImAddress.Replace("SIP:","")
        }       
        add-content -path $FileName ("EMAIL;PREF;INTERNET:" + $Recpient.PrimarySmtpAddress)		
        if ($IncludePhoto.IsPresent) {
            $EndPoint = "https://graph.microsoft.com/v1.0/Users"
            $RequestURL = $EndPoint + "('" + $Identity + "')/Photos/360x360/`$value" 
            if($Recpient.RecipientTypeDetails -eq "GroupMailbox"){
                $EndPoint = "https://graph.microsoft.com/v1.0/Groups"
                $RequestURL = $EndPoint + "('" + $Recpient.ExternalDirectoryObjectId + "')/Photos/360x360/`$value"
            }         
             
            $photoBytes = (Invoke-MgGraphRequest -Uri $RequestURL -ContentType "image/jpeg" -OutputType HttpResponseMessage).Content.ReadAsByteArrayAsync().Result 
            add-content -path $FileName "PHOTO;TYPE=JPEG;ENCODING=BASE64:"
            $ExportString = ""
            $ImageString = [System.Convert]::ToBase64String($photoBytes, [System.Base64FormattingOptions]::InsertLineBreaks)
            ForEach ($line in $($ImageString -split "`r`n"))
            {
               $ExportString += " " + $line + "`r`n"
            }      
            
            add-content -path $FileName $ExportString        
            add-content -path $FileName "`r`n"
        }
        add-content -path $FileName ("X-RecipientType:" + $Recpient.RecipientType)        
        add-content -path $FileName ("X-RecipientTypeDetails:" + $Recpient.RecipientTypeDetails)   
        add-content -path $FileName "END:VCARD"
        Write-Verbose ("Contact exported to $FileName")
        return $FileName
    } 
}




