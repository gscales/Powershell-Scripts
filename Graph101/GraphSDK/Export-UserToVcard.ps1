####################### 
<# 
.SYNOPSIS 
 This is an Extension cmldet Exo cmdlets that exports a Recpient from Get-ExoRecipient to a VCF file
 
.DESCRIPTION 
   Exports a Contact in an Exchange Online a VCF file 
  
  Requires the Exo Cmdlets
  
.EXAMPLE
 
	Example 1 Export a Contact to a VCF file 
	Export-MGPContactToVcard -FileName 'c:\temp\cc.vcf' -UserId user@contso.com  -ContactId $ContactIdVar

    Example 2 Export a Contact to a VCF file with the Contact Photo (if availble) 
	Export-MGPContactToVcard -FileName 'c:\temp\cc.vcf' -UserId user@contso.com -ContactId $ContactIdVar -IncludePhoto

#> 
########################
function Export-UserToVcard {

	   [CmdletBinding()] 
    param( 
        [Parameter(Position = 1, Mandatory = $false)] [string]$Identity,
        [Parameter(Position = 2, Mandatory = $false)] [switch]$IncludePhoto,
        [Parameter(Position = 3, Mandatory = $true)]
        [string]
        $FileName
    )  
    Begin {
        $User = Get-User -Identity $Identity 
        Write-Verbose $User
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
        if ($User.FirstName) {
            $firstName = $User.FirstName
        }       
        $lastName = ""
        if ($User.LastName) {
            $lastName = $User.LastName
        }
        add-content -path $FileName ("N:" + $lastName + ";" + $firstName)
        add-content -path $FileName ("FN:" + $User.DisplayName)
        $department = "";
        if ($User.Department) {
            $department = $User.Department
        }		
        $CompanyName = "";
        if ($User.Company) {
            $CompanyName = $User.Company
        }
        add-content -path $FileName ("ORG:" + $companyName + ";" + $department)
        if (![String]::IsNullOrEmpty($User.Title)) {
            add-content -path $FileName ("TITLE:" + $User.Title)
        }
        if (![String]::IsNullOrEmpty($User.Phone)) {
            add-content -path $FileName ("TEL;WORK;VOICE:" + $User.Phone)
        }
        if (![String]::IsNullOrEmpty($User.mobilePhone)) {
            add-content -path $FileName ("TEL;CELL;VOICE:" + $User.mobilePhone)
        }
        if (![String]::IsNullOrEmpty($User.homePhone)) {
            add-content -path $FileName ("TEL;HOME;VOICE:" + $User.homePhone)
        }
        if (![String]::IsNullOrEmpty($User.Fax)) {
            add-content -path $FileName ("TEL;WORK;FAX:" + $User.Fax)
        }
        if (![String]::IsNullOrEmpty($User.City)) {
            if (![String]::IsNullOrEmpty($User.CountryOrRegion)){
                $Country = $User.CountryOrRegion.Replace("`n", "")
            }
            if (![String]::IsNullOrEmpty($User.City)) {
                $City = $User.City.Replace("`n", "")
            }
            if (![String]::IsNullOrEmpty($User.PostalCode)) {
                $PCode = $User.PostalCode.Replace("`n", "")
            }
            if (![String]::IsNullOrEmpty($User.StreetAddress)) {
                $StreetAddress = $User.StreetAddress.Replace("`n", "")
            }
            $addr = "ADR;WORK;PREF:;;" + $StreetAddress + ";" + $City + ";" + $StateOrProvince + ";" + $PCode + ";" + $Country
            add-content -path $FileName $addr
        }
        if (![String]::IsNullOrEmpty($User.WebPage)) {
            add-content -path $FileName ("URL;WORK:" + $User.WebPage)
        }        
        add-content -path $FileName ("EMAIL;PREF;INTERNET:" + $User.WindowsEmailAddress)		
        if ($IncludePhoto.IsPresent) {
            $EndPoint = "https://graph.microsoft.com/v1.0/Users"
            $RequestURL = $EndPoint + "('" + $Identity + "')/Photos/360x360/`$value"  
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
        add-content -path $FileName ("X-RecipientType:" + $User.RecipientType)        
        add-content -path $FileName ("X-RecipientTypeDetails:" + $User.RecipientTypeDetails)   
        add-content -path $FileName "END:VCARD"
        Write-Verbose ("Contact exported to $FileName")
        return $FileName
    } 
}




