####################### 
<# 
.SYNOPSIS 
 This is an Extension cmldet from the Microsoft Graph Powershell SDK that Exports a Contact in an Exchange Online Mailbox to a VCF file
 
.DESCRIPTION 
   Exports a Contact in an Exchange Online Mailbox to a VCF file using the Microsoft Graph API
  
  Requires the Microsoft Graph Powershell SDK
  
.EXAMPLE
    This cmldet is an extension for the Microsoft Graph Powershell SDK and requires you have a connection context with at least  Contacts.Read eg

    Connect-MgGraph -Scopes "Contacts.Read" 

    You then need to find the Id of the contact you wish to export eg to the get the last 10 contacts in a Mailbox's contacts folder

    $Contacts = Get-MgUserContact -UserId user@contso.com

    to export the first contact asign that id to a varible (or use it directly)eg

    $ContactIdVar = $Contacts[0].id
 
	Example 1 Export a Contact to a VCF file 
	Export-MGPContactToVcard -FileName 'c:\temp\cc.vcf' -UserId user@contso.com  -ContactId $ContactIdVar

    Example 2 Export a Contact to a VCF file with the Contact Photo (if availble) 
	Export-MGPContactToVcard -FileName 'c:\temp\cc.vcf' -UserId user@contso.com -ContactId $ContactIdVar -IncludePhoto

#> 
########################
function Export-MGPContactToVcard {

	   [CmdletBinding()] 
    param( 
        [Parameter(Position = 1, Mandatory = $false)] [string]$UserId,
        [Parameter(Position = 2, Mandatory = $true)] [string]$ContactId,
        [Parameter(Position = 3, Mandatory = $false)] [switch]$IncludePhoto,
        [Parameter(Position = 4, Mandatory = $true)]
        [string]
        $FileName
    )  
    Begin {
        $EndPoint = "https://graph.microsoft.com/v1.0/users"
        $RequestURL = $EndPoint + "('$UserId')/contacts/" + $ContactId
        $RequestURL += "?`$expand=SingleValueExtendedProperties(`$filter=(Id eq 'Boolean {00062004-0000-0000-C000-000000000046} Id 0x8015'))"
        Write-Verbose $RequestURL
        $ClientResult = Invoke-MgGraphRequest -Uri $RequestURL
        $Contact = $ClientResult       
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
        $givenName = ""
        if ($Contact.GivenName) {
            $givenName = $Contact.givenName
        }       
        $surname = ""
        if ($Contact.Surname) {
            $surname = $Contact.surname
        }
        add-content -path $FileName ("N:" + $surname + ";" + $givenName)
        add-content -path $FileName ("FN:" + $Contact.displayName)
        $Department = "";
        if ($Contact.department) {
            $Department = $Contact.department
        }		
        $CompanyName = "";
        if ($Contact.companyName) {
            $CompanyName = $Contact.companyName
        }
        add-content -path $FileName ("ORG:" + $CompanyName + ";" + $Department)
        if ($Contact.jobTitle) {
            add-content -path $FileName ("TITLE:" + $Contact.jobTitle)
        }
        if ($Contact.mobilePhone) {
            add-content -path $FileName ("TEL;CELL;VOICE:" + $Contact.mobilePhone)
        }
        if ($Contact.homePhones) {
            add-content -path $FileName ("TEL;HOME;VOICE:" + $Contact.homePhones)
        }
        if ($Contact.businessPhones) {
            add-content -path $FileName ("TEL;WORK;VOICE:" + $Contact.businessPhones)
        }
        if ($Contact.businessPhones) {
            add-content -path $FileName ("TEL;WORK;FAX:" + $Contact.businessFaxs)
        }
        if ($Contact.businessHomePage) {
            add-content -path $FileName ("URL;WORK:" + $Contact.businessHomePage)
        }
        if ($Contact.businessAddress) {
            if ($Contact.businessAddress.countryOrRegion) {
                $Country = $Contact.businessAddress.countryOrRegion.Replace("`n", "")
            }
            if ($Contact.businessAddress.city) {
                $City = $Contact.businessAddress.city.Replace("`n", "")
            }
            if ($Contact.businessAddress.street) {
                $Street = $Contact.businessAddress.street.Replace("`n", "")
            }
            if ($Contact.businessAddress.state) {
                $State = $Contact.businessAddress.state.Replace("`n", "")
            }
            if ($Contact.businessAddress.postalCode) {
                $PCode = $Contact.businessAddress.postalCode.Replace("`n", "")
            }
            $addr = "ADR;WORK;PREF:;" + $Country + ";" + $Street + ";" + $City + ";" + $State + ";" + $PCode + ";" + $Country
            add-content -path $FileName $addr
        }
        if ($Contact.imAddresses) {
            add-content -path $FileName ("X-MS-IMADDRESS:" + $Contact.imAddresses)
        }
        $emCnt = 1;
        foreach ($emailAddress in $Contact.emailAddresses) {
            if ($emCnt -eq 1) {
                add-content -path $FileName ("EMAIL;PREF;INTERNET:" + $emailAddress.address)
            }
            else {
                add-content -path $FileName ("EMAIL;" + $emCnt + ";INTERNET:" + $emailAddress.address)
            }
            $emCnt++		
        }     
		
        if ($IncludePhoto.IsPresent -band (Invoke-ExpandContactExtendedProperties -item $Contact)) {
            $RequestURL = $EndPoint + "('" + $UserId + "')/Contacts('" + $ContactId + "')/Photo/`$value"  
            $photoBytes = (Invoke-MgGraphRequest -Uri $RequestURL -OutputType HttpResponseMessage).Content.ReadAsByteArrayAsync().Result 
            add-content -path $FileName "PHOTO;ENCODING=BASE64;TYPE=JPEG:"
            $ImageString = [System.Convert]::ToBase64String($photoBytes, [System.Base64FormattingOptions]::InsertLineBreaks)
            add-content -path $FileName $ImageString
            add-content -path $FileName "`r`n"
        }
        add-content -path $FileName "END:VCARD"
        Write-Verbose ("Contact exported to $FileName")
        return $FileName
    } 
}



function Invoke-ExpandContactExtendedProperties {
    [CmdletBinding()] 
    param (
        [Parameter(Position = 1, Mandatory = $false)]
        [psobject]
        $Item
    )
 
    process {
        if ($Item.singleValueExtendedProperties -ne $null) {
            foreach ($Prop in $Item.singleValueExtendedProperties) {
                Switch ($Prop.Id) {
                    "Boolean {00062004-0000-0000-C000-000000000046} Id 0x8015" {
                        return [bool]::Parse($Prop.Value) 
                    }                
                }
            }
        }
        return $false
    }
}
