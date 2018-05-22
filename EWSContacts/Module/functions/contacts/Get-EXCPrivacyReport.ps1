function Get-EXCPrivacyReport {

    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true,  ValueFromPipeline=$true) ]
        [Microsoft.Exchange.WebServices.Data.Contact]$Contact	)
    Begin {
        $rptObject = "" | Select EmailAddresses, PhoneNumbers, PhysicalAddresses, BirthDay, JobDetails, HomeAddress, HomePhone, MobilePhone,HasNotes, PrivacyPoints
        #Set Defaults
        $rptObject.EmailAddresses = 0
        $rptObject.PhoneNumbers = 0
        $rptObject.PhysicalAddresses = 0
        $rptObject.BirthDay = $false
        $rptObject.JobDetails = $false
        $rptObject.HomeAddress = $false
        $rptObject.HomePhone = $false
        $rptObject.MobilePhone = $false
        $rptObject.HasNotes = $false
        $rptObject.PrivacyPoints = 0
        $BusinessPhone = $null
        $MobilePhone = $null
        $HomePhone = $null
        $HomePhone2 = $null
        $OtherTelephone = $null
        if ($Contact.PhoneNumbers -ne $null) {
            if ($Contact.PhoneNumbers.TryGetValue([Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone, [ref]$BusinessPhone)) {
                $rptObject.PhoneNumbers++
                $rptObject.PrivacyPoints++
            }
            if ($Contact.PhoneNumbers.TryGetValue([Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone, [ref]$MobilePhone)) {
                $rptObject.PhoneNumbers++
                $rptObject.PrivacyPoints++
                $rptObject.MobilePhone = $true
            }
            if ($Contact.PhoneNumbers.TryGetValue([Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::HomePhone, [ref]$HomePhone)) {
                $rptObject.PhoneNumbers++
                $rptObject.PrivacyPoints++
                $rptObject.HomePhone = $true
            }
            if ($Contact.PhoneNumbers.TryGetValue([Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::HomePhone2, [ref]$HomePhone2)) {
                $rptObject.PhoneNumbers++
                $rptObject.PrivacyPoints++
                $rptObject.HomePhone = $true
            }
            if ($Contact.PhoneNumbers.TryGetValue([Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::OtherTelephone, [ref]$OtherTelephone)) {
                $rptObject.PhoneNumbers++
                $rptObject.PrivacyPoints++
                $rptObject.HomePhone = $true
            }
            
        }
        $EmailAddress1 = $null
        $EmailAddress2 = $null
        $EmailAddress3 = $null
        if ($Contact.EmailAddresses -ne $null){
            if ($Contact.EmailAddresses.TryGetValue([Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1,[ref]$EmailAddress1)) {
                $rptObject.EmailAddresses++
                $rptObject.PrivacyPoints++
            }
            if ($Contact.EmailAddresses.TryGetValue([Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress2,[ref]$EmailAddress2)) {
                $rptObject.EmailAddresses++
                $rptObject.PrivacyPoints++
            }
            if ($Contact.EmailAddresses.TryGetValue([Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress3,[ref]$EmailAddress3)) {
                $rptObject.EmailAddresses++
                $rptObject.PrivacyPoints++
            }
        }

        $HomeAddress = $null
        $BusinessAddress = $null
        $OtherAddress = $null
        if ($Contact.PhysicalAddresses -ne $null) {
            if ($Contact.PhysicalAddresses.TryGetValue([Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Home, [ref]$HomeAddress)) {
                $rptObject.HomeAddress = $true
                $rptObject.PhysicalAddresses++
                $rptObject.PrivacyPoints++
            }
            if ($Contact.PhysicalAddresses.TryGetValue([Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business, [ref]$BusinessAddress)) {
                $rptObject.PhysicalAddresses++
                $rptObject.PrivacyPoints++
            }
            if ($Contact.PhysicalAddresses.TryGetValue([Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Other, [ref]$OtherAddress)) {
                $rptObject.PhysicalAddresses++
                $rptObject.PrivacyPoints++
            }
            
        }
        if (![String]::IsNullOrEmpty($Contact.Birthday)) {
            $rptObject.BirthDay = $true
            $rptObject.PrivacyPoints++
        }
        if(![String]::IsNullOrEmpty($Contact.Department)){
            $rptObject.JobDetails = $true
            $rptObject.PrivacyPoints++
        }
        if(![String]::IsNullOrEmpty($Contact.JobTitle)){
            $rptObject.JobDetails = $true
            $rptObject.PrivacyPoints++
        }
        if(![String]::IsNullOrEmpty($Contact.Profession)){
            $rptObject.JobDetails = $true
            $rptObject.PrivacyPoints++
        }
        if(![String]::IsNullOrEmpty($Contact.Body.Text)){
            $rptObject.HasNotes = $true
            $rptObject.PrivacyPoints++
        }
        return $rptObject

    }
}
