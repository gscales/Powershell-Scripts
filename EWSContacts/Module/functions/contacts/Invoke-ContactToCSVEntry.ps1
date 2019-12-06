function Invoke-ContactToCSVEntry
{
<#
    Converts a Contact to CSV Entry
#>
	
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[Microsoft.Exchange.WebServices.Data.Contact]
		$Contact
	)
	Begin
	{
        $csvObj = "" | select DisplayName,GivenName,Surname,Department,CompanyName,Gender,Email1DisplayName,Email1Type,Email1EmailAddress,Email2DisplayName,Email2Type,Email2EmailAddress,Email3DisplayName,Email3Type,Email3EmailAddress,ImAddress,BusinessPhone,MobilePhone,BusinessFax,HomePhone,BusinessHomePage,BusinessStreet,BusinessCity,BusinessState,BusinessCountryOrRegion,BusinessPostalCode,HomeStreet,HomeCity,HomeState,HomeCountryOrRegion,HomePostalCode,JobTitle

        if ($Contact.GivenName -ne $null)
        {
            $csvObj.GivenName = $Contact.GivenName
        }
        $surname = ""
        if ($Contact.Surname -ne $null)
        {
            $csvObj.Surname = $Contact.Surname
        }

        if ($Contact.Department -ne $null)
        {
            $csvObj.Department = $Contact.Department
        }   

        if ($Contact.CompanyName -ne $null)
        {
            $csvObj.CompanyName = $Contact.CompanyName
        }

        if ($Contact.JobTitle -ne $null)
        {
            $csvObj.JobTitle = $Contact.JobTitle
        }
        if ($Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone] -ne $null)
        {
            $csvObj.MobilePhone = $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone]
        }
        if ($Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::HomePhone] -ne $null)
        {
            $csvObj.HomePhone = $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::HomePhone]
        }
        if ($Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone] -ne $null)
        {
            $csvObj.BusinessPhone = $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone]
        }
        if ($Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessFax] -ne $null)
        {
            $csvObj.BusinessFax = $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessFax]
        }
        if ($Contact.BusinessHomePage -ne $null)
        {
            $csvObj.BusinessHomePage =  $Contact.BusinessHomePage
        }
        if ($Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business] -ne $null)
        {

            if ($Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].CountryOrRegion -ne $null)
            {
                $csvObj.BusinessCountryOrRegion = $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].CountryOrRegion.Replace("`n", "")
            }
            if ($Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].City -ne $null)
            {
                $csvObj.BusinessCity = $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].City.Replace("`n", "")
            }
            if ($Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].Street -ne $null)
            {
                $csvObj.BusinessStreet = $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].Street.Replace("`n", "")
            }
            if ($Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].State -ne $null)
            {
                $csvObj.BusinessState = $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].State.Replace("`n", "")
            }
            if ($Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].PostalCode -ne $null)
            {
                $csvObj.BusinessPostalCode = $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].PostalCode.Replace("`n", "")
            }
        }
        if ($Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Home] -ne $null)
        {

            if ($Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Home].CountryOrRegion -ne $null)
            {
                $csvObj.HomeCountryOrRegion = $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Home].CountryOrRegion.Replace("`n", "")
            }
            if ($Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Home].City -ne $null)
            {
                $csvObj.HomeCity = $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Home].City.Replace("`n", "")
            }
            if ($Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Home].Street -ne $null)
            {
                $csvObj.HomeStreet = $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Home].Street.Replace("`n", "")
            }
            if ($Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Home].State -ne $null)
            {
                $csvObj.HomeState = $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Home].State.Replace("`n", "")
            }
            if ($Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Home].PostalCode -ne $null)
            {
                $csvObj.HomePostalCode = $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Home].PostalCode.Replace("`n", "")
            }
        }
        if ($Contact.ImAddresses[[Microsoft.Exchange.WebServices.Data.ImAddressKey]::ImAddress1] -ne $null)
        {
            $csvObj.ImAddress =  $Contact.ImAddresses[[Microsoft.Exchange.WebServices.Data.ImAddressKey]::ImAddress1]
        }
        if($Contact.EmailAddresses.Contains([Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1)){                  
            $csvObj.Email1DisplayName = $Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Name  
            $csvObj.Email1Type = $Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].RoutingType  
            $csvObj.Email1EmailAddress = $Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address  
        }  
        if($Contact.EmailAddresses.Contains([Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress2)){                  
            $csvObj.Email2DisplayName = $Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress2].Name  
            $csvObj.Email2Type = $Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress2].RoutingType  
            $csvObj.Email2EmailAddress = $Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress2].Address  
        }  
        if($Contact.EmailAddresses.Contains([Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress3)){                  
            $csvObj.Email3DisplayName = $Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress3].Name  
            $csvObj.Email3Type = $Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress3].RoutingType  
            $csvObj.Email3EmailAddress = $Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress3].Address  
        }
        return  $csvObj  
        				

	}
}
