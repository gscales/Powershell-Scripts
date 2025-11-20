function Get-M365TenantAudienceMetadata {
<#
.SYNOPSIS
    Retrieves the 'allowedAudiences' metadata for a given Microsoft 365 tenant ID (GUID).

.DESCRIPTION
    This function takes a tenant ID (GUID) and queries the public
    accounts.accesscontrol.windows.net metadata endpoint using the GUID.
    It then parses the JSON response, extracts all domains from the 
    'allowedAudiences' array, consolidates subdomains down to their root
    domain (e.g., a.domain.com -> domain.com), and outputs the unique list
    of root domains.

.PARAMETER TenantId
    The GUID (Tenant ID) of the tenant to query. This is a mandatory parameter.

.EXAMPLE
    Get-M365TenantAudienceMetadata -TenantId "00000000-0000-0000-0000-000000000000"

    This will query the metadata for the specified Tenant ID and output the allowed audiences.

.NOTES
    This relies on a public, unauthenticated endpoint.
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $true,
               Position = 0,
               HelpMessage = "Enter the M365 tenant ID (GUID)")]
    [string]$TenantId
)

    Write-Verbose "Querying metadata for tenant ID: $TenantId"

    # Construct the metadata URI using the Tenant ID (GUID)
    $MetadataUri = "https://accounts.accesscontrol.windows.net/$TenantId/metadata/json/1"

    try {
        Write-Verbose "Sending web request to: $MetadataUri"
        
        # Send the web request and get the content
        # Use -ErrorAction Stop to force exceptions into the catch block for clean handling
        $response = Invoke-WebRequest -Uri $MetadataUri -UseBasicParsing -ErrorAction Stop
        
        # Parse the JSON content from the response
        $metadata = $response.Content | ConvertFrom-Json -ErrorAction Stop
        
        # Initialize a dictionary (hash table) to store unique root domains
        $RawDomainList = @{}
        
        # Check if 'allowedAudiences' property exists 
        if ($metadata.PSObject.Properties.Name -contains 'allowedAudiences') {
            foreach($entry in $metadata.allowedAudiences){
                # 1. Clean up the entry by removing the constant URI prefix
                $cleanedDomain = $entry.Replace("00000001-0000-0000-c000-000000000000/accounts.accesscontrol.windows.net@","")
                
                # 2. Get the root domain using the helper function
                $rootDomain = Get-RootDomain -Domain $cleanedDomain
                
                if(![String]::IsNullOrEmpty($rootDomain)){
                    # 3. Add the unique root domain to the list
                    if(!$RawDomainList.ContainsKey($rootDomain)){
                        $RawDomainList.Add($rootDomain, $true) 
                    } 
                } 
            }
        }
        else {
            Write-Warning "The 'allowedAudiences' property was not found in the response for Tenant ID $TenantId."
        }
    }
    catch [System.Net.WebException] {
        # Handle HTTP errors (e.g., 404 Not Found, 400 Bad Request)
        $statusCode = $_.Exception.Response.StatusCode
        Write-Error "Error retrieving metadata for Tenant ID $TenantId. Status code: $statusCode"
        Write-Error "Please ensure the Tenant ID '$TenantId' is correct and represents a valid tenant."
    }
    catch {
        # Catch any other unexpected errors
        Write-Error "An unexpected error occurred: $($_.Exception.Message)"
    }
    
    # Output the unique list of root domains
    Write-Output $RawDomainList.Keys
}

# --- 1. HELPER FUNCTION ---
# This function calculates the "Root Domain" (e.g., "a.domain.com" -> "domain.com") 
# while correctly handling multi-part TLDs (e.g., "sub.example.com.au" -> "example.com.au").
function Get-RootDomain {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Domain
    )

    # Filter out non-domain strings (like UUIDs) that lack a dot
    if ($Domain -notmatch '\.') {
        return $null
    }

    # Handle the special M365 default domain
    if($Domain -match "onmicrosoft.com"){
        return $Domain
    }

    # Split the domain by the dot and remove any empty parts
    $parts = $Domain.Split('.') | Where-Object { $_ }

    # Define common two-part public suffixes that act as the TLD itself
    # This list allows us to correctly identify the root domain in three-part cases.
    $publicSuffixes = @(
        "com.au", "net.au", "org.au", "gov.au",
        "co.uk", "org.uk",
        "co.jp", "ne.jp",
        "com.br", "com.tr",
        "co.in"
    )

    # Check if the domain ends with a recognized two-part public suffix (e.g., 'com.au')
    if ($parts.Count -ge 2) {
        $lastTwoParts = ($parts[-2], $parts[-1]) -join '.'
        
        if ($publicSuffixes -contains $lastTwoParts) {
            # This is a multi-part TLD (e.g., .com.au, .co.uk).
            if ($parts.Count -ge 3) {
                # The root domain is the SLD (parts[-3]) plus the two-part suffix.
                # Example: a.management.com.au -> management.com.au
                return ($parts[-3], $lastTwoParts) -join '.'
            }
        }
    }
    
    # Default behavior: Assume standard SLD.TLD (e.g., a.example.com -> example.com)
    if ($parts.Count -ge 2) {
        # Return the last two parts
        return ($parts[-2], $parts[-1]) -join '.'
    }

    # Fallback for single-part names or unexpected formats
    return $Domain
}

