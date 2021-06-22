<#
  .SYNOPSIS
  Get Exchange Server schema and version related information for an exisiting Exchange Organization
   
  Thomas Stensitzki
	
  THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
  RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
  Version 1.0, 2018-05-22

  .DESCRIPTION

  This script gets the Exchange schema version from the Active Directory schema partition

  The Exchange organization name is fetched from Active Directory automatically
  
  The script fetches at forest level:
  - objectVersion of MESO Container
  - rangeUpper of ms-Exch-Schema-Version-Pt 
  - msExchProductId of Exchange Organization container
  - objectVersion of Exchange Organization container

  The script fetches at forest level:
  - objectVersion of MESO Container

  .LINK 
  http://scripts.granikos.eu

  .NOTES 
  Requirements 
  - Windows Server 2012 R2, Windows Server 2016

  Revision History 
  -------------------------------------------------------------------------------- 
  1.0     Initial release

  .EXAMPLE
  Fetch all version information in the Active Directory forest
  .\Get-ExchangeServerVersionInfo.ps1
#>
[CmdletBinding()]
param(
)

Import-Module -Name ActiveDirectory

#region Functions 

function Get-ExchangeOrganizationName {
  <#
    .SYNOPSIS
    This function fetches Exchange organization name from Active Directory configuration partition

    .DESCRIPTION
    The function determines the forest root domain and queries the Microsoft Exchange container
    in the Active Directory configuration partition to get name of the msExchOrganizationContainer.
  #>

  # Get Active Directory Forest Distinguihsed Name
  $ForestNameDN = Get-ADDomain -Identity (Get-ADDomain).Forest | Select-Object -ExpandProperty DistinguishedName
  
  # Fetch Exchange Services hive from Active Directory Configuration Partition
  $Configuration = [ADSI]('LDAP://CN=Microsoft Exchange,CN=Services,CN=Configuration,{0}' -f $ForestNameDN) 

  # Get Exchange Organization Name from Exchange Services hive
  $OrganizationName = ($Configuration.psbase.children | Where-Object {$_.objectClass -eq 'msExchOrganizationContainer'}).Name

  return $OrganizationName

}

function Get-ExchangeSchemaVersion {
  <#
    .SYNOPSIS
    This function fetches the rangeUpper value from Exchange schema object ms-Exch-Schema-Version-Pt,CN=Schema

    .DESCRIPTION
    The function determines the forest root domain and connects to the schema partition to read the
    rangeUpper value of the ms-Exch-Schema-Version-Pt object.
  #>
  
  # Get Active Directory Forest Distinguished Name
  $ForestNameDN = Get-ADDomain -Identity (Get-ADDomain).Forest | Select-Object -ExpandProperty DistinguishedName

  # Get rangeUpper attribute
  $RangeUpper =([ADSI]('LDAP://CN=ms-Exch-Schema-Version-Pt,CN=Schema,CN=Configuration,{0}' -f $ForestNameDN)).rangeUpper

  return $RangeUpper

}

function Get-ExchangeDomainInformation {
  <#
    .SYNOPSIS
    Fetches the objectVersion attribute value of the MESO container object. 
    Fetches the msExchProductId and objectVersion attributes of Exchange Organization object, if the domain is the forest root domain.

    .DESCRIPTION
    The script determines whether to domain is the forest root domain. If the domain is the forest root the script fetches the following attributes:
    - MESO container objectVersion
    - Exchange organization msExchProductId
    - Exchange organization objectVersion

    If the domain is not the forest root domain, the script fetches the following attribute:
    - MESO container objectVersion

    .PARAMETER DomainName
    The Active Directory domain name of the domain to query

    .PARAMETER ExchangeOrganizationName
    The Exchange organization name

    .EXAMPLE
    Get-ExchangeDomainInformation -DomainName varunagroup.de -ExchangeOrganizationName Varuna-Group
    Get the Exchange related domain information for Active Directory domain varunagroup.de and Exchange organization named Varuna-Group
  #>

  param (
    [Parameter(Mandatory,HelpMessage='Active Directory Domain Name')][string]$DomainName,
    [Parameter(Mandatory,HelpMessage='Provide the Exchange Organization Name')][string]$ExchangeOrganizationName
  )

  $DomainDN = Get-ADDomain -Identity $DomainName | Select-Object -ExpandProperty DistinguishedName

  # Get Active Directory Forest Distinguihsed Name
  $ForestNameDN = Get-ADDomain -Identity (Get-ADDomain).Forest | Select-Object -ExpandProperty DistinguishedName

  # Get MESO Container object Version  
  $MESOObjectVersion =  ([ADSI]('LDAP://CN=Microsoft Exchange System Objects,{0}' -f $DomainDN)).objectVersion
  Write-Host ('MESO Container objectVersion           : {0}' -f $($MESOObjectVersion))

  if($DomainDN -eq $ForestNameDN) { 

    # Get Exchange ProductId (Version) of Exchange Organisation 
    $ConfigurationProductId =  ([ADSI]('LDAP://CN={0},CN=Microsoft Exchange,CN=Services,CN=Configuration,{1}' -f $ExchangeOrganizationName, $DomainDN)).msExchProductId
    Write-Host ('Exchange Configuration msExchProductId : {0}' -f $($ConfigurationProductId))

    # Get Exchange ObjectVersion of Exchange Organisation 
    $ConfigurationObjectVersion =  ([ADSI]('LDAP://CN={0},CN=Microsoft Exchange,CN=Services,CN=Configuration,{1}' -f $ExchangeOrganizationName, $DomainDN)).objectVersion
    Write-Host ('Exchange Configuration objectVersion   : {0}' -f $($ConfigurationObjectVersion))

  }
}

#endregion

## MAIN ##########################################

# Write forest root domain
Write-Host
Write-Host "Exchange Server Schema and Object Information for forest [$((Get-ADForest).Name.ToUpper())]" -ForegroundColor Gray 

# Fetch Exchange Organization name
$ExchangeOrgName = Get-ExchangeOrganizationName
Write-Host ('Exchange Organization Name        : {0}' -f $ExchangeOrgName)

# Write Exchange schema version
Write-Host ('Active Directory Schema rangeUpper: {0}' -f (Get-ExchangeSchemaVersion))

# Fetch all domains in the forest
$ForestDomains = (Get-ADForest).Domains

foreach($Domain in $ForestDomains) {

  Write-Host
  Write-Host ('Working on {0}' -f ($Domain.ToUpper())) 

  # Get domain related information  
  Get-ExchangeDomainInformation -DomainName $Domain -ExchangeOrganizationName $ExchangeOrgName

}