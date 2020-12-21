<#
Example:
./CreateUsers -Path ou=Sales,ou=My Users,dc=wisesoft,dc=local -NumUsers 500

Requirements:
MaleFirstNames.txt
	- Text file containing a list of male first names
	http://www.census.gov/genealogy/names/dist.male.first
FemaleFirstNames.txt
	- Text file containing a list of female first names
	http://www.census.gov/genealogy/names/dist.female.first
LastNames.txt
	- Text file containing a list of surnames
	http://www.census.gov/genealogy/names/dist.all.last


#>
param(
	[Parameter(Mandatory=$true)]
	[int]$numusers,
	[Parameter(Mandatory=$true)]
	[string]$path,
	[string]$Department,
	[System.DateTime]$AccountExpirationDate,
	[string]$AccountPassword,
	[bool] $ChangePasswordAtLogon=$true,
	[string]$City,
	[string]$Company,
	[string]$Country,
	[string]$Description,
	[string]$Division,
	[bool]$Enabled=$true,
	[string]$Office,
	[string]$Organization
)
Import-Module ActiveDirectory
[Reflection.Assembly]::LoadWithPartialName("System.Web") 
Add-Type -AssemblyName Microsoft.VisualBasic

# Load a list of male names.  Remove all text after first space and convert to proper case
Write-Host "Loading Male first names...."
$MaleNames = Get-Content "MaleFirstNames.txt" | ForEach-Object{ 
	$name = $_
	If ($name.IndexOf(" ") -ge 0) {
		$name = $name.Substring(0,($name.IndexOf(" ")))
	}
	$name = [Microsoft.VisualBasic.Strings]::StrConv($name,'ProperCase')
	$name
}
# Load a list of female names.  Remove all text after first space and convert to proper 
Write-Host "Loading Female first names...."
$FemaleNames = Get-Content "FemaleFirstNames.txt" | ForEach-Object{ 
	$name = $_
	if ($name.IndexOf(" ") -ge 0) {
		$name = $name.Substring(0,($name.IndexOf(" ")))
	}
	$name = [Microsoft.VisualBasic.Strings]::StrConv($name,'ProperCase')
	$name
}
# Load a list of surnames.  Remove all text after first space and convert to proper 
Write-Host "Loading Surnames names...."
$Surnames = Get-Content "LastNames.txt" | ForEach-Object{ 
	$name = $_
	if ($name.IndexOf(" ") -ge 0) {
		$name = $_.Substring(0,($name.IndexOf(" ")))
	}
	$name = [Microsoft.VisualBasic.Strings]::StrConv($name,'ProperCase')
	$name
}
# Function to generate a random female name
Function GetRandomFemaleName(){
	$random = Get-Random -Minimum 0 -Maximum ($FemaleNames.Count)
	return $FemaleNames.GetValue($random)
}
# Function to generate a random male name
Function GetRandomMaleName(){
	$random = Get-Random -Minimum 0 -Maximum ($MaleNames.Count)
	return $MaleNames.GetValue($random) 
}
# Function to generate a random surname
Function GetRandomSurname() {
	$random = Get-Random -Minimum 0 -Maximum ($Surnames.Count)
	return $Surnames.GetValue($random)
}
# Password supplied via command line
if ($AccountPassword.Length -gt 0){
	$password=$AccountPassword
	#Password needs to be converted to a secure string type for New-ADUser
	$passwordss = ConvertTo-SecureString -AsPlainText $password -Force
}

for ($i = 0; $i -lt $numusers; $i++) {
	# Even distribution of male/female names are used
	if (($i % 2) -eq 0){
		$givenName = GetRandomMaleName
	}
	else {
		$givenName = GetRandomFemaleName
	}
	$sn = GetRandomSurname
	$displayName="$givenName $sn"
	$sAMAccountName="$givenname.$sn"
	$sAMAccountName = $sAMAccountName.Substring(0,[System.Math]::Min($sAMAccountName.Length,20))
	# Generate a random password if not supplied via the command line
	if ($AccountPassword.Length -eq 0)
	{
		$password = [System.Web.Security.Membership]::GeneratePassword(20,2)
		#Password needs to be converted to a secure string type for New-ADUser
		$passwordss = ConvertTo-SecureString -AsPlainText $password -Force
	}
	
	$retryCount =0
	do{
		# Loop will be exited unless an account with the same name already exists (in which it will retry with a uniquefier)
		$completed=$true
		$cn = $sAMAccountName
		$email = "$sAMAccountName@$env:USERDNSDOMAIN"
		Write-Host "Creating user"($i+1)"of $numusers.....$sAMAccountName | Pass