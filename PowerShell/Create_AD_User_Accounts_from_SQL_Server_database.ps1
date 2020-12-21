Import-Module ActiveDirectory

######## Edit Here ########

# Path to create users in
$path="OU=AdventureWorks,DC=WiseSoft,DC=local"

# Connection string to SQL Server database
$connectionString = "Server=WISESQL01;Initial Catalog=AdventureWorks2012;Integrated Security=SSPI;"

# Select statement to return new user accounts
# Needs to return "sAMAccountName" & "Password" columns
# Note: Other columns names should match AD attribute name
$sql="SELECT LEFT(REPLACE(FirstName + '.' + LastName,' ',''),20)  as sAMAccountName,
		FirstName + '.' + LastName + '@wisesoft.co.uk' as userPrincipalName,
		FirstName as givenName,
		LastName as sn,
		JobTitle as title,
		phoneNumber as telePhoneNumber,
		AddressLine1 + ISNULL(CHAR(13) + CHAR(10) + AddressLine2,'') as streetAddress,
		City as l,
		StateProvinceName as st,
		PostalCode as postalCode,
		EmailAddress as mail,
		'MyP@$$w0rdG3n3r@t3dFr0mD@t@b@s3' as Password
FROM [HumanResources].[vEmployee]"

###########################

$cn = new-object system.data.sqlclient.sqlconnection
$cn.ConnectionString = $connectionString
$cn.Open()
$cmd = New-Object System.Data.SqlClient.SqlCommand
$cmd.CommandText = $sql
$cmd.connection = $cn
$dr = $cmd.ExecuteReader()

$colCount = $dr.FieldCount
$sAMAccountNameOrdinal = $dr.GetOrdinal("sAMAccountName")
$PasswordOrdinal = $dr.GetOrdinal("Password")

while ($dr.Read()) {
	# Get value of sAMAccountName column
	$sAMAccountName = $dr.GetValue($sAMAccountNameOrdinal)
	# Get value password column (converted to secure string for New-ADUser Cmdlet)
	$password = ConvertTo-SecureString -AsPlainText $dr.GetValue($PasswordOrdinal) -Force
		
	write-host "Creating user account..." $sAMAccountName

	$otherAttributes = New-Object System.Collections.HashTable

	# Create a hash table of attribute names and attribute values
	# Used to populate other attributes. 
	for($i=0;$i -le $colCount-1;$i++)
	{
		$attribute = $dr.GetName($i)

		switch ($attribute)
		{
			"Password"{} 		#Ignore
			"SAMAccountName" {} 	#Ignore
			default{
				$otherAttributes.Add($attribute,$dr.GetValue($i))
			}
		}
	}
	# Create Active Directory User Account
	New-ADUser -sAMAccountName $sAMAccountName -Name $sAMAccountName -Path $path -otherAttributes $otherAttributes -Enable $true -AccountPassword $password 

}

$dr.Close()
$cn.Close()