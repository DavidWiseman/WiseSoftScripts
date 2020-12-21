# Length of the password to be generated
$PasswordLength = 20

# VB.Net source code for SecureRandom class.  Used to generate random numbers using the RNGCryptoServiceProvider
# which is more secure than using System.Random. Class is based on code provided by Jim Mischel
# http://www.informit.com/guides/content.aspx?g=dotnet&seqNum=775
$source="
Imports System.Security.Cryptography
Public Class SecureRandom
	Private Const BufferSize As Integer = 1024	' must be a multiple of 4
	Private RandomBuffer As Byte()
	Private BufferOffset As Integer
	Private rng As RNGCryptoServiceProvider
	Public Sub New()
		RandomBuffer = New Byte(BufferSize - 1) {}
		rng = New RNGCryptoServiceProvider()
		BufferOffset = RandomBuffer.Length
	End Sub
	Private Sub FillBuffer()
		rng.GetBytes(RandomBuffer)
		BufferOffset = 0
	End Sub
	Public Function [Next]() As Integer
		If BufferOffset >= RandomBuffer.Length Then
			FillBuffer()
		End If
		Dim val As Integer = System.BitConverter.ToInt32(RandomBuffer, BufferOffset) And &H7fffffff
		BufferOffset += 4
		Return val
	End Function
	Public Function [Next](maxValue As Integer) As Integer
		Return [Next]() Mod maxValue
	End Function
	Public Function [Next](minValue As Integer, maxValue As Integer) As Integer
		If maxValue < minValue Then
			Throw New System.ArgumentOutOfRangeException(""maxValue must be greater than or equal to minValue"")
		End If
		Dim range As Integer = maxValue - minValue
		Return minValue + [Next](range)
	End Function
	Public Function NextDouble() As Double
		Dim val As Integer = [Next]()
		Return CDbl(val) / Integer.MaxValue
	End Function
	Public Sub GetBytes(buff As Byte())
		rng.GetBytes(buff)
	End Sub
End Class"

# Add the SecureRandom class so we can use it in PowerShell
Add-Type -TypeDefinition $source -Language VisualBasic

# Create a new instance of SecureRandom, used to generate random numbers.
$rngRand = New-Object SecureRandom

# Used to store an array of characters that can be used for the password
$CharPool = New-Object System.Collections.ArrayList
# Add characters a-z to the arraylist
for ($index = 97; $index -le 122; $index++) { [Void]$CharPool.Add([char]$index) }
# Add characters A-Z to the arraylist
for ($index = 65; $index -le 90; $index++) { [Void]$CharPool.Add([Char]$index) }
# Add digits 0-9 to the arraylist
$CharPool.AddRange(@("0","1","2","3","4","5","6","7","8","9"))
# Add a range of special characters to the arraylist
$CharPool.AddRange(@("!","""","#","$","%","&","'","(",")","*","+","-",".","/",":",";","<","=",">","?","@","[","\","]","^","_","{","|","}","~","!"))

$password=""

# Generate password by appending a random value from the array list until desired length of password is reached
1..$PasswordLength | foreach { $password = $password + $CharPool[$rngRand.Next(0,$CharPool.Count)] }	
# Print password
$password