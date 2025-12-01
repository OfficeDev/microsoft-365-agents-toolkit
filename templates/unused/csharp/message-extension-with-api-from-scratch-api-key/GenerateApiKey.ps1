# Define the length of the random string
$length = 12

# Create a new byte array of the specified length
$randomBytes = New-Object Byte[] $length

# Fill the byte array with random numbers
$randomNumberGenerator = [System.Security.Cryptography.RandomNumberGenerator]::Create()
$randomNumberGenerator.GetBytes($randomBytes)

# Convert the byte array to a Base64 string
$randomString = [Convert]::ToBase64String($randomBytes)

# Get a substring of the Base64 string with the specified length
$key = $randomString.Substring(0, [Math]::Min($length, $randomString.Length))

# Output the generated key
Write-Host "Generated a new API Key: $key"