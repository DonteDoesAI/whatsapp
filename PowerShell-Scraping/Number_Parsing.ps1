#Requires -Version 7.0

Import-Module ImportExcel

# Installing PhoneNumberParser
# Register-PackageSource `
#     -Name NuGet `
#     -ProviderName NuGet `
#     -Location "https://www.nuget.org/api/v2" `
#     -ErrorAction SilentlyContinue 

# Install-Package `
#     PhoneNumberParser `
#     -Scope CurrentUser `
#     -Force `
#     -Verbose `
#     -SkipDependencies

$package = [System.IO.Path]::combine(
    $env:LOCALAPPDATA,
    "PackageManagement",
    "Nuget",
    "Packages",
    "PhoneNumberParser.3.1.0",
    "Lib",
    "net7.0",
    "PhoneNumbers.dll"
)

Add-Type -Path $package
Add-Type -AssemblyName "System.Runtime"

# [PhoneNumbers.PhoneNumber]("+15712150398")

$phone_numbers = $null
# | foreach {
#     try {
#         [PhoneNumbers.PhoneNumber]($_.'phone_number') 
#     } catch {
#         $_.'phone_number'
#     }
}

# foreach ($number in $phone_numbers) {
    
# }

Set-WhatsAppMessage `
    -Message $phone_numbers.phone_number `
    -Web_Driver $driver

