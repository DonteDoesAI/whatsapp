# Required Modules
# https://chromedriver.chromium.org/downloads - need a chrome driver that matches chrome version
# https://www.selenium.dev/selenium/docs/api/dotnet/html/T_OpenQA_Selenium_Keys.htm
# https://www.twilio.com/en-us/messaging
# https://medium.com/data-science-community-srm/whatsapp-automation-with-selenium-a4fbaff625a0
$ErrorActionPreference = "Stop"
$DebugPreference = "Continue"

Set-Location $PSScriptRoot
$modules = @(
    "$($PSScriptRoot)\Modules\SeleniumWhatsApp.psm1" # Custom
)

# Import all modules necessary
$modules | ForEach-Object {
    Write-Output "Importing '$($_)'..."
    $_ | Import-Module -ErrorAction Stop
}

$CHROMEDRIVER_PATH = [System.IO.Path]::Combine(
    (Get-Item $PSScriptRoot).Parent.FullName,
    "utils"
)

# Kill all Chrome Browsers if they're open.
if (Get-Process chrome -ErrorAction SilentlyContinue) { Stop-Process -Force -ErrorAction SilentlyContinue }

# Start Google Chrome with a defined set of preferences
$CHROME_DOWNLOAD_DIR_PATH = $([System.IO.Path]::combine($PSScriptRoot,"Chrome.Downloads"))
$CHROME_USER_PREFS_PATH = $([System.IO.Path]::combine($PSScriptRoot,"Chrome.User_Data")) 

$driver = Start-ChromeDriver `
    -Chrome_Driver_Path $CHROMEDRIVER_PATH `
    -Download_Directory  $CHROME_DOWNLOAD_DIR_PATH `
    -User_Data $CHROME_USER_PREFS_PATH

Start-Sleep -Seconds 3

# Navigate to WhatsApp
$driver.Navigate().GoToUrl("https://web.whatsapp.com")

Read-Host "Press Enter after the QR Code is Entered, and the Main Menu is Launched"


# $hyperlinks = Get-WhatsAppMessageHyperLinks -Web_Driver $driver
# Test-WhatsAppAccount -Web_Driver $driver

# $contacts = @()

# Get-WhatsAppMessages `
#     -Web_Driver $driver `
#     -Contact_List $contacts

# #------------------------------- MESSAGE TEST -------------------------------#

# # Selecting own chat
# $user_page = Get-WhatsAppChat `
#     -Web_Driver $driver `
#     -Access_Method SearchBar `
#     -Name "Dontavious Jennings"

# $user_page.click()

# # Composing and sending Messages
# $message = Get-Content .\Input\numbers.txt -Raw

# Set-WhatsAppMessageBarText `
#     -Web_Driver $driver `
#     -Message $message 

# Send-WhatsAppMessage -Web_Driver $driver


# #------------------------------- MESSAGE TEST -------------------------------#

$chat_object = Get-WhatsAppChat `
    -Web_Driver $driver `
    -Access_Method SearchBar `
    -Search_Self

$phone_numbers = Import-Csv .\Input\

$phone_numbers = $phone_numbers.phone_number | Where-Object {$_ -ne $null}

$phone_numbers = $phone_numbers[0..100]

Select-WhatsAppElement -Element $chat_object

Set-WhatsAppMessage `
    -Web_Driver $driver `
    -Message $phone_numbers

Send-WhatsAppMessage `
    -Web_Driver $driver

# Sending attachments
$attachment_file_path = Get-Item ([System.IO.Path]::combine(
    "Input",
    "test.txt"
))

Set-WhatsAppAttachment `
    -Web_Driver $driver `
    -Path $attachment_file_path `
    -Message "Uh-oh, Spaghettio!" 

Send-WhatsAppAttachment `
    -Web_Driver $driver    

$Links = Get-WhatsAppMessageHyperLinks `
    -Web_Driver $driver

foreach ($link in $links) {
    $link | Add-Member `
        -MemberType NoteProperty `
        -Name "Has_WhatsApp" `
        -Value $(
            Test-WhatsAppAccount `
                -Web_Driver $driver `
                -Link_Element $link
        ) `
        -Force
}

$links | Out-GridView