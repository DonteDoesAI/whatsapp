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
    $_ | Import-Module -ErrorACtion Stop
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


# $user_page = Get-WhatsAppChat -Web_Driver $driver -Access_Method SnapViaURL -Phone_Number "15712150398"

# $user_page.click()
# Messages
$message = Get-Content .\Input\test.txt -Raw

$message_bar = Get-WhatsAppMessageBar -Web_Driver $driver
$message_bar.SendKeys($message)
$message_bar.SendKeys([OpenQA.Selenium.Keys]::Enter)

$attachment_file_path = Get-Item ([System.IO.Path]::combine(
    "Input",
    "test.txt"
))

$attachment_button = $driver.FindElementByXPath(
    '//div[@title = "Attach"]'
)
$attachment_button.click()

$image_box = $driver.FindElementByXPath(
    '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]'
)

$image_box.SendKeys($attachment_file_path.FullName)

Start-Sleep 3

$send_button = $driver.FindElementByXPath('//span[@data-icon="send"]')
$send_button.Click()

# $contacts = @()

# Get-WhatsAppMessages `
#     -Web_Driver $driver `
#     -Contact_List $contacts