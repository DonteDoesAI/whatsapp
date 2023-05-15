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


# Selecting own chat
$user_page = Get-WhatsAppChat `
    -Web_Driver $driver `
    -Access_Method SearchBar `
    -Search_Self

$user_page.click()

# Composing and sending Messages
$message = Get-Content .\Input\numbers.txt -Raw

Set-WhatsAppMessageBarText `
    -Web_Driver $driver `
    -Message $message 

$message_bar = Get-WhatsAppMessageBar -Web_Driver $driver 
$message_bar.SendKeys([OpenQA.Selenium.Keys]::Enter)


# Sending attachments
$attachment_file_path = Get-Item ([System.IO.Path]::combine(
    "Input",
    "test.txt"
))

$attachment_bar = Set-WhatsAppAttachment `
    -Web_Driver $driver `
    -Path $attachment_file_path `
    -Message "Uh-oh, Spaghettio!" 

$attachment_bar.SendKeys([OpenQA.Selenium.Keys]::Enter)


# Test

Function Get-WhatsAppMessageHyperLinks {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][Object]$Web_Driver
    )
    Write-Debug "Get-WhatsAppMessageHyperLinks"
    $text_class = "_11JPr selectable-text copyable-text"
    $hyperlink_style = "cursor: pointer;"

    $xpath = '//*[@class="{0}" and @style="{1}"]' -f ($text_class, $hyperlink_style)

    $hyperlinked_messages = $driver.FindElementsByXPath($xpath)

    return $hyperlinked_messages
}

# FIXME - for some reason this doesn't return the correct answer but will parse correctly if run outside of a function
Function Test-WhatsAppAccount {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][Object]$Web_Driver
    )
    Write-Debug "Test-WhatsAppAccount"
    $xpath = '//*[@data-testid="mi-message-on-whatsapp"]'
    
    try {
        $element = $Web_Driver.FindElementsByXPath(
            $xpath
        )
        return $true
    }
    catch {
        Write-Debug "Failure!"
        return $false
    }
}

# $hyperlinks = Get-WhatsAppMessageHyperLinks -Web_Driver $driver
# Test-WhatsAppAccount -Web_Driver $driver

# $contacts = @()

# Get-WhatsAppMessages `
#     -Web_Driver $driver `
#     -Contact_List $contacts