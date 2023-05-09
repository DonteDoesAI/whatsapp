# Required Modules
# https://chromedriver.chromium.org/downloads - need a chrome driver that matches chrome version
# https://www.selenium.dev/selenium/docs/api/dotnet/html/T_OpenQA_Selenium_Keys.htm
$ErrorActionPreference = "Stop"
$DebugPreference = "Continue"

Set-Location $PSScriptRoot
$modules = @(
    "$($PSScriptRoot)\Modules\SeleniumWhatsApp.psm1" # Custom
)

# If the modules don't exist, install for the current scope
$modules | ForEach-Object {
    if (!(Get-Module -ListAvailable -Name $_)) {
        Write-Output "'$($_)' does not exist; installing...";
        Install-Module $_ `
            -Scope CurrentUser `
            -ErrorAction Stop `
            -Force 
    }
    else {
        Write-Output "'$($_)' exists!"
    }
}

# Import all modules necessary
$modules | ForEach-Object {
    Write-Output "Importing '$($_)'..."
    $_ | Import-Module -ErrorACtion Stop
}

$CHROMEDRIVER_PATH = [System.IO.Path]::Combine(
    (Get-Item $PSScriptRoot).Parent.FullName,
    "utils"
)

$CHROME_DOWNLOAD_DIR_PATH = $([System.IO.Path]::combine($PSScriptRoot,"Chrome.Downloads"))
$CHROME_USER_PREFS_PATH = $([System.IO.Path]::combine($PSScriptRoot,"Chrome.User_Data")) 

$driver = Start-ChromeDriver `
    -Chrome_Driver_Path $CHROMEDRIVER_PATH `
    -Download_Directory  $CHROME_DOWNLOAD_DIR_PATH `
    -User_Data $CHROME_USER_PREFS_PATH

Start-Sleep -Seconds 3

# Navigate to WhatsApp
$driver.Navigate().GoToUrl("https://web.whatsapp.com")

Read-Host "Press Enter after the QR Code is Entered"

try {
    Write-Output "Getting contact list..." -ForegroundColor Green
    
    # Helper method that returns a hashset of contacts in WhatsApp
    $contacts = Get-WhatsAppContacts -Web_Driver $driver `
    | Where-Object {$_ -ne $null} `
        | Where-Object {$_.GetType().Name -eq "String"}
    
    try {
        # Helper method that exports conversations + media in WhatsApp
        $conversations = Get-WhatsAppMessages `
            -Web_Driver $driver `
            -Contact_List $contacts
    }
    catch {
        Write-Error $error[0]
    }
        

    Write-Output "$($contacts.count) contacts retrieved!" -ForegroundColor Green

    Write-Output "$($conversations.count) conversations retreived!" -ForegroundColor Green

    $output_folder_path = [System.IO.Path]::Combine(
        "collected_data"
    )
    if (!(Test-Path $output_folder_path)) {
        New-Item $output_folder_path -ItemType Directory
    }
    $file_path = [System.IO.Path]::Combine(
        $output_folder_path,
        "all.json"
    )

    $all_conversations = $conversations | ConvertTo-Json -Depth 100
    $all_conversations | Out-File $file_path

    $driver.quit()

}
catch {
    $Error[0]
    $driver.Quit()
}