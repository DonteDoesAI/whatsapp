# Required Modules
# https://chromedriver.chromium.org/downloads - need a chrome driver that matches chrome version
# https://www.selenium.dev/selenium/docs/api/dotnet/html/T_OpenQA_Selenium_Keys.htm
$modules = @(
    "$($PSScriptRoot)\Modules\WhatsApp-Selenium.psm1" # Custom
)

# If the modules don't exist, install for the current scope
$modules | ForEach-Object {
    if (!(Get-Module -ListAvailable -Name $_)) {
        Write-Host "'$($_)' does not exist; installing..." -ForegroundColor Green;
        Install-Module $_ `
            -Scope CurrentUser `
            -ErrorAction Stop `
            -Force 
    }
    else {
        Write-Host "'$($_)' exists!" -ForegroundColor Green;
    }
}

# Import all modules necessary
$modules | ForEach-Object {
    Write-Host "Importing '$($_)'..." -ForegroundColor Green;
    $_ | Import-Module -ErrorACtion Stop
}

$CHROMEDRIVER_PATH = [System.IO.Path]::Combine(
    ".",
    "utils"
)

$user_data_path = [System.IO.Path]::Combine(
    $PSScriptRoot,
    "$($driver_choice).User_Data"
)

$whatsapp_url = "https://web.whatsapp.com/"

$arguments = @(
    "user-data-dir=$($user_data_path)"
)

$ChromeOptions = New-Object OpenQA.Selenium.Chrome.ChromeOptions
$ChromeOptions.AddArguments($arguments)

# Starting a driver with the given cmdlet is a pain in the butt; apparently calling the C# method directly
# is much more stable, using the Chrome Driver DIRECTORY (vice the explicit path).
$driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($CHROMEDRIVER_PATH, $ChromeOptions)
Enter-SeUrl $whatsapp_url -Driver $driver # Navigate to WhatsApp
Read-Host "Press Enter after the QR Code is Entered"

try {
    Write-Host "Getting contact list..." -ForegroundColor Green
    
    # Helper method that returns a hashset of contacts in WhatsApp
    $contacts = Get-WhatsAppContacts -Web_Driver $driver
    $contacts = $contacts | Where-Object {$_ -ne $null} | Where-Object {$_.GetType().Name -eq "String"}
    
    try {
        # Helper method that exports conversations + media in WhatsApp
        $conversations = Get-WhatsAppMessages `
            -Web_Driver $driver `
            -Contact_List $contacts
    }
    catch {
        Write-Error $error[0]
    }
        

    Write-Host "$($contacts.count) contacts retrieved!" -ForegroundColor Green

    Write-Host "$($conversations.count) conversations retreived!" -ForegroundColor Green

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