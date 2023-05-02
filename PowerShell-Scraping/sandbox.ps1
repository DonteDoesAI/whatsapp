# Required Modules
# https://chromedriver.chromium.org/downloads
$modules = @(
    "ImportExcel" # https://github.com/dfinke/ImportExcel
    "Selenium"    # https://github.com/adamdriscoll/selenium-powershell
)

# $CHROMEBROWSER_PATH = $(Get-Package -Name "*Chrome").swidtags.Metadata["DisplayIcon"].split(",")[0]

$CHROMEDRIVER_PATH = [System.IO.Path]::Combine(
    ".",
    "utils"
)

# If the modules don't exist, install for the current scope
$modules | ForEach {
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
$modules | ForEach {
    Write-Host "Importing '$($_)'..." -ForegroundColor Green;
    $_ | Import-Module 
}

# Expose all options available for PS Selenium
$driver_options = @(
    "Edge"
    "Chrome"
    "Firefox"
)

$driver_choice = $driver_options[1]
Write-Host "USing '$($driver_choice)' as the Selenium Driver..." -ForegroundColor Green

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