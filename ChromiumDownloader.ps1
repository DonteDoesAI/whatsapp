# Get Chrome Version
Set-Location $PSScriptRoot
$ErrorActionPreference = "Stop"
try {
    $Chrome_Application = (
        Get-Item `
            (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe').'(Default)'
    ).VersionInfo
} catch {
    Write-Error -Message "Can't find Chrome! Did you install it?"
}

$Chrome_Driver_Download_Page = Invoke-WebRequest -Uri "https://chromedriver.chromium.org/downloads"

# Check what Chrome Drivers are available, using a discrete class.
$hrefs = $Chrome_Driver_Download_Page.ParsedHtml.Links `
    | Where-Object {$_.className -like "XqQF9c"} `
        | Where-Object {$_.ie8_href -like "https://chromedriver.storage.googleapis.com/index.html?path=*"} `
            | Select-Object -Property @( # Alias the fieldnames, where E is the original field.
                @{N="Version";E={[System.Version]($_.innerText.replace("ChromeDriver","").trim())}}, # Cast as a version object
                @{N="Link";E={$_.ie8_href}}
            ) -Unique

# Select drivers where the major revision matches the version of Google Chrome installed,
# Cast the Application as a System.Version Object, and place most recent version
#  at the top of the list
$valid_drivers = $hrefs `
    | Where-Object {
        $_.Version.Major `
            -eq `
        ([System.Version]$Chrome_Application.FileVersion).Major 
    } `
    | Sort-Object -Property Version -Descending 

$selected_driver = $valid_drivers[0]

$driver_uri = "https://chromedriver.storage.googleapis.com/$($selected_driver.Version)/chromedriver_win32.zip"
$zip_path = $([System.IO.Path]::combine("utils","chromedriver_win32.zip"))
Write-Debug $zip_path
# Download the correct chrome driver.

Invoke-WebRequest `
    -Uri $driver_uri `
    -OutFile $zip_path

Expand-Archive $zip_path -DestinationPath "utils" -Force