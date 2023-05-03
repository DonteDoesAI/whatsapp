# Required Modules
# https://chromedriver.chromium.org/downloads
$modules = @(
    "ImportExcel" # https://github.com/dfinke/ImportExcel
    "Selenium"    # https://github.com/adamdriscoll/selenium-powershell
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


$CONTACTS_CLASS = "_21S-L"
$MESSAGES_CLASS = "n5hs2j7m oq31bsqd gx1rr48f qh5tioqs" #"_2gzeB"
$MESSAGE_PANE

function Get-Messages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][Object]$Web_Driver,
        [Parameter(Mandatory=$true)][Object]$Contact_List
    )

    Write-Host "Getting Messages..." -ForegroundColor Green

    foreach ($contact in $contact_list) {
        Write-Host "Contact: $($contact)" -ForegroundColor Green

        Start-Sleep -Seconds 2

        # XPath to search bar:
        $search_bar = $driver.FindElementbyXPath('//*[@id="side"]/div[1]/div/div/div[2]/div/div[1]/p')
        $search_bar.Click()

        #FIXME - need to figure out how to remove text from search bar.
        if ($search_bar.Text -ne "") {
            $search_bar.SendKeys([System.Windows.Forms.SendKeys]::SendWait("^a")) | Out-Null # Clear previous searches
            Start-Sleep -Seconds 2
        }

        $search_bar.SendKeys($contact)

        # Xpath to click the conversation head
        $xpath = '//span[contains(@title, "{0}")]' -f ($contact)
        $user = $Web_Driver.FindElementByXpath($xpath)
        $user.click()

        # Give 3 seconds for the page to render
        Start-Sleep -Seconds 3
        # Xpath to identify a conversation
        $xpath = "//div[@class='$($MESSAGES_CLASS)']"
        $conversation_pane = $Web_Driver.FindElementByXpath($xpath)

        $messages = New-Object System.Collections.Generic.HashSet[String]

        $length = 0

        $scroll = $SCROLL_SIZE

        While ($true) {

            # Get all rows that are rendered; remove rows that have no text.
            $elements = $Web_Driver.FindElementsByXpath("//div[@class='copyable-text']") | `
                Where-Object {$_.TagName -eq "div"}

            foreach ($e in $elements) {
                $header = $e.GetAttribute('data-pre-plain-text')
                $message_body = $e.text
                $full_message = "$($header.trim()) $($message_body)"
                $messages.Add($full_message)
                Write-Host $messages -ForegroundColor Green
            }

            if ($length -eq $messages.Count) {
                break
            }

            else {
                $length = $messages.Count
            }

            $driver.ExecuteScript(
                ('arguments[0].scrollTop = -{0}' -f ($scroll)), 
                $conversation_pane
            )
            Start-Sleep -Seconds 2
            $scroll += $SCROLL_SIZE
        }
        Scroll-Pane -Web_Driver $Web_Driver
    }

    Write-Host "Messages -> $($messages)" -ForegroundColor White
    
    $conversations.add($messages)

    $file_directory = [System.IO.Path]::Combine(
        "Output",
        "collected_data",
        "conversations"
    )

    if (!(Test-Path $file_directory)) {
        New-Item $file_directory -ItemType Directory
    }

    $file_path = [System.IO.Path]::Combine(
        $file_directory,
        "$($contact).json"
    )

    $json_string = $messages | ConvertTo-Json -Depth 100

    $json_string | Out-File $file_path

    return $conversations

}

$scroll_dict = [Ordered]@{
    "pane-side"   = $SCROLL_SIZE
    "pane-message"= $SCROLL_SIZE
}

Function Scroll-Pane {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][Object]$Web_Driver,
        [Parameter(Mandatory=$false)][ValidateSet("pane-side","pane-message")][Object]$Pane="pane-side"
    )
    Write-Host "Scrolling Side Pane..." -ForegroundColor Green

    $pane_class = ""
    switch ($Pane) {
        "pane-side" {
            $pane_class = "pane-side"
        }
        "pane-message" {
            $pane_class = "_2Ts6i _2xAQV"
        }
    }

    $side_pane = $Web_Driver.FindElementById(
        $pane_class
    )
    $Web_Driver.ExecuteScript(
        $("arguments[0].scrollTop = {0}" -f ($scroll_dict[$Pane])),
        $side_pane
    )
    Start-Sleep -Seconds 3
    $scroll_dict[$Pane] += $SCROLL_SIZE

    return
}

# Globals
$SCROLL_TO = 600
$SCROLL_SIZE = 600
$conversations = [System.Collections.ArrayList]::new()

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
    
    $contacts = New-Object System.Collections.Generic.HashSet[String]
    
    $length = 0

    While ($true) {
        $contacts_sel = $driver.FindElementsByClassName($CONTACTS_CLASS)

        $contacts_sel = $contacts_sel.text

        $conversations.Add(
            (Get-Messages -Web_Driver $driver -Contact_List ( $contacts_sel | Where-Object {$_ -notin $contacts}) )
        )

        $contacts.Add($contacts_sel)

        if ( ($length -eq $contacts.Count) -and $length -ne 0 ) {
            break
        } 
        else {
            $length = $contacts.Count
        }

        Scroll-Pane -Web_Driver $driver -Pane "pane-side"
    }
    Write-Host "$($contacts.count) contacts retrieved!" -ForegroundColor Green

    Write-Host "$($conversations.count) conversations retreived!" -ForegroundColor Green

    $file_directory = [System.IO.Path]::Combine(
        "collected_data"
    )
    if (!(Test-Path $file_directory)) {
        New-Item $file_directory -ItemType Directory
    }
    $file_path = [System.IO.Path]::Combine(
        $file_directory,
        "all.json"
    )

    $all_conversations = $conversations | ConvertTo-Json -Depth 100

    $all_conversations | Out-File $file_path

}
catch {
    $Error[0]
    $driver.Quit()
}