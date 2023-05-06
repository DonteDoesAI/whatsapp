# Required Modules
# https://chromedriver.chromium.org/downloads
# https://www.selenium.dev/selenium/docs/api/dotnet/html/T_OpenQA_Selenium_Keys.htm
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

# XPaths are key to finding elements in the XML - use the 'inspect' tool
# in Chrome to find and generate xpaths for data; there is no need for a whole
# lot of guesswork.
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

    $conversations = @{}

    foreach ($contact in $contact_list) {
        Write-Host "Contact: $($contact)" -ForegroundColor Green

        Start-Sleep -Seconds 2

        # XPath to search bar:
        $search_bar = $driver.FindElementbyXPath('//*[@id="side"]/div[1]/div/div/div[2]/div/div[1]/p')
        $search_bar.Click()

        # Check to clear search bar in case it has text.
        if ($search_bar.Text -ne "") {
            $search_bar.SendKeys(
                [OpenQA.Selenium.Keys]::Control + `
                "A"
            )
            $search_bar.SendKeys(
                [OpenQA.Selenium.Keys]::Backspace
            )

            # Refresh the element selection - when we clear it, the element actually changes.
            $search_bar = $driver.FindElementbyXPath('//*[@id="side"]/div[1]/div/div/div[2]/div/div[1]/p')
            $search_bar.Click()
        }

        $search_bar.SendKeys($contact)

        $xpath = '//span[contains(@title, "{0}")]' -f ($contact)
        $user = $Web_Driver.FindElementByXpath($xpath)

        Start-Sleep -Seconds 5

        $user.click()

        # Give 3 seconds for the page to render
        Start-Sleep -Seconds 3

        # Xpath to identify a conversation
        $xpath = "//div[@class='$($MESSAGES_CLASS)']"
        $conversation_pane = $Web_Driver.FindElementByXpath($xpath)
        $conversation_pane_last_height = $conversation_pane_size.size.height

        $user_details = $Web_Driver.FindElementByXpath("/html/body/div[1]/div/div/div[5]/div/header/div[1]")
        $user_details.click()

        Start-Sleep -Seconds 5

        try {
            $phone_number = $Web_Driver.FindElementByXPath("/html/body/div[1]/div/div/div[6]/span/div/span/div/div/section/div[1]/div/div[2]/div").Text
        } catch {
            $phone_number = "NO_PHONE"
        }

        try {
            $full_name = $Web_Driver.FindElementByXpath('//*[@id="app"]/div/div/div[6]/span/div/span/div/div/section/div[1]/div/div[2]/h2').Text
        } catch {
            $full_name = "NO_NAME"
        }

        if ($full_name -eq "") {
            $full_name = "NO_NAME"
        }

        $subject_folder_name = "$($phone_number).$($full_name)"

        $subject_folder_path = [System.IO.Path]::Combine(
            $PSScriptRoot,
            "Output",
            "collected_data",
            "conversations",
            $subject_folder_name
        )
    
        if (!(Test-Path $subject_folder_path)) {
            New-Item $subject_folder_path -ItemType Directory
            New-Item ([System.IO.Path]::combine($subject_folder_path,"Media")) -ItemType Directory
        }

        # Download all media 
        try{ 
            # Click the media pane
            $media_pane = $Web_Driver.FindElementByXPath(
                '//*[@data-testid="section-media"]'
            )
            $media_pane.Click()

            # Click the 'Download' Button on all files
            # to preload the media; there is an empty try-catch block
            # that essentially ignores any errors when attempting to click a download button.
            $Web_Driver.FindElementsByXPath(
                '//*[@data-testid="media-download"]'
            ) | ForEach-Object {
                try {$_.Click()} catch {}
            }

            # 
            $media_files = $Web_Driver.FindElementsByXPath(
                '//*[@data-testid="media-canvas"]'
            ) 

            $media_files | ForEach-Object {
                $_.Click()
                Start-Sleep -Seconds 3

                $Web_Driver.FindElementsByXPath(
                    '//*[@aria-label="Download"]'
                ).Click()
                Start-Sleep -Seconds 3

                $(
                    (Get-ChildItem "C:\Users\$($env:USERNAME)\Downloads" `
                        | Sort-Object -Descending LastWriteTime)[0] `
                            | Move-Item -Destination $([System.IO.Path]::combine($subject_folder_path,"Media")) -Force
                )

                $Web_Driver.FindElementsByXPath(
                    '//*[@aria-label="Close"]'
                ).Click()
            }
        }
        catch {}


        $messages = New-Object System.Collections.Generic.HashSet[String]

        $length = 0

        # $scroll = $SCROLL_SIZE
        
        # Use "Page Up" to move backwards in time to get more messages; somewhat
        # bad practice to use a discrete amount of times to measure distance but
        # this should catch 90% of cases.
        $pg_up_num_times = 0
        While ($true) {
            $conversation_pane.SendKeys([OpenQA.Selenium.Keys]::PageUp)
            if ($conversation_pane.Size.Height -eq $conversation_pane_last_height) {
                if ($pg_up_num_times -ge 10) {
                    Write-Host "End of conversation hit! Scraping..." -ForegroundColor Green
                    break
                }
                $pg_up_num_times +=1 
                continue
            }
            $conversation_pane_last_height = $conversation_pane.Size.Height
            $pg_up_num_times = 0
        }

        While ($true) {
            # Get all rows that are rendered; remove rows that have no text.
            $elements = $Web_Driver.FindElementsByXpath("//div[@class='copyable-text']") | `
                Where-Object {$_.TagName -eq "div"}

            foreach ($e in $elements) {
                $header = $e.GetAttribute('data-pre-plain-text')
                $message_body = $e.text
                $full_message = "$($header.trim()) $($message_body)"
                $messages.Add($full_message)               
                Write-Host $full_message -ForegroundColor Green
            }

            # If all messages have been collected, export to a json and continue with
            # the next contact.
            if ($length -eq $messages.Count) {
                $conversations[$contact] = $messages
            
                $conversation_file_path = [System.IO.Path]::Combine(
                    $subject_folder_path,
                    "conversation.txt"
                )

                if (Test-Path $conversation_file_path) {
                    Remove-Item $conversation_file_path
                }
            
                $json_string = $messages | ConvertTo-Json -Depth 100
                $json_string | Out-File $conversation_file_path
                # Write-Host "Messages -> $($messages)" -ForegroundColor White
                break
            }

            else {
                $length = $messages.Count
            }
            Start-Sleep -Seconds 2
        }
    }
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
$downloads_folder = [System.IO.Path]::Combine(
    $PSScriptRoot,
    "Temp"
)
New-Item $downloads_folder -ItemType Directory -ErrorAction Ignore

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
    }
    Write-Host "$($contacts.count) contacts retrieved!" -ForegroundColor Green

    Write-Host "$($conversations.count) conversations retreived!" -ForegroundColor Green

    $subject_folder_path = [System.IO.Path]::Combine(
        "collected_data"
    )
    if (!(Test-Path $subject_folder_path)) {
        New-Item $subject_folder_path -ItemType Directory
    }
    $file_path = [System.IO.Path]::Combine(
        $subject_folder_path,
        "all.json"
    )

    $all_conversations = $conversations | ConvertTo-Json -Depth 100

    $all_conversations | Out-File $file_path

}
catch {
    $Error[0]
    $driver.Quit()
}