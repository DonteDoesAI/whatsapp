$modules = @(
    "ImportExcel" # https://github.com/dfinke/ImportExcel
    "Selenium"    # https://github.com/adamdriscoll/selenium-powershell
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
    $_ | Import-Module 
}

# Globals
$SCROLL_SIZE = 9000

Function Clear-WhatsAppContactSearchBar {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][Object]$Web_Driver,
        [Parameter(Mandatory=$true)][Object]$WhatsApp_Search_Bar_Object
    )

    # Check to clear search bar in case it has text.
    if ($WhatsApp_Search_Bar_Object.Text -ne "") {
        $WhatsApp_Search_Bar_Object.SendKeys(
            [OpenQA.Selenium.Keys]::Control + `
            "A"
        )
        $WhatsApp_Search_Bar_Object.SendKeys(
            [OpenQA.Selenium.Keys]::Backspace
        )

        # Refresh the element selection - when we clear it, the element actually changes.
        $WhatsApp_Search_Bar_Object = $Web_Driver.FindElementbyXPath('//*[@id="side"]/div[1]/div/div/div[2]/div/div[1]/p')
        $WhatsApp_Search_Bar_Object.Click()
    }
    return $Web_Driver
}

Function Move-WhatsAppConversationPane {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][Object]$Web_Driver,
        [Parameter(Mandatory=$true)][Object]$Conversation_Pane,
        [Parameter(Mandatory=$false)][Object]$Page_Up_Max=50
    )
    $conversation_pane_last_height = $Conversation_Pane.size.height
    $pg_up_num_times = 0
    While ($true) {
        # send a Pg-Up command
        $Conversation_Pane.SendKeys([OpenQA.Selenium.Keys]::PageUp)
        if ($Conversation_Pane.Size.Height -eq $conversation_pane_last_height) {
            if ($pg_up_num_times -ge $Page_Up_Max) {
                Write-Host "End of conversation hit! Scraping..." -ForegroundColor Green
                break
            }
            $pg_up_num_times +=1 
            continue
        }
        Start-Sleep -Seconds 5
        $conversation_pane_last_height = $conversation_pane.Size.Height
        $pg_up_num_times = 0
    }

    return $Web_Driver
}

Function Get-WhatsAppMedia {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][Object]$Web_Driver,
        [Parameter(Mandatory=$true)][Object]$Subject_Folder_Path
    )

    # Download all media 
    try{ 
        # Click the media pane
        $media_pane = $Web_Driver.FindElementByXPath(
            '//*[@data-testid="section-media"]'
        )
        $media_pane.Click()

        Start-Sleep -Seconds 5

        # Click the 'Download' Button on all files
        # to preload the media; there is an empty try-catch block
        # that essentially ignores any errors when attempting to click a download button.
        $Web_Driver.FindElementsByXPath(
            '//*[@data-testid="media-download"]'
        ) | ForEach-Object {
            try {$_.Click()} catch {Write-Error $error[0]}
        }

        # 
        $media_files = $Web_Driver.FindElementsByXPath(
            '//*[@data-testid="media-canvas"]'
        ) 

        $media_files | ForEach-Object {
            $_.Click()
            Start-Sleep -Seconds 5

            $Web_Driver.FindElementsByXPath(
                '//*[@aria-label="Download"]'
            ).Click()
            Start-Sleep -Seconds 5

            $(
                (Get-ChildItem "C:\Users\$($env:USERNAME)\Downloads" `
                    | Sort-Object -Descending LastWriteTime)[0] `
                        | Move-Item -Destination $([System.IO.Path]::combine($Subject_Folder_Path,"Media")) -Force
            )

            $Web_Driver.FindElementsByXPath(
                '//*[@aria-label="Close"]'
            ).Click()
        }
    }
    catch {}
    return $Web_Driver
}

Function Get-WhatsAppContacts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][Object]$Web_Driver
    )

    try {
        Write-Host "Getting contact list..." -ForegroundColor Green
        
        $contacts = New-Object System.Collections.Generic.HashSet[String]
        
        $length = 0
    
        While ($true) {
            # Add contacts repeatedly.
            $Web_Driver.FindElementsByXPath('//div[@data-testid="cell-frame-title"]').text `
                | Sort-Object | ForEach-Object {
                    $contacts.add($_) | Out-Null
                }
    
            if ( ($length -eq $contacts.Count) -and $length -ne 0 ) {
                break
            } 
            else {
                $length = $contacts.Count
            }
    
            Move-WhatsAppPane `
                -Web_Driver $Web_Driver `
                -Pane pane-side
            Start-Sleep -Seconds 5
        }
    } catch {
        Write-Error $error[0]
        break
    }

    $deduped_sorted_contacts = $contacts | Where-Object {$_ -ne $null} | Sort-Object
    return $deduped_sorted_contacts
}

Function Move-WhatsAppPane {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][Object]$Web_Driver,
        [Parameter(Mandatory=$false)][ValidateSet("pane-side","pane-message")][Object]$Pane="pane-side"
    )
    Write-Host "Scrolling Side Pane..." -ForegroundColor Green

    
    $scroll_dict = [Ordered]@{
        "pane-side"   = $SCROLL_SIZE
        "pane-message"= $SCROLL_SIZE
    }

    $pane_class = ""
    switch ($Pane) {
        "pane-side" {
            $pane_class = "pane-side"
        }
        "pane-message" {
            $pane_class = "_2Ts6i _2xAQV"
        }
    }

    $pane_object = $Web_Driver.FindElementById(
        $pane_class
    )
    
    $Web_Driver.ExecuteScript(
        $("arguments[0].scrollTop = {0}" -f ($scroll_dict[$Pane])),
        $pane_object
    )
    Start-Sleep -Seconds 3
    $scroll_dict[$Pane] += $SCROLL_SIZE

    return $Web_Driver
}

Function Get-WhatsAppMessages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][Object]$Web_Driver,
        [Parameter(Mandatory=$true)][Object]$Contact_List
    )
    Write-Host "$($Contact_List.Count) contacts identified." -ForegroundColor Green
    Write-Host "Getting Messages..." -ForegroundColor Green

    $conversations = @{}

    foreach ($contact in $Contact_List) {
        Write-Host "Contact: $($contact)" -ForegroundColor Green

        Start-Sleep -Seconds 5

        # XPath to search bar:
        $search_bar = $Web_Driver.FindElementbyXPath('//*[@id="side"]/div[1]/div/div/div[2]/div/div[1]/p')
        $search_bar.Click()

        # Check to clear search bar in case it has text.
        if ($search_bar.Text -ne "") {
            $Web_Driver = Clear-WhatsAppContactSearchBar `
                -Web_Driver $Web_Driver `
                -WhatsApp_Search_Bar_Object $search_bar
        }

        $search_bar = $Web_Driver.FindElementbyXPath('//*[@id="side"]/div[1]/div/div/div[2]/div/div[1]/p')
        $search_bar.SendKeys($contact)
        
        Start-Sleep -Seconds 5

        $xpath = '//span[contains(@title, "{0}")]' -f ($contact)
        $user = $Web_Driver.FindElementsByXpath($xpath) `
            | Where-Object {$_.Text -eq $contact}

        $user.click()

        # Give 5 seconds for the page to render
        Start-Sleep -Seconds 5

        # Xpath to identify a conversation
        $MESSAGES_CLASS = "n5hs2j7m oq31bsqd gx1rr48f qh5tioqs"
        $xpath = "//div[@class='$($MESSAGES_CLASS)']"
        $conversation_pane = $Web_Driver.FindElementByXpath($xpath)

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

        try {
            $group_name = $Web_Driver.FindElementsByXPath(
                '//*[@id="app"]/div/div/div[6]/span/div/span/div/div/div/section/div[1]/div/div/div[2]'
            )
            if ($group_name -eq "") {
                $group_name = "NO_GROUP"
            }
        } catch {
            $group_name = "NO_GROUP"
        }

        if ($full_name -eq "") {
            $full_name = "NO_NAME"
        }

        $subject_folder_name = "$($contact).$($phone_number).$($full_name)"

        if ($contact -ne $full_name) {
            Write-Host "Stop!"
        }

        $subject_folder_path = [System.IO.Path]::Combine(
            $(Get-Item $PSScriptRoot `
                | Select-Object -ExpandProperty Parent `
                | Select-Object -ExpandProperty FullName
            ),
            "Output",
            "collected_data",
            "conversations",
            $subject_folder_name
        )
    
        if (!(Test-Path $subject_folder_path)) {
            New-Item $subject_folder_path -ItemType Directory
            New-Item ([System.IO.Path]::combine($subject_folder_path,"Media")) -ItemType Directory
        }

        # Helper method designed to trawl through the Web Page, and download all associated
        # media files.
        $Web_Driver = Get-WhatsAppMedia -Web_Driver $Web_Driver -Subject_Folder_Path $subject_folder_path

        $messages = New-Object System.Collections.Generic.HashSet[String]

        $length = 0
        
        # Use "Page Up" to move backwards in time to get more messages; somewhat
        # bad practice to use a discrete amount of times to measure distance but
        # this should catch 90% of cases.
        $Web_Driver = Move-WhatsAppConversationPane `
            -Web_Driver $Web_Driver `
            -Conversation_Pane $conversation_pane `
            -Page_Up_Max 50

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
            Start-Sleep -Seconds 5
        }
    }
    return $conversations

}