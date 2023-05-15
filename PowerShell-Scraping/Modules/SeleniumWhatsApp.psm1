$modules = @(
    "ImportExcel" # https://github.com/dfinke/ImportExcel
    "Selenium"    # https://github.com/adamdriscoll/selenium-powershell
)

# If the modules don't exist, install for the current scope
$modules | ForEach-Object {
    if (!(Get-Module -ListAvailable -Name $_)) {
        Write-Debug "'$($_)' does not exist; installing..." 
        Install-Module $_ `
            -Scope CurrentUser `
            -ErrorAction Stop `
            -Force 
    }
    else {
        Write-Debug "'$($_)' exists!" 
    }
}

# Import all modules necessary
$modules | ForEach-Object {
    Write-Debug "Importing '$($_)'..." 
    $_ | Import-Module 
}

Function Start-ChromeDriver {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][Object]$Chrome_Driver_Path,
        [Parameter(Mandatory=$true)][Object]$User_Data,
        [Parameter(Mandatory=$true)][Object]$Download_Directory
    )
    Write-Debug "Start-ChromeDriver"

    if (!(Test-Path $Download_Directory)) {
        New-Item -ItemType Directory $Download_Directory | Out-Null
    }

    $ChromeOptions = New-Object OpenQA.Selenium.Chrome.ChromeOptions
    $ChromeOptions.AddArguments("user-data-dir=$($User_Data)")
    $ChromeOptions.AddArguments("--browser.download.folderList=2");
    $ChromeOptions.AddArguments("--browser.helperApps.neverAsk.saveToDisk=image/jpg");
    $ChromeOptions.AddArguments("--browser.download.dir=$($Download_Directory)");
    $ChromeOptions.AddUserProfilePreference("download.default_directory", $Download_Directory);
    
    # Starting a driver with the given cmdlet is a pain in the butt; apparently calling the C# method directly
    # is much more stable, using the Chrome Driver DIRECTORY (vice the explicit path).
    $driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($Chrome_Driver_Path, $ChromeOptions)

    return $driver
}

Function Clear-WhatsAppContactSearchBar {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][Object]$Web_Driver,
        [Parameter(Mandatory=$true)][Object]$WhatsApp_Search_Bar_Object
    )
    Write-Debug "Clear-WhatsAppContactSearchBar"

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
    Write-Debug "Move-WhatsAppConversationPane"
    $conversation_pane_last_height = $Conversation_Pane.size.height
    $pg_up_num_times = 0
    While ($true) {
        # send a Pg-Up command
        $Conversation_Pane.SendKeys([OpenQA.Selenium.Keys]::PageUp)
        if ($Conversation_Pane.Size.Height -eq $conversation_pane_last_height) {
            if ($pg_up_num_times -ge $Page_Up_Max) {
                # Write-Debug "End of conversation hit! Scraping..."                 
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

Function Get-WhatsAppFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][Object]$Web_Driver,
        [Parameter(Mandatory=$true)][Object]$Subject_Folder_Path,
        [Parameter(Mandatory=$false)][ValidateSet("Media","Docs")][Object]$File_Type
    )
    Write-Debug "Get-WhatsAppFiles, File_Type: $($File_Type)"

    # Download all media 
    try{ 
        # Click the media pane (if it isn't clicked already)
        try {
            $media_links_docs_pane = $Web_Driver.FindElementByXPath(
                '//*[@data-testid="section-media"]'
            )
            $media_links_docs_pane.Click()

            # Start-Sleep -Seconds 5
        }
        catch {}

        $media_pane = $Web_Driver.FindElementByXPath(
            '//*[@title="{0}"]' -f ($File_Type)
        )
        $media_pane.Click()

        Switch ($File_Type) {
            ("Media") {
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

                Write-Debug "$($media_files.count) Media files found."

                # Routine to download items
                foreach ($media_file in $media_files) {
                    $media_file.click()
                    Start-Sleep -Seconds 5

                    $Web_Driver.FindElementsByXPath(
                        '//*[@aria-label="Download"]'
                    ).Click()
                    Start-Sleep -Seconds 5

                    $download_folder = $(
                        [System.IO.Path]::combine(
                            (Get-Location).Path,
                            "Chrome.Downloads"
                        )
                    )
                    
                    $most_recent_download = (Get-ChildItem $download_folder `
                        | Sort-Object -Descending LastWriteTime)[0]

                    $most_recent_download | Move-Item -Destination $(
                        [System.IO.Path]::combine($Subject_Folder_Path,$File_Type)
                    )

                    $Web_Driver.FindElementsByXPath(
                        '//*[@aria-label="Close"]'
                    ).Click()
                }
            }
            ("Docs") {
                $docs_download_class = "_233fJ"

                Start-Sleep -Seconds 5

                $docs = $Web_Driver.FindElementsByXPath(
                    '//*[@class="{0}"]' -f $($docs_download_class)
                ) 
                Write-Debug "$($docs.count) Doc files found."
                # Routine to download items
                foreach ($doc in $docs) {
                    $doc.click()
                    Start-Sleep -Seconds 5

                    $download_folder = $(
                        [System.IO.Path]::combine(
                            (Get-Location).Path,
                            "Chrome.Downloads"
                        )
                    )

                    $most_recent_download = (Get-ChildItem $download_folder `
                        | Sort-Object -Descending LastWriteTime)[0]

                    $most_recent_download | Move-Item -Destination $(
                        [System.IO.Path]::combine($Subject_Folder_Path,$File_Type)
                    )
                }                
            }
        }
    }
    catch {}

    $back_button = $Web_Driver.FindElementByXPath(
        '//*[@id="app"]/div/div/div[6]/span/div/span/div/header/div/div[1]/div/span'
    )
    $back_button.click()

    return $Web_Driver
}

Function Get-WhatsAppContacts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][Object]$Web_Driver,
        [Parameter(Mandatory=$false)][Object]$Max_Times_Check=5
    )
    Write-Debug "Get-WhatsAppContacts"

    try {
        Write-Debug "Getting contact list..."
        
        $contacts = New-Object System.Collections.Generic.HashSet[String]

        $scroll_level = 0

        $max_height = Get-WhatsAppScrollMaxHeight `
        -Web_Driver $Web_Driver `
        -Arguments $(
            Get-WhatsAppScrollBar `
                -Web_Driver $Web_Driver `
                -Pane pane-side
        )

        while ($true) {
            $scroll_level += 600
            $Web_Driver.FindElementsByXPath('//div[@data-testid="cell-frame-title"]').text `
            | Sort-Object | ForEach-Object {
                $contacts.add($_) | Out-Null
            }

            if ($scroll_level -ge $max_height) {
                break
            }

            # Iteratively scroll to the bottom
            Move-WhatsAppPane `
                -Web_Driver $Web_Driver `
                -Pane pane-side `
                -Scroll_Level $scroll_level

            Start-Sleep -Seconds 2
        }

    } catch {
        Write-Error $error[0]
        break
    }

    $deduped_sorted_contacts = $contacts `
        | Where-Object {$_ -ne $null} `
            | Where-Object {$_.GetType().Name -eq "String"} `
                | Where-Object {$_ -notlike "*(You)"} `
                    | Sort-Object
    return $deduped_sorted_contacts
}
Function Get-WhatsAppScrollMaxHeight {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][Object]$Web_Driver,
        [Parameter(Mandatory=$true)][Object]$Arguments
    )
    Write-Debug "Get-WhatsAppScrollMaxHeight"
    return $Web_Driver.ExecuteScript(
        "return arguments[0].scrollHeight",
        $Arguments
    )
}
Function Get-WhatsAppScrollBar {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][Object]$Web_Driver,
        [Parameter(Mandatory=$false)][ValidateSet("pane-side","pane-message")][Object]$Pane="pane-side"
    )
    Write-Debug "Get-WhatsAppScrollBar"
    $pane_class = ""
    switch ($Pane) {
        "pane-side" {
            $pane_class = "pane-side"
            $pane_object = $Web_Driver.FindElementById(
                $pane_class
            )
        }
        "pane-message" {
            $pane_class = "n5hs2j7m oq31bsqd gx1rr48f qh5tioqs"
            $pane_object = $Web_Driver.FindElementByXPath(
               '//*[@class="{0}"]' -f ($pane_class)
            )
        }
    }

    return $pane_object 
}
Function Get-WhatsAppSearchBar {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]$Web_Driver
    )
    Write-Debug "Get-WhatsAppSearchBar"
    $xpath = '//*[@title="Search input textbox"]'
    $search_bar = $Web_Driver.FindElementByXPath($xpath)
    return $search_bar
}
Function Move-WhatsAppPane {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][Object]$Web_Driver,
        [Parameter(Mandatory=$false)][ValidateSet("pane-side","pane-message")][Object]$Pane="pane-side",
        [Parameter(Mandatory=$true)][Object]$Scroll_Level
    )
    Write-Debug "Move-WhatsAppPane"
    # Write-Debug "Scrolling $($pane_class)..." | Out-Null

    $pane_object = Get-WhatsAppScrollBar -Web_Driver $Web_Driver -Pane $Pane

    # Scroll by a discrete amount
    $Web_Driver.ExecuteScript(
        "arguments[0].scrollTop = $($Scroll_Level)",
        $pane_object
    )

    return $Web_Driver
}
Function Get-WhatsAppMessages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][Object]$Web_Driver,
        [Parameter(Mandatory=$true)][Object]$Contact_List,
        [Parameter(Mandatory=$false)][Switch]$Export_Conversations=$false
    )
    Write-Debug "Get-WhatsAppMessages"
    Write-Debug "$($Contact_List.Count) contact(s) identified."
    Write-Debug "Getting Messages..." 
    $conversations = @{}

    foreach ($contact in $Contact_List) {
        Write-Debug "Contact: $($contact)" 

        $search_bar = Get-WhatsAppSearchBar -Web_Driver $Web_Driver
        $search_bar.Click()

        # Check to clear search bar in case it has text.
        if ($search_bar.Text -ne "") {
            $Web_Driver = Clear-WhatsAppContactSearchBar `
                -Web_Driver $Web_Driver `
                -WhatsApp_Search_Bar_Object $search_bar
        }

        $search_bar = Get-WhatsAppSearchBar -Web_Driver $Web_Driver
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
    
        if (!(Test-Path $subject_folder_path) -and $Export_Conversations) {
            New-Item $subject_folder_path -ItemType Directory
            New-Item ([System.IO.Path]::combine($subject_folder_path,"Media")) -ItemType Directory
            New-Item ([System.IO.Path]::combine($subject_folder_path,"Docs")) -ItemType Directory
        }

        # Helper method designed to trawl through the Web Page, and download all associated
        # media files.
        if ($Export_Conversations) {
            $Web_Driver = Get-WhatsAppFiles -Web_Driver $Web_Driver -Subject_Folder_Path $subject_folder_path -File_Type Media
            $Web_Driver = Get-WhatsAppFiles -Web_Driver $Web_Driver -Subject_Folder_Path $subject_folder_path -File_Type Docs
        }
        $messages = New-Object System.Collections.Generic.HashSet[String]

        $length = 0
        
        # Use "Page Up" to move backwards in time to get more messages; somewhat
        # bad practice to use a discrete amount of times to measure distance but
        # this should catch 90% of cases.
        $Web_Driver = Move-WhatsAppConversationPane `
            -Web_Driver $Web_Driver `
            -Conversation_Pane $conversation_pane `
            -Page_Up_Max 100

        While ($true) {
            # Get all rows that are rendered; remove rows that have no text.
            $elements = $Web_Driver.FindElementsByXpath("//div[@class='copyable-text']") | `
                Where-Object {$_.TagName -eq "div"}

            foreach ($e in $elements) {
                $header = $e.GetAttribute('data-pre-plain-text')
                $message_body = $e.text
                $full_message = "$($header.trim()) $($message_body)"
                $messages.Add($full_message) | Out-Null
                Write-Debug $messages[0]
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
                if ($Export_Conversations) {
                    $json_string = $messages | ConvertTo-Json -Depth 100
                    $json_string | Out-File $conversation_file_path
                }
                # # Write-Debug "Messages -> $($messages)"                 
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

Function Get-WhatsAppChat {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,ParameterSetName="SearchBar")]
        [Parameter(Mandatory=$true,ParameterSetName="SnapViaURL")]
        [Object]$Web_Driver,
        
        [Parameter(Mandatory=$true,ParameterSetName="SnapViaURL")]
        [Object]$Phone_Number,

        [Parameter(Mandatory=$false,ParameterSetName="SearchBar")]
        [Object]$Name="",
        
        [Parameter(Mandatory=$false,ParameterSetName="SearchBar")]
        [Switch]$Search_Self=$false,

        [Parameter(Mandatory=$false,ParameterSetName="SnapViaURL")]
        [Object]$Message="",

        [Parameter(Mandatory=$false,ParameterSetName="SearchBar")][ValidateSet("SnapViaURL","SearchBar")]
        [Parameter(Mandatory=$false,ParameterSetName="SnapViaURL")][ValidateSet("SnapViaURL","SearchBar")]
        [Object]$Access_Method="SearchBar"
    )
    Write-Debug "Get-WhatsAppContactPage; '$($Access_Method)' mode"

    if ($false -eq $Search_Self) {
        if ( ($null -eq $Phone_Number) -and ($null -eq $Name) ) {
            Write-Error "Need at least a Name or a Phone Number!"
            return
        }
    }

    switch ($Access_Method) {
        ("SnapViaURL") {
            $url_encoded_message = [System.Web.HttpUtility]::UrlEncode($Message)

            $uri = "https://web.whatsapp.com/send?phone=$([System.Web.HttpUtility]::UrlEncode($Phone_Number))" + `
                "&text=$($url_encoded_message)" + `
                "&source&data&app_absent"
            Write-Debug "Navigating to '$($uri)'..."
            $Web_Driver.Navigate().GoToUrl($uri)
            Start-Sleep -Seconds 5
        
            return $Web_Driver
        }
        ("Searchbar") {
            Move-WhatsAppPane -Web_Driver $Web_Driver -Pane pane-side -Scroll_Level 0 | Out-Null

            $search_bar = Get-WhatsAppSearchBar `
                -Web_Driver $driver
        
            $Web_Driver = Clear-WhatsAppContactSearchBar -Web_Driver $Web_Driver `
                -WhatsApp_Search_Bar_Object $search_bar
        
            $search_bar.Click()

            if ($Search_Self) {
                $search_bar.SendKeys("(You)")
            } 
            else {
                $search_bar.SendKeys($Name)
            }

            Start-Sleep -Seconds 5
            
            # This is the prefix for each class of text that corresponds to a subject's name in the search bar.
            # The owning/controlling user account has a different tag than other contacts; it's ordered like
            # this: "{prefix} {chat_suffix}"
            # $contact_label = "ggj6brxn gfz4du6o r7fjleex g0rxnol2 lhj4utae le5p0ye3 l7jjieqr _11JPr"
            # $self_label    = "tvf2evcx gfz4du6o r7fjleex g0rxnol2 lhj4utae le5p0ye3 l7jjieqr _11JPr"

            $self_prefix_message_class = "ggj6brxn"
            $default_prefix_message_class = "tvf2evcx"
            $chat_suffix_message_class = "gfz4du6o r7fjleex g0rxnol2 lhj4utae le5p0ye3 l7jjieqr _11JPr"
            $self_text_class = "seopfc61 i539y0ga _11JPr"

            if ($Search_Self) {
                $xpath = '//*[contains(@class,"{0}") and contains(@data-testid,"you-label")]' -f ($self_text_class)
            }
            else {
                $xpath = '//*[contains(@class,"{0}") and contains(@title,"{1}")]' -f ($chat_suffix_message_class, $Name)
            }

            Write-Debug "Xpath: '$($xpath)'"
            
            $user = $Web_Driver.FindElementsByXpath($xpath) `
                | Where-Object {($_.Text -like "*$($name)*" -or ($_.Text -like "*$($Phone_Number)*"))}

            if ($null -eq $user) {
                Write-Error "Could not find requested information!"
            }
            return $user
        }
    }
}

Function Get-WhatsAppMessageBar {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][Object]$Web_Driver
    )
    Write-Debug "Get-WhatsAppMessageBar"
    $message_bar_class = "conversation-compose-box-input"
    $message_bar = $driver.FindElementByXPath(
        '//*[@data-testid="{0}"]' -f ($message_bar_class)
    )
    return $message_bar
}

Function Set-WhatsAppMessageBarText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][Object]$Message,
        [Parameter(Mandatory=$true)][Object]$Web_Driver
    )
    Write-Debug "Set-WhatsAppMessageBarText"
    $Message_Bar = Get-WhatsAppMessageBar -Web_Driver $Web_Driver
    $Message_Bar.SendKeys([OpenQA.Selenium.Keys]::Control + "A")

    $split_message = $message.split("`n")

    foreach ($line in $split_message) {
        $Message_Bar.SendKeys($line)

        if ($split_message.IndexOf($line) -lt $split_message.count-1) {
            $Message_Bar.SendKeys(
                [OpenQA.Selenium.Keys]::Shift + 
                [OpenQA.Selenium.Keys]::Enter
            ) 
        }
    }
    return
}

Function Set-WhatsAppAttachment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [Object]$Web_Driver,

        [Parameter(Mandatory=$true)]
        [Object]$Path,

        [Parameter(Mandatory=$false)]
        [Object]$Message="",
        
        [Parameter(Mandatory=$false)]
        [Switch]$SendImmediately=$false
    )
    Write-Debug "Send-WhatsAppAttachment"
    
    if (Test-Path $Path) {
        $file = Get-Item $Path
        $attachment_button = $driver.FindElementByXPath(
            '//div[@title = "Attach"]'
        )
        $attachment_button.click()
    
        $image_box = $driver.FindElementByXPath(
            '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]'
        )
    
        $image_box.SendKeys($file.FullName)

        # Start-Sleep -Seconds 3

        $attachment_message_div = $Web_Driver.FindElementByXPath(
            '//*[@data-testid="media-caption-input-container"]'
        )
        $attachment_message_div.SendKeys($Message)

        if ($SendImmediately) {
            $attachment_message_div.SendKeys([OpenQA.Selenium.Keys]::Enter)
            return
        }

        return $attachment_message_div
    }
    else {
        Write-Error "Could not find item at '$($file.FullName)'!"
        return
    }
}