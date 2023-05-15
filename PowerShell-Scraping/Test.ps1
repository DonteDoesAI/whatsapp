Function Test-WhatsAppAccount {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [Object]$Web_Driver,
        
        [Parameter(Mandatory=$true)]
        [Object]$Link_Element
    )
    Write-Debug "Test-WhatsAppAccount"

    $Link_Element.click()    

    $xpath = '//*[@aria-label="Chat with "]'
    
    try {
        $element = $Web_Driver.FindElementsByXPath(
            $xpath
        )
        if ("" -eq $element) {
            $Link_Element.click()
            return $false
        }

        $Link_Element.click()
        return $true
    }
    catch {
        Write-Debug "Failure!"
        $Link_Element.click()
        return $false
    }
}

$Links = Get-WhatsAppMessageHyperLinks `
    -Web_Driver $driver

foreach ($link in $links) {
    Test-WhatsAppAccount `
        -Web_Driver $driver `
        -Link_Element $link
}

