# Change Log

## [0.4.0.0] - 15 May 2023
### Changes
- Fixed a bug in `PowerShell-Scraping\Modules\SeleniumWhatsApp.psm1::Set-WhatsAppMessageBarText` where messages with line breaks/carriage returns would be sent automatically.
- Added `PowerShell-Scraping\Modules\SeleniumWhatsApp.psm1::Set-WhatsAppAttachment` to add attachments to the message, with an instant send or a lazy send option.

## [0.3.0.0] - 11 May 2023
### Changes
- Added `PowerShell-Scraping\Modules\SeleniumWhatsApp.psm1::Get-WhatsAppSearchBar`.
- Added `PowerShell-Scraping\Sandbox.ps1`.
- Added `PowerShell-Scraping\Modules\SeleniumWhatsApp.psm1::Get-WhatsAppChat` to provide an interface to grab a chat using 2 different methods via URL Redirect, or via the search bar.
- Added `PowerShell-Scraping\Modules\SeleniumWhatsApp.psm1::Get-WhatsAppMessageBar`.
- Added `PowerShell-Scraping\Modules\SeleniumWhatsApp.psm1::Set-WhatsAppMessageBarText`.

## [0.2.0.0] - 10 May 2023
### Changes
- Added `ChromiumDownloader.ps1` to automate the installation of the Chrome Driver, designed to match the computer's installed copy of Google Chrome.
- Fixed a bug in `PowerShell-Scraping\Modules\SeleniumWhatsApp.psm1::Get-WhatsAppContacts` where the contact "You" would be included within the contact list.
- Added debug messages for `PowerShell-Scraping\Modules\SeleniumWhatsApp.psm1::Get-WhatsAppMessages`; added a switch parameter to allow users to decide whether conversations should be exported or not.
- Added a paragraph in README.md to include notes covering PowerShell automation of ChromeDriver downloads.
- Fixed bug in `PowerShell-Scraping\Modules\SeleniumWhatsApp.psm1::Get-WhatsAppFiles`where Documents would not download because the app would click too quickly through the tab.

## [0.1.0.0] - 09 May 2023
### Changes
- Added capability to download Doc files.
- Moved off more functionalities to `SeleniumWhatsApp.psm1`.
- Added `.gitignore`.

## [0.0.0.0] - 07 May 2023
### Changes
- Initial operating prototype for PowerShell

