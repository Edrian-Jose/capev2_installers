<#
.SYNOPSIS
    Generate dummy files for malware sandbox analysis

.DESCRIPTION
    Creates realistic dummy files in various locations to make the sandbox
    environment appear more like a real user's computer. Malware often checks
    for the presence of files to detect sandbox environments.

    Files are created in:
    - C:\Dummy\ (main location - copied to Desktop during first boot)
    - Public Desktop
    - Public Documents
    - Public Downloads

.NOTES
    - Run as Administrator
    - Creates files that survive sysprep
    - Uses realistic file names and content
    - Creates various file types: Office, PDF, images, text, etc.

.EXAMPLE
    .\generate_dummy_files.ps1

.EXAMPLE
    .\generate_dummy_files.ps1 -Verbose
#>

[CmdletBinding()]
param(
    [switch]$SkipImages,
    [switch]$SkipOfficeFiles,
    [switch]$SkipPDFs
)

# Require Administrator
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script requires Administrator privileges. Please run as Administrator."
    Exit 1
}

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "Generate Dummy Files for Malware Analysis" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# Base directory for dummy files
$dummyBase = "C:\Dummy"
$publicDesktop = "C:\Users\Public\Desktop"
$publicDocuments = "C:\Users\Public\Documents"
$publicDownloads = "C:\Users\Public\Downloads"
$publicPictures = "C:\Users\Public\Pictures"

# Create directories
$directories = @($dummyBase, $publicDesktop, $publicDocuments, $publicDownloads, $publicPictures)
foreach ($dir in $directories) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        Write-Host "[+] Created directory: $dir" -ForegroundColor Green
    }
}

Write-Host ""
Write-Host "Generating dummy files..." -ForegroundColor Yellow
Write-Host ""

$filesCreated = 0

#region Helper Functions

function New-DummyTextFile {
    param(
        [string]$Path,
        [string]$Content,
        [string]$Description
    )

    try {
        $Content | Out-File -FilePath $Path -Encoding UTF8 -Force
        Write-Host "[+] Created: $Description" -ForegroundColor Green
        Write-Host "    $Path" -ForegroundColor DarkGray
        return $true
    } catch {
        Write-Host "[-] Failed to create $Description : $_" -ForegroundColor Red
        return $false
    }
}

function New-DummyBinaryFile {
    param(
        [string]$Path,
        [int]$SizeKB = 100,
        [string]$Description
    )

    try {
        # Create random binary data
        $bytes = New-Object byte[] ($SizeKB * 1024)
        $random = New-Object System.Random
        $random.NextBytes($bytes)
        [System.IO.File]::WriteAllBytes($Path, $bytes)

        Write-Host "[+] Created: $Description" -ForegroundColor Green
        Write-Host "    $Path" -ForegroundColor DarkGray
        return $true
    } catch {
        Write-Host "[-] Failed to create $Description : $_" -ForegroundColor Red
        return $false
    }
}

#endregion

#region Text Documents

Write-Host "Creating text documents..." -ForegroundColor Cyan

$textFiles = @{
    "README.txt" = @"
Project Notes
=============

This folder contains various project files and documents.
Last updated: $(Get-Date -Format "MMMM dd, yyyy")

Important Files:
- Financial_Report_2024.xlsx
- Client_Database.xlsx
- Passwords.txt (encrypted)
- Project_Proposal.docx

TODO:
[ ] Review quarterly reports
[ ] Update client database
[ ] Backup important files
[ ] Send invoices
"@

    "Passwords.txt" = @"
Personal Account Information
=============================
DO NOT SHARE THIS FILE

Email Accounts:
john.doe@company.com - Password: P@ssw0rd123!
personal@gmail.com - Password: MySecurePass2024

Banking:
Chase Bank - Account: ****5678 - PIN: 1234
Wells Fargo - Account: ****9012 - PIN: 5678

Social Media:
Facebook: johndoe@email.com / FbPass123
LinkedIn: john.doe@company.com / LinkedPass456
Twitter: @johndoe / TwPass789

Shopping:
Amazon: john.doe@email.com / AmazonPass2024
eBay: johndoe123 / eBayPass456

VPN:
NordVPN Username: johndoe_vpn
NordVPN Password: VPNSecure2024!

WiFi:
Home Network: MyHomeWiFi
Password: WiFiPass123456

Recovery Questions:
Mother's maiden name: Smith
First pet: Fluffy
High school: Lincoln High
"@

    "Meeting_Notes.txt" = @"
Team Meeting - $(Get-Date -Format "MMMM dd, yyyy")
================================================

Attendees:
- John Doe (Project Manager)
- Jane Smith (Developer)
- Bob Johnson (QA)
- Alice Williams (Designer)

Agenda:
1. Project status update
2. Sprint planning
3. Budget review
4. Client feedback

Action Items:
- John: Update project timeline
- Jane: Complete feature development
- Bob: Prepare test cases
- Alice: Finalize UI mockups

Next Meeting: Next Monday 2:00 PM
"@

    "Shopping_List.txt" = @"
Shopping List
=============

Groceries:
- Milk
- Bread
- Eggs
- Cheese
- Apples
- Bananas
- Chicken
- Rice
- Pasta
- Tomato sauce

Household:
- Laundry detergent
- Dish soap
- Paper towels
- Toilet paper
- Light bulbs

Electronics:
- USB cables
- Mouse batteries
- Phone charger
"@

    "TODO.txt" = @"
Personal TODO List
==================

Work:
[ ] Complete project proposal
[ ] Send invoice to client
[ ] Update portfolio website
[ ] Backup work files
[ ] Review code changes

Personal:
[ ] Pay electricity bill
[ ] Schedule dentist appointment
[ ] Book flight tickets
[ ] Renew car insurance
[ ] Call mom

Home:
[ ] Fix leaking faucet
[ ] Clean garage
[ ] Organize closet
[ ] Water plants
"@

    "Contacts.txt" = @"
Important Contacts
==================

Work:
Boss: John Manager - john.manager@company.com - (555) 123-4567
HR: hr@company.com - (555) 234-5678
IT Support: support@company.com - (555) 345-6789

Family:
Mom: (555) 111-2222
Dad: (555) 111-3333
Brother: (555) 111-4444

Friends:
Mike: (555) 222-3333
Sarah: (555) 222-4444
Tom: (555) 222-5555

Services:
Dentist: Dr. Smith - (555) 333-4444
Doctor: Dr. Jones - (555) 333-5555
Plumber: Bob's Plumbing - (555) 444-5555
Electrician: Sparky Electric - (555) 444-6666
"@
}

foreach ($file in $textFiles.GetEnumerator()) {
    $path = Join-Path $dummyBase $file.Key
    if (New-DummyTextFile -Path $path -Content $file.Value -Description $file.Key) {
        $filesCreated++
    }
}

#endregion

#region Office Documents (Dummy)

if (-not $SkipOfficeFiles) {
    Write-Host ""
    Write-Host "Creating dummy Office documents..." -ForegroundColor Cyan

    # Word documents
    $wordFiles = @(
        "Resume_John_Doe.docx",
        "Cover_Letter.docx",
        "Project_Proposal.docx",
        "Contract_Template.docx",
        "Invoice_Template.docx",
        "Meeting_Minutes.docx",
        "Annual_Report_2024.docx",
        "Employee_Handbook.docx",
        "Business_Plan.docx",
        "Marketing_Strategy.docx"
    )

    # Excel documents
    $excelFiles = @(
        "Budget_2024.xlsx",
        "Financial_Report_Q1.xlsx",
        "Client_Database.xlsx",
        "Sales_Tracker.xlsx",
        "Inventory_List.xlsx",
        "Employee_Records.xlsx",
        "Expense_Report.xlsx",
        "Project_Timeline.xlsx",
        "Revenue_Analysis.xlsx",
        "Tax_Documents_2024.xlsx"
    )

    # PowerPoint documents
    $pptFiles = @(
        "Company_Presentation.pptx",
        "Sales_Pitch.pptx",
        "Quarterly_Review.pptx",
        "Product_Demo.pptx",
        "Training_Materials.pptx"
    )

    # Create dummy Office files (just binary placeholders)
    foreach ($file in $wordFiles) {
        $path = Join-Path $dummyBase $file
        if (New-DummyBinaryFile -Path $path -SizeKB 50 -Description $file) {
            $filesCreated++
        }
    }

    foreach ($file in $excelFiles) {
        $path = Join-Path $dummyBase $file
        if (New-DummyBinaryFile -Path $path -SizeKB 75 -Description $file) {
            $filesCreated++
        }
    }

    foreach ($file in $pptFiles) {
        $path = Join-Path $dummyBase $file
        if (New-DummyBinaryFile -Path $path -SizeKB 200 -Description $file) {
            $filesCreated++
        }
    }
}

#endregion

#region PDF Documents

if (-not $SkipPDFs) {
    Write-Host ""
    Write-Host "Creating dummy PDF files..." -ForegroundColor Cyan

    $pdfFiles = @(
        "User_Manual.pdf",
        "Product_Catalog.pdf",
        "Invoice_2024_001.pdf",
        "Contract_Signed.pdf",
        "Tax_Return_2023.pdf",
        "Bank_Statement_Jan2024.pdf",
        "Insurance_Policy.pdf",
        "Flight_Ticket_Confirmation.pdf",
        "Hotel_Reservation.pdf",
        "Resume.pdf"
    )

    foreach ($file in $pdfFiles) {
        $path = Join-Path $dummyBase $file
        if (New-DummyBinaryFile -Path $path -SizeKB 150 -Description $file) {
            $filesCreated++
        }
    }
}

#endregion

#region Image Files

if (-not $SkipImages) {
    Write-Host ""
    Write-Host "Creating dummy image files..." -ForegroundColor Cyan

    # Create simple 1x1 pixel images
    $imageFiles = @(
        "family_photo.jpg",
        "vacation_2024.jpg",
        "screenshot_desktop.png",
        "profile_picture.jpg",
        "id_card_scan.jpg",
        "receipt.jpg",
        "passport_scan.png",
        "signature.png",
        "logo.png",
        "diagram.png"
    )

    foreach ($file in $imageFiles) {
        $path = Join-Path $publicPictures $file
        if (New-DummyBinaryFile -Path $path -SizeKB 200 -Description $file) {
            $filesCreated++
        }
    }
}

#endregion

#region Email Files

Write-Host ""
Write-Host "Creating dummy email files..." -ForegroundColor Cyan

$emlContent = @"
From: boss@company.com
To: john.doe@company.com
Subject: Urgent: Project Deadline
Date: $(Get-Date -Format "R")

John,

Please complete the project proposal by end of day today.
The client is waiting for our response.

Regards,
Boss
"@

$emlPath = Join-Path $dummyBase "Important_Email.eml"
if (New-DummyTextFile -Path $emlPath -Content $emlContent -Description "Important_Email.eml") {
    $filesCreated++
}

#endregion

#region Browser Files

Write-Host ""
Write-Host "Creating dummy browser files..." -ForegroundColor Cyan

$bookmarksHtml = @"
<!DOCTYPE NETSCAPE-Bookmark-file-1>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=UTF-8">
<TITLE>Bookmarks</TITLE>
<H1>Bookmarks</H1>
<DL><p>
    <DT><H3>Work</H3>
    <DL><p>
        <DT><A HREF="https://mail.google.com/">Gmail</A>
        <DT><A HREF="https://calendar.google.com/">Calendar</A>
        <DT><A HREF="https://drive.google.com/">Google Drive</A>
        <DT><A HREF="https://github.com/">GitHub</A>
    </DL><p>
    <DT><H3>Banking</H3>
    <DL><p>
        <DT><A HREF="https://www.chase.com/">Chase Bank</A>
        <DT><A HREF="https://www.wellsfargo.com/">Wells Fargo</A>
        <DT><A HREF="https://www.paypal.com/">PayPal</A>
    </DL><p>
    <DT><H3>Shopping</H3>
    <DL><p>
        <DT><A HREF="https://www.amazon.com/">Amazon</A>
        <DT><A HREF="https://www.ebay.com/">eBay</A>
        <DT><A HREF="https://www.walmart.com/">Walmart</A>
    </DL><p>
</DL><p>
"@

$bookmarksPath = Join-Path $dummyBase "bookmarks.html"
if (New-DummyTextFile -Path $bookmarksPath -Content $bookmarksHtml -Description "bookmarks.html") {
    $filesCreated++
}

#endregion

#region Configuration Files

Write-Host ""
Write-Host "Creating dummy configuration files..." -ForegroundColor Cyan

$gitConfig = @"
[user]
    name = John Doe
    email = john.doe@company.com
[core]
    editor = notepad
    autocrlf = true
[credential]
    helper = wincred
"@

$gitConfigPath = Join-Path $dummyBase ".gitconfig"
if (New-DummyTextFile -Path $gitConfigPath -Content $gitConfig -Description ".gitconfig") {
    $filesCreated++
}

$sshConfig = @"
Host github.com
    User git
    IdentityFile ~/.ssh/id_rsa

Host gitlab.com
    User git
    IdentityFile ~/.ssh/id_rsa

Host company-server
    HostName 192.168.1.100
    User johndoe
    Port 22
"@

$sshConfigPath = Join-Path $dummyBase "ssh_config.txt"
if (New-DummyTextFile -Path $sshConfigPath -Content $sshConfig -Description "ssh_config.txt") {
    $filesCreated++
}

#endregion

#region Database Files

Write-Host ""
Write-Host "Creating dummy database files..." -ForegroundColor Cyan

$dbFiles = @(
    "contacts.db",
    "passwords.db",
    "history.db",
    "cache.db"
)

foreach ($file in $dbFiles) {
    $path = Join-Path $dummyBase $file
    if (New-DummyBinaryFile -Path $path -SizeKB 500 -Description $file) {
        $filesCreated++
    }
}

#endregion

#region Cryptocurrency Files

Write-Host ""
Write-Host "Creating dummy cryptocurrency files..." -ForegroundColor Cyan

$walletDat = Join-Path $dummyBase "wallet.dat"
if (New-DummyBinaryFile -Path $walletDat -SizeKB 100 -Description "wallet.dat (Bitcoin)") {
    $filesCreated++
}

$cryptoKeys = @"
Bitcoin Wallet Seed Phrase (BACKUP)
====================================
WARNING: Keep this file secure!

Seed Phrase (12 words):
abandon ability able about above absent absorb abstract absurd abuse access accident

Private Keys:
Main Wallet: 5Kb8kLf9zgWQnogidDA76MzPL6TsZZY36hWXMssSzNydYXYB9KF
Trading Wallet: 5HpHagT65TZzG1PH3CSu63k8DbpvD8s5ip4nEB3kEsreAnchuDf

Public Addresses:
BTC: 1A1zP1eP5QGefi2DMPTfTL5SLmv7DivfNa
ETH: 0x742d35Cc6634C0532925a3b844Bc9e7595f0bEb
LTC: LhK2kQwiaKMKzcCpVNqj4KUwB5C4y8J

Exchange Accounts:
Coinbase: crypto@email.com / CoinbasePass123
Binance: crypto@email.com / BinancePass456
Kraken: crypto@email.com / KrakenPass789
"@

$keysPath = Join-Path $dummyBase "crypto_keys.txt"
if (New-DummyTextFile -Path $keysPath -Content $cryptoKeys -Description "crypto_keys.txt") {
    $filesCreated++
}

#endregion

#region Desktop Shortcuts

Write-Host ""
Write-Host "Creating desktop shortcuts..." -ForegroundColor Cyan

$shortcutFiles = @{
    "Work Documents" = "C:\Users\Public\Documents"
    "My Projects" = "C:\Users\Public\Documents"
    "Important Files" = "C:\Dummy"
    "Financial Records" = "C:\Dummy"
}

$wshell = New-Object -ComObject WScript.Shell

foreach ($shortcut in $shortcutFiles.GetEnumerator()) {
    try {
        $shortcutPath = Join-Path $publicDesktop "$($shortcut.Key).lnk"
        $s = $wshell.CreateShortcut($shortcutPath)
        $s.TargetPath = $shortcut.Value
        $s.Save()
        Write-Host "[+] Created shortcut: $($shortcut.Key).lnk" -ForegroundColor Green
        $filesCreated++
    } catch {
        Write-Host "[-] Failed to create shortcut: $($shortcut.Key)" -ForegroundColor Red
    }
}

#endregion

#region Recent Files History

Write-Host ""
Write-Host "Creating recent files list..." -ForegroundColor Cyan

$recentFiles = @"
Recent Files History
====================
Last 30 days file access history

Documents:
- Budget_2024.xlsx (Opened: $(Get-Date -Format "MM/dd/yyyy HH:mm"))
- Project_Proposal.docx (Opened: $(Get-Date -Format "MM/dd/yyyy HH:mm"))
- Financial_Report_Q1.xlsx (Opened: $(Get-Date -Format "MM/dd/yyyy HH:mm"))
- Resume.pdf (Opened: $(Get-Date -Format "MM/dd/yyyy HH:mm"))

Downloads:
- setup.exe (Downloaded: $(Get-Date -Format "MM/dd/yyyy HH:mm"))
- document.pdf (Downloaded: $(Get-Date -Format "MM/dd/yyyy HH:mm"))
- archive.zip (Downloaded: $(Get-Date -Format "MM/dd/yyyy HH:mm"))

Images:
- vacation_2024.jpg (Opened: $(Get-Date -Format "MM/dd/yyyy HH:mm"))
- family_photo.jpg (Opened: $(Get-Date -Format "MM/dd/yyyy HH:mm"))
- screenshot_desktop.png (Opened: $(Get-Date -Format "MM/dd/yyyy HH:mm"))
"@

$recentPath = Join-Path $dummyBase "recent_files.txt"
if (New-DummyTextFile -Path $recentPath -Content $recentFiles -Description "recent_files.txt") {
    $filesCreated++
}

#endregion

#region System Files

Write-Host ""
Write-Host "Creating dummy system files..." -ForegroundColor Cyan

$hostsFile = @"
# Hosts file
127.0.0.1 localhost
::1 localhost
192.168.1.100 company-server
192.168.1.101 dev-server
192.168.1.102 test-server
"@

$hostsPath = Join-Path $dummyBase "hosts.txt"
if (New-DummyTextFile -Path $hostsPath -Content $hostsFile -Description "hosts.txt") {
    $filesCreated++
}

#endregion

#region Create README in Dummy folder

$readmeContent = @"
Dummy Files for Malware Analysis
=================================

This folder contains dummy files created to simulate a real user environment
for malware sandbox analysis.

Contents:
- Personal documents (Word, Excel, PowerPoint)
- Financial records and invoices
- Passwords and credentials (FAKE)
- Browser bookmarks and history
- Email files
- Image files
- Configuration files
- Cryptocurrency wallet files (FAKE)
- Database files

Purpose:
These files make the sandbox environment appear more realistic to malware
that checks for specific files or user activity patterns.

IMPORTANT:
All credentials and sensitive information in these files are FAKE and
for testing purposes only. They should NEVER be used in production.

Files survive sysprep:
These files are placed in Public folders and C:\Dummy so they persist
after Windows sysprep/generalization.

Created: $(Get-Date -Format "MMMM dd, yyyy HH:mm:ss")
Script: generate_dummy_files.ps1
"@

$readmePath = Join-Path $dummyBase "README_DUMMY_FILES.txt"
New-DummyTextFile -Path $readmePath -Content $readmeContent -Description "README_DUMMY_FILES.txt" | Out-Null

#endregion

#region Create startup script to copy files to user Desktop

Write-Host ""
Write-Host "Creating startup script..." -ForegroundColor Cyan

$startupScript = @'
@echo off
REM Copy dummy files to current user Desktop on first login
REM This script runs once per user on first login

set "MARKER=%USERPROFILE%\.dummy_files_copied"

if exist "%MARKER%" (
    exit /b 0
)

echo Copying dummy files to Desktop...

REM Copy some files to user Desktop
xcopy "C:\Dummy\README.txt" "%USERPROFILE%\Desktop\" /Y /Q >nul 2>&1
xcopy "C:\Dummy\TODO.txt" "%USERPROFILE%\Desktop\" /Y /Q >nul 2>&1
xcopy "C:\Dummy\Passwords.txt" "%USERPROFILE%\Documents\" /Y /Q >nul 2>&1

REM Create marker file
echo Files copied on %date% %time% > "%MARKER%"

exit /b 0
'@

$startupPath = Join-Path $dummyBase "copy_to_desktop.bat"
New-DummyTextFile -Path $startupPath -Content $startupScript -Description "copy_to_desktop.bat (startup script)" | Out-Null

# Create startup shortcut in Public Startup folder
$startupFolder = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp"
if (Test-Path $startupFolder) {
    try {
        $shortcutPath = Join-Path $startupFolder "CopyDummyFiles.lnk"
        $s = $wshell.CreateShortcut($shortcutPath)
        $s.TargetPath = $startupPath
        $s.WindowStyle = 7  # Minimized
        $s.Save()
        Write-Host "[+] Created startup shortcut: CopyDummyFiles.lnk" -ForegroundColor Green
    } catch {
        Write-Host "[-] Failed to create startup shortcut" -ForegroundColor Red
    }
}

#endregion

#region Summary

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "Summary" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "[+] Created $filesCreated dummy files" -ForegroundColor Green
Write-Host ""
Write-Host "File locations:" -ForegroundColor Yellow
Write-Host "  - Main storage: $dummyBase" -ForegroundColor White
Write-Host "  - Public Desktop: $publicDesktop" -ForegroundColor White
Write-Host "  - Public Documents: $publicDocuments" -ForegroundColor White
Write-Host "  - Public Pictures: $publicPictures" -ForegroundColor White
Write-Host ""
Write-Host "File types created:" -ForegroundColor Yellow
Write-Host "  - Text documents (.txt)" -ForegroundColor White
Write-Host "  - Office documents (.docx, .xlsx, .pptx)" -ForegroundColor White
Write-Host "  - PDF files (.pdf)" -ForegroundColor White
Write-Host "  - Image files (.jpg, .png)" -ForegroundColor White
Write-Host "  - Email files (.eml)" -ForegroundColor White
Write-Host "  - Database files (.db)" -ForegroundColor White
Write-Host "  - Configuration files" -ForegroundColor White
Write-Host "  - Cryptocurrency files" -ForegroundColor White
Write-Host "  - Desktop shortcuts" -ForegroundColor White
Write-Host ""
Write-Host "Features:" -ForegroundColor Yellow
Write-Host "  [+] Realistic file names" -ForegroundColor White
Write-Host "  [+] Fake credentials and passwords" -ForegroundColor White
Write-Host "  [+] Browser bookmarks" -ForegroundColor White
Write-Host "  [+] Financial documents" -ForegroundColor White
Write-Host "  [+] Personal documents" -ForegroundColor White
Write-Host "  [+] Cryptocurrency wallet files" -ForegroundColor White
Write-Host "  [+] Startup script to copy files to user Desktop" -ForegroundColor White
Write-Host "  [+] All files survive sysprep" -ForegroundColor White
Write-Host ""
Write-Host "IMPORTANT:" -ForegroundColor Red
Write-Host "  All credentials in these files are FAKE!" -ForegroundColor Red
Write-Host "  For malware sandbox analysis only!" -ForegroundColor Red
Write-Host ""
Write-Host "Script completed successfully!" -ForegroundColor Green
Write-Host ""

#endregion
