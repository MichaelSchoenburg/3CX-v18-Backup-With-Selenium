<#
    .SYNOPSIS
    Short description
    
    .DESCRIPTION
    Long description
    
    .PARAMETER Url
    Parameter description
    
    .PARAMETER UserName
    Parameter description
    
    .PARAMETER UserPassword
    Parameter description
    
    .PARAMETER FTPServerUrl
    Parameter description
    
    .PARAMETER FTPServerUserName
    Parameter description
    
    .PARAMETER FTPServerUserPassword
    Parameter description
    
    .PARAMETER DownloadPath
    Parameter description
    
    .PARAMETER SetBackupLocation
    Parameter description
    
    .PARAMETER DownloadLatestBackup
    Parameter description
    
    .PARAMETER CreateBackup
    Parameter description
    
    .EXAMPLE
    An example
    
    .NOTES
    General notes
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]
    $Url,
    [Parameter(Mandatory = $false)]
    [string]
    $UserName = 'admin',
    [Parameter(Mandatory = $true)]
    [string]
    $UserPassword,
    [Parameter(Mandatory = $false)]
    [string]
    $FTPServerUrl,
    [Parameter(Mandatory = $false)]
    [string]
    $FTPServerUserName,
    [Parameter(Mandatory = $false)]
    [string]
    $FTPServerUserPassword,
    [Parameter(Mandatory = $false)]
    [string]
    $DownloadPath,
    [switch]
    $SetBackupLocation,
    [switch]
    $DownloadLatestBackup,
    [switch]
    $CreateBackup,
    [Parameter(Mandatory = $false)]
    [string]
    $BackupName = "Initiales_Backup_nach_Backup-Ziel-Umstellung_$((Get-Date).Month)m_$((Get-Date).Day)d_$((Get-Date).Year)y_$((Get-Date).Hour)h_$((Get-Date).Minute)m",
    [switch]
    $LeaveBrowserOpen = $false
)

Import-Module Selenium

# $Driver = Start-SeFirefox
$firefoxOptions = [OpenQA.Selenium.Firefox.FirefoxOptions]::new()
$firefoxOptions.AddAdditionalCapability('acceptInsecureCerts', $true, $true)
$firefoxOptions.SetPreference('browser.download.folderList', 2)
$firefoxOptions.SetPreference('browser.download.dir', $DownloadPath)
    <# 
    This can be set to either 0, 1, or 2. 
    When set to 0, Firefox will save all files on the user’s desktop. 
    1 saves the files in the Downloads folder and 
    2 saves file at the location specified for the most recent download.
    #>
$firefoxOptions.SetPreference('browser.download.manager.showWhenStarting', $false)
    <# 
    It allows the user to specify whether or not the Download Manager window is displayed 
    when a file download is initiated.
    #>
$firefoxOptions.SetPreference('browser.helperApps.alwaysAsk.force', $false)
    <# 
    Always ask what to do with an unknown MIME type, and disable option to remember 
    what to open it with False (default): Opposite of above
    #>
$firefoxOptions.SetPreference('browser.helperApps.neverAsk.saveToDisk', 'application/x-compressed, application/x-zip-compressed, application/zip, multipart/x-zip')
    <# 
    A comma-separated list of MIME types to save to disk without asking what to use to open the file.
    Here is the path to find the MIME of the different files:
    https://www.sitepoint.com/web-foundations/mime-types-complete-list/
    #>

# Create $DownloadPath if it doesn't exist already
if (-not (Test-Path -Path $DownloadPath)) {
    Write-Verbose "The specified download path $( $DownloadPath ) doesn't exist. It will now be created..."
    New-Item -Path $DownloadPath -ItemType Directory
} else {
    Write-Verbose "The specified download path $( $DownloadPath ) does exist already."
}

# Create new driver
$Driver = New-Object -TypeName "OpenQA.Selenium.Firefox.FirefoxDriver" -ArgumentList @($firefoxOptions)

# Log in to 3CX
Write-Verbose 'Logging in to 3CX admin...'
Write-Verbose 'Logging in to 3CX admin: Opening url...'
Enter-SeUrl -Driver $Driver -Url $Url

Write-Verbose 'Logging in to 3CX admin: Setting user name...'
$Element = Find-SeElement -Driver $Driver -xpath '/html/body/div/div/div/div/form/div/div[1]/input' -Wait -Timeout 5
Send-SeKeys -Element $Element -Keys $UserName

Write-Verbose 'Logging in to 3CX admin: Setting user password...'
$Element = Find-SeElement -Driver $Driver -xpath '/html/body/div/div/div/div/form/div/div[2]/input' -Wait -Timeout 5
Send-SeKeys -Element $Element -Keys $UserPassword

Write-Verbose 'Selecting langauge...'
$Element = Find-SeElement -Driver $Driver -xpath '/html/body/div/div/div/div/form/div/div[3]/div/span/span[1]' -Wait -Timeout 5
Invoke-SeClick -Element $Element
$Element = Find-SeElement -Driver $Driver -xpath '/html/body/div/div/div/div/form/div/div[3]/div/ul/li[3]/a' -Wait -Timeout 5
Invoke-SeClick -Element $Element

Write-Verbose 'Logging in to 3CX admin: Clicking button "login"...'
$Element = Find-SeElement -Driver $Driver -xpath '/html/body/div/div/div/div/form/button' -Wait -Timeout 5
Invoke-SeClick -Element $Element

# Enter Backup and Restore
Write-Verbose 'Enter Backup and Restore...'
$Element = Find-SeElement -Driver $Driver -XPath "//*[text()='Sichern/Wiederh.']" # This button is not always at the same position, so we have to find it via text
Invoke-SeClick -Element $Element

if ($SetBackupLocation) {
    Write-Verbose 'Clicking button "location"...'
    $XPathButtonLocation = '//*[@id="btnLocation"]'
    $ButtonLocation = Find-SeElement -Driver $Driver -XPath $XPathButtonLocation -Wait -Timeout 5
    Invoke-SeClick -Element $ButtonLocation

    Write-Verbose 'Selecting "FTP" from drop down menu...'
    $XPathDropDownLocationType = '/html/body/div[1]/div/div/div/div[2]/location-settings/div[1]/select-enum-control/div/select'
    $DropDownLocationType = Find-SeElement -Driver $Driver -XPath $XPathDropDownLocationType -Wait -Timeout 5
    $DropDownLocationType.SendKeys('FTP') # Workaround to select FTP

    Write-Verbose 'Setting FTP server url...'
    $XPathInputFTPServerUrl = '/html/body/div[1]/div/div/div/div[2]/location-settings/div[4]/text-control[1]/div/input'
    $InputFTPServerUrl = Find-SeElement -Driver $Driver -XPath $XPathInputFTPServerUrl -Wait -Timeout 5
    SeType -Element $InputFTPServerUrl -Keys $FTPServerUrl -ClearFirst

    Write-Verbose 'Setting FTP user name...'
    $XPathInputFTPServerUserName = '/html/body/div[1]/div/div/div/div[2]/location-settings/div[4]/text-control[2]/div/input'
    $InputFTPServerUserName = Find-SeElement -Driver $Driver -XPath $XPathInputFTPServerUserName -Wait -Timeout 5
    do {
        SeType -Element $InputFTPServerUserName -Keys $FTPServerUserName -ClearFirst -SleepSeconds 1
    } until ($InputFTPServerUserName.GetAttribute('value') -eq $FTPServerUserName)

    Write-Verbose 'Setting FTP user password...'
    $XPathInputFTPServerUserPassword = '/html/body/div[1]/div/div/div/div[2]/location-settings/div[4]/password-control/div/div/form/input'
    $InputFTPServerUserPassword = Find-SeElement -Driver $Driver -XPath $XPathInputFTPServerUserPassword -Wait -Timeout 5
    do {
        SeType -Element $InputFTPServerUserPassword -Keys $FTPServerUserPassword -ClearFirst -SleepSeconds 1
    } until ($InputFTPServerUserPassword.GetAttribute('value') -eq $FTPServerUserPassword)

    Write-Verbose 'Pressing button "OK"...'
    $CssSelectorButtonLocationOk = '.modal-footer > button:nth-child(1)'
    $ButtonLocationOk = Find-SeElement -Driver $Driver -Wait -Timeout 5 -CssSelector $CssSelectorButtonLocationOk
    Send-SeClick -Element $ButtonLocationOk

    Start-Sleep -Seconds 1
}

if ($CreateBackup) {
    Write-Verbose 'Clicking button "Backup"...'
    $CssSelectorButtonBackup = 'html.ng-scope body div#content.h-full.ng-scope div#app.app.ng-scope.app-header-fixed div.app-content-container div#app-container.app-content-scrollable div.app-content div.app-content-body.ng-scope list-control.ng-scope div div.hbox.hbox-auto-xs.hbox-auto-sm.ng-scope div.wrapper-md.ng-scope div.panel.panel-default div.panel-body div.panel-body button#btnBackup.btn.btn-sm.btn-default.btn-responsive.ng-scope'
    $ButtonBackup = Find-SeElement -Driver $Driver -CssSelector $CssSelectorButtonBackup -Wait -Timeout 5
    Send-SeClick -Element $ButtonBackup

    Write-Verbose 'Setting backup name...'
    $XPathInputBackupName = '/html/body/div[1]/div/div/div/div[2]/form/div[1]/input'
    $InputBackupName = Find-SeElement -Driver $Driver -XPath $XPathInputBackupName
    Send-SeKeys -Element $InputBackupName -Keys $BackupName

    Write-Verbose 'Clicking checkbox "Custom Templates, Logos, Firmwares and Faxes"'
    $XPathCheckboxCustomStuff = '/html/body/div[1]/div/div/div/div[2]/form/div[3]/label/i'
    $InputCheckboxCustomStuff = Find-SeElement -Driver $Driver -XPath $XPathCheckboxCustomStuff
    Send-SeClick -Element $InputCheckboxCustomStuff

    Write-Verbose 'Clicking checkbox "Voicemails"'
    $XPathCheckboxVoicemails = '/html/body/div[1]/div/div/div/div[2]/form/div[4]/label/i'
    $InputCheckboxVoicemails = Find-SeElement -Driver $Driver -XPath $XPathCheckboxVoicemails
    Send-SeClick -Element $InputCheckboxVoicemails

    Write-Verbose 'Clicking checkbox "Recordings (Backup and Restore will take longer)"'
    $XPathCheckboxRecordings = '/html/body/div[1]/div/div/div/div[2]/form/div[5]/label/i'
    $InputCheckboxRecordings = Find-SeElement -Driver $Driver -XPath $XPathCheckboxRecordings
    Send-SeClick -Element $InputCheckboxRecordings

    Write-Verbose 'Pressing button "OK"...'
    $CssSelectorButtonBackupOk = 'html.ng-scope body.modal-open div.modal.fade.ng-scope.ng-isolate-scope.in div.modal-dialog.modal-lg div.modal-content div.modal-content.ng-scope div.modal-footer button#btnOk.btn.btn-default.ng-scope'
    $ButtonBackupOk = Find-SeElement -Driver $Driver -Wait -Timeout 5 -CssSelector $CssSelectorButtonBackupOk
    Send-SeClick -Element $ButtonBackupOk

    Write-Verbose 'Pressing button "OK"...'
    $CssSelectorButtonBackupOk2 = 'html.ng-scope body.modal-open div.modal.fade.ng-scope.ng-isolate-scope.in div.modal-dialog.modal-md div.modal-content div.modal-content.ng-scope div.modal-footer button#btnOk.btn.btn-default.ng-scope'
    $ButtonBackupOk2 = Find-SeElement -Driver $Driver -Wait -Timeout 5 -CssSelector $CssSelectorButtonBackupOk2
    Send-SeClick -Element $ButtonBackupOk2

    # Wait for backup to be finished
    $Stopwatch =  [system.diagnostics.stopwatch]::StartNew()
    $TableBackups = $null

    Write-Verbose 'Looking for the saved backup...'
    do {
        Write-Verbose '  Refreshing the site...'
        $Driver.Navigate().Refresh()

        try {
            Write-Verbose '  Looking for table of backups...'
            $TableBackups = Find-SeElement -Driver $Driver -TagName 'td' -Wait -Timeout 5
            Write-Verbose '  Found tabel of backups.'
            Write-Verbose '  Checking if backup can be found in the table...'
            try {
                if ($TableBackups.Text.Contains($BackupName)) {
                    $BackupLookUpSuccessful = $true
                    Write-Verbose '    Backup has been found.'
                } else {
                    throw # Going over to catch-sequence
                }
            }
            catch {
                Write-Verbose '    Backup could not be found.'
                Write-Verbose '    Starting five seconds cooldown...'
                Start-Sleep -Seconds 5   
            }
        }
        catch {
            Write-Verbose '  Could not find table of backups.'
        }
    } until ($BackupLookUpSuccessful)

    $Stopwatch.Stop()
    $TextTimeItTook = @()
    $TextTimeItTook += "It took $( $Stopwatch.Elapsed.Hours ) hours, "
    $TextTimeItTook += "$( $Stopwatch.Elapsed.Minutes ) minutes and "
    $TextTimeItTook += "$( $Stopwatch.Elapsed.Seconds ) seconds to find the backup."
    $TextTimeItTook = -join $TextTimeItTook
    Write-Verbose $TextTimeItTook
}

if ($DownloadLatestBackup) {
    # Find fist element in the table
    Write-Verbose 'Looking for table of backups...'
    
    # Waiting for elements to be loaded
    Start-Sleep -Seconds 1

    # Get table contents
    $Table = Get-SeElement -Driver $Driver -TagName 'td'
    $Dates = $Table.Text | Where-Object {$_ -like "*.*.*:*:*"}
        
    # Get latest date
    $DatesFinal = @()
    foreach ($date in $Dates) {
        $DatesFinal += Get-Date $date
    }
    $LastDate = $DatesFinal | Sort-Object | Select-Object -Last 1
    Write-Verbose "  Last date = $( $LastDate  )"

    Write-Verbose "  All dates:"
    foreach ($date in $DatesFinal) {
        Write-Verbose "    $( $date )"
    }

    Write-Verbose 'Downloading backup...'
    Write-Verbose '  Finding latest backup element...'
    $LatestBackup = Find-SeElement -Driver $Driver -XPath "//*[text()='$( Get-Date $LastDate -Format 'dd.MM.yyyy HH:mm:ss' )']"
    Write-Verbose '  Clicking on the backup element...'
    Invoke-SeClick -Element $LatestBackup

    Write-Verbose '  Clicking on button "download"...'
    $XPath = '//*[@id="btnDownload"]'
    $DownloadButton = Find-SeElement -Driver $Driver -XPath $XPath -Wait -Timeout 5
    Invoke-SeClick -Element $DownloadButton
}

if ($LeaveBrowserOpen) {
    Write-Verbose 'Leaving browser open.'
} else {
    Stop-SeDriver -Target $Driver
}
