# Definir enums
enum InstrumentConfig {
    T711
    T711WithLabac
    TT511
}

# Global Variables
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$global:SOFTWARE_INSTALLER_PATH = Join-Path -Path $scriptPath -ChildPath "Installer\c4c_System_Software.msi"
$global:DEVBUNDLE_INSTALLER_PATH = Join-Path -Path $scriptPath -ChildPath "Installer\c4c_Development_Bundle.exe"
$global:TSN_DROP_FOLDER = "C:\ProgramData\Roche\c4c\TsnDrop"
$global:SIM_RAWDATA_FOLDER = "C:\Users\Public\Documents\Roche\c4c System Software\RawData\SimRawData"
$global:EBARCODES_SOURCE_FOLDER = Join-Path -Path $scriptPath -ChildPath "Resources\MasterDataForDeveloperTests\FullFeatureSmokeTest"
$global:RAWDATA_SOURCE_FOLDER = Join-Path -Path $scriptPath -ChildPath "Resources\SimRawData"
$global:SIMULATOR_TOOLS_PATH = "C:\Program Files\Roche\c4c\SystemSoftware\DevelopmentTools"
$global:ICSIMULATOR_FOLDER = "C:\Program Files (x86)\Roche\c4c\ICSimulator"
$global:FRONTEND_FOLDER = "C:\Program Files\Roche\c4c\SystemSoftware\Frontend"
$global:LOGFILES_FOLDER = "C:\Users\Public\Documents\Roche\c4c System Software\log"

# New global variables for executable names
$global:ICSIMULATOR_EXE = "Roche.C4C.ICSimulator.UserInterface.exe"
$global:FRONTEND_EXE = "Roche.C4C.ServiceHosting.Frontend.WPF.exe"

function Show-Menu {
    param (
        [string]$Title = 'Menu'
    )
    cls
    Write-Host "==================== $Title ====================" -ForegroundColor Cyan
    Write-Host "=== Installation ===" -ForegroundColor Cyan
    Write-Host "    1. Install all for T711"
    Write-Host "    2. Install all for TT511"
    Write-Host "    3. Install all for T711 with labac"
    Write-Host "=== Components ===" -ForegroundColor Cyan
    Write-Host "    4. Install system software"
    Write-Host "    5. Install development bundle"
    Write-Host "    6. Copy eBarcodes & RawData"
    Write-Host "    7. Configure system software for simulator"
    Write-Host "    8. Start system software"
    Write-Host "    9. Restart system software"
    Write-Host "    10. Create shortcuts"
    Write-Host "    11. Configure IC Simulator for T711"
    Write-Host "    12. Configure IC Simulator for T711 with labac"
    Write-Host "    13. Configure IC Simulator for TT511"
    Write-Host "    14. Configure timezone"
    Write-Host "Q. Quit" -ForegroundColor Cyan
}

function Write-TimestampMessage {
    param (
        [string]$message,
        [string]$color = "White"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] $message" -ForegroundColor $color
}

function Install-SystemSoftware {
    Write-TimestampMessage "Starting Install-SystemSoftware..." -color Yellow
    #Start-Process msiexec.exe -ArgumentList "/i `"$global:SOFTWARE_INSTALLER_PATH`"" -Wait
	
	# Construct the argument list for msiexec
    # Construct the argument list for msiexec
	$args = "/i `"$global:SOFTWARE_INSTALLER_PATH`"" # normal start
	# NOT WORKING $args = "/i `"$global:SOFTWARE_INSTALLER_PATH`" InstrumentType=t711_LABAC" # normal start with param	
    # NOT WORKING $args = "/i `"$global:SOFTWARE_INSTALLER_PATH`" /qn /norestart InstrumentType=t711_LABAC" # silent start with param
	
	# Start the installation process
    Start-Process -FilePath "msiexec.exe" -ArgumentList $args -Wait
	
    Write-TimestampMessage "Completed Install-SystemSoftware." -color Green
}

function Install-DevelopmentBundle {
    Write-TimestampMessage "Starting Install-DevelopmentBundle..." -color Yellow
    Start-Process -FilePath $global:DEVBUNDLE_INSTALLER_PATH -Wait
    Write-TimestampMessage "Completed Install-DevelopmentBundle." -color Green
}

function Copy-eBarcodesAndRawData {
    Write-TimestampMessage "Starting Copy-eBarcodesAndRawData..." -color Yellow
    # Copy FullFeatureSmokeTest
    $sourceFolder1 = $global:EBARCODES_SOURCE_FOLDER
    $destinationFolder1 = $global:TSN_DROP_FOLDER

    if (-Not (Test-Path -Path $destinationFolder1)) {
        Write-TimestampMessage "Creating directory: $destinationFolder1" -color Yellow
        New-Item -ItemType Directory -Path $destinationFolder1 -Force
    }
    Copy-Item -Path "$sourceFolder1\*" -Destination $destinationFolder1 -Recurse

    # Copy RawData
    $sourceFolder2 = $global:RAWDATA_SOURCE_FOLDER
    $destinationFolder2 = $global:SIM_RAWDATA_FOLDER

    if (-Not (Test-Path -Path $destinationFolder2)) {
        Write-TimestampMessage "Creating directory: $destinationFolder2" -color Yellow
        New-Item -ItemType Directory -Path $destinationFolder2 -Force
    }
    Copy-Item -Path "$sourceFolder2\*" -Destination $destinationFolder2 -Recurse
    Write-TimestampMessage "Completed Copy-eBarcodesAndRawData." -color Green
}

function Configure-SystemSoftwareForSimulator {
    Write-TimestampMessage "Starting Configure-SystemSoftwareForSimulator..." -color Yellow
    $configureCUPath = Join-Path -Path $global:SIMULATOR_TOOLS_PATH -ChildPath "ConfigureControlUnitForSimulator.bat"
    $configureSimDataPath = Join-Path -Path $global:SIMULATOR_TOOLS_PATH -ChildPath "ConfigureSimulatedRawData.bat"

    Start-Process -FilePath $configureCUPath -Wait
    Start-Process -FilePath $configureSimDataPath -Wait
    Write-TimestampMessage "Completed Configure-SystemSoftwareForSimulator." -color Green
}

function Start-ProcessIfNotRunning {
    param (
        [string]$path,
        [string]$filename
    )
    $processPath = Join-Path -Path $path -ChildPath $filename
    $processName = [System.IO.Path]::GetFileNameWithoutExtension($filename)

    Write-TimestampMessage "Starting $filename if not already running..." -color Yellow

    if (-Not (Get-Process -Name $processName -ErrorAction SilentlyContinue)) {
        Start-Process -FilePath $processPath -WorkingDirectory $path
    } else {
        Write-TimestampMessage "$filename is already running." -color Green
    }

    Write-TimestampMessage "Completed check for $filename." -color Green
}

function Start-SystemSoftware {
    Write-TimestampMessage "Starting Start-SystemSoftware..." -color Yellow
    Start-ProcessIfNotRunning -path $global:ICSIMULATOR_FOLDER -filename $global:ICSIMULATOR_EXE
    Start-ProcessIfNotRunning -path $global:FRONTEND_FOLDER -filename $global:FRONTEND_EXE
    Write-TimestampMessage "Completed Start-SystemSoftware." -color Green
}

function Restart-SystemSoftware {
    Write-TimestampMessage "Starting Restart-SystemSoftware..." -color Yellow
    $restartScriptPath = Join-Path -Path $global:SIMULATOR_TOOLS_PATH -ChildPath "RestartSystemSoftware.bat"
    Start-Process -FilePath $restartScriptPath -Wait
    Start-SystemSoftware
    Write-TimestampMessage "Completed Restart-SystemSoftware." -color Green
}

function Create-Shortcut {
    param (
        [string]$targetPath,
        [string]$shortcutName
    )

    Write-TimestampMessage "Creating shortcut for $shortcutName..." -color Yellow

    $WScriptShell = New-Object -ComObject WScript.Shell
    $desktopPath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath('Desktop'), "$shortcutName.lnk")
    $shortcut = $WScriptShell.CreateShortcut($desktopPath)
    $shortcut.TargetPath = $targetPath
    $shortcut.Save()

    Write-TimestampMessage "Created shortcut for $shortcutName." -color Green
}

function Create-Shortcuts {
    Write-TimestampMessage "Starting Create-Shortcuts..." -color Yellow
    Create-Shortcut -targetPath $global:SIM_RAWDATA_FOLDER -shortcutName "SimRawData"
    Create-Shortcut -targetPath $global:TSN_DROP_FOLDER -shortcutName "TSNDrop"
    Create-Shortcut -targetPath $global:LOGFILES_FOLDER -shortcutName "LogFiles"
    Write-TimestampMessage "Completed Create-Shortcuts." -color Green
}

function Create-DefaultConfigurationXml {
    param (
        [string]$path
    )
    
    Write-TimestampMessage "Creating default Configuration.xml at $path..." -color Yellow
    $xmlContent = @'
<?xml version="1.0" encoding="utf-8"?>
<Configurations>
  <instrumentType>high</instrumentType>
  <hasLabac>true</hasLabac>
  <hasBalcony>true</hasBalcony>
  <testsPerHour>6000</testsPerHour>
  <initializationDuration>2000</initializationDuration>
  <loadFromDatabaseOnStartup>true</loadFromDatabaseOnStartup>
  <useDatabase>true</useDatabase>
  <connectionString>Server=c4cdb_c4c;Connect Mode=Default;Persist Security Info=True</connectionString>
  <directoryPath>C:\ProgramData\Roche\c4c\TsnDrop</directoryPath>
  <rawDataPath>C:\Users\Public\Documents\Roche\c4c Instrument Control Simulator\SourceRawData</rawDataPath>
  <bufferSize>2</bufferSize>
  <windowStaysOnTop>false</windowStaysOnTop>
  <loadMasterDataFromDatabase>False</loadMasterDataFromDatabase>
  <initializationSequence>0</initializationSequence>
  <labacOperationMode>1</labacOperationMode>
  <labacStatusNormalOperation>true</labacStatusNormalOperation>
</Configurations>
'@
    $xmlContent | Out-File -FilePath $path -Encoding utf8
    Write-TimestampMessage "Default Configuration.xml created." -color Green
}

function Configure-ICSimulator {
    param (
        [InstrumentConfig]$configType
    )

    Write-TimestampMessage "Starting Configure-ICSimulator with configuration: $configType..." -color Yellow
    $configFolderPath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath('ApplicationData'), "c4c\ICSimulator")
    $configFilePath = [System.IO.Path]::Combine($configFolderPath, "Configuration.xml")

    if (-Not (Test-Path -Path $configFolderPath)) {
        Write-TimestampMessage "Creating directory: $configFolderPath" -color Yellow
        New-Item -ItemType Directory -Path $configFolderPath -Force | Out-Null
    }

    if (-Not (Test-Path -Path $configFilePath)) {
        Create-DefaultConfigurationXml -path $configFilePath
    }

    [xml]$xmlConfig = Get-Content -Path $configFilePath

    switch ($configType) {
        "T711" {
            $xmlConfig.Configurations.instrumentType = "high"
            $xmlConfig.Configurations.hasLabac = "false"
            $xmlConfig.Configurations.testsPerHour = "6000"
            $xmlConfig.Configurations.loadFromDatabaseOnStartup = "true"
            $xmlConfig.Configurations.useDatabase = "true"
        }
        "T711WithLabac" {
            $xmlConfig.Configurations.instrumentType = "high"
            $xmlConfig.Configurations.hasLabac = "true"
            $xmlConfig.Configurations.testsPerHour = "6000"
            $xmlConfig.Configurations.loadFromDatabaseOnStartup = "true"
            $xmlConfig.Configurations.useDatabase = "true"
        }
        "TT511" {
            $xmlConfig.Configurations.instrumentType = "mid"
            $xmlConfig.Configurations.hasLabac = "false"
            $xmlConfig.Configurations.testsPerHour = "6000"
            $xmlConfig.Configurations.loadFromDatabaseOnStartup = "true"
            $xmlConfig.Configurations.useDatabase = "true"
        }
    }

    # Save the changes
    $xmlConfig.Save($configFilePath)
    Write-TimestampMessage "Configuration.xml has been updated with configuration: $configType." -color Green

    Write-TimestampMessage "Completed Configure-ICSimulator." -color Green
}

function Install-All {
    param (
        [InstrumentConfig]$configType
    )

    Write-TimestampMessage "Starting Install-All with configuration: $configType..." -color Yellow
	Run-ConfigureTimezone
    Install-SystemSoftware
    Install-DevelopmentBundle
    Run-ICSimulatorAndKill
    Copy-eBarcodesAndRawData
    Configure-SystemSoftwareForSimulator
    Configure-ICSimulator -configType $configType
    Create-Shortcuts
    Start-SystemSoftware
    Write-TimestampMessage "Completed Install-All with configuration: $configType." -color Green
}

function Run-ICSimulatorAndKill {
    Write-TimestampMessage "Starting Run-ICSimulatorAndKill..." -color Yellow

    # Start the IC Simulator
    $processPath = Join-Path -Path $global:ICSIMULATOR_FOLDER -ChildPath $global:ICSIMULATOR_EXE
    $processName = [System.IO.Path]::GetFileNameWithoutExtension($global:ICSIMULATOR_EXE)

    Start-Process -FilePath $processPath -WorkingDirectory $global:ICSIMULATOR_FOLDER

    # Wait for the process to start
    Start-Sleep -Seconds 10

    # Check if the process is running
    $process = Get-Process -Name $processName -ErrorAction SilentlyContinue

    if ($process) {
        Write-TimestampMessage "$processName is running. Now stopping it..." -color Yellow
        # Kill the process
        Stop-Process -Id $process.Id -Force
        Write-TimestampMessage "$processName has been stopped." -color Green
    } else {
        Write-TimestampMessage "$processName failed to start or is not running." -color Red
    }

    Write-TimestampMessage "Completed Run-ICSimulatorAndKill." -color Green
}

# Function to get available timezones
function Get-Timezones {
    $timezonesRaw = tzutil /l
    $timezones = @()
    for ($i = 0; $i -lt $timezonesRaw.Count; $i += 2) {
        $displayName = $timezonesRaw[$i]
        $timezoneId = $timezonesRaw[$i + 1]
        if ($displayName -match '^\(UTC([^\)]+)\)\s+(.*)') {
            $timezones += [PSCustomObject]@{
                UtcOffset  = $matches[1].Trim()
                DisplayName = $displayName
                TimezoneId  = $timezoneId.Trim()
            }
        }
    }
    return $timezones
}

# Function to display the ordered timezones menu
function Show-TimezoneMenu {
    cls
    Write-Host "==================== Timezones ====================" -ForegroundColor Cyan
    $timezones = Get-Timezones
    # Sort timezones based on UTC offset
    $sortedTimezones = $timezones | Sort-Object UtcOffset

    $option = 1
    $timezoneList = @()

    foreach ($timezone in $sortedTimezones) {
        Write-Host "$option. $($timezone.DisplayName) ($($timezone.TimezoneId))"
        $timezoneList += $timezone.TimezoneId
        $option++
    }
    return $timezoneList
}

# Function to set the selected timezone
function Configure-Timezone {
    param (
        [string]$timezoneId
    )

    Write-Host "Setting timezone to $timezoneId..."
    tzutil /s "$timezoneId"
    if ($?) {
        Write-Host "Timezone set to $timezoneId successfully." -ForegroundColor Green
    } else {
        Write-Host "Failed to set timezone. Please try again." -ForegroundColor Red
    }
}

# Function to run timezone configuration
function Run-ConfigureTimezone {
    $timezoneList = Show-TimezoneMenu
    $selectedOption = Read-Host "Please select a timezone by number"

    # Validate if selection is a number
    if ([int]::TryParse($selectedOption, [ref]$selectedOption)) {
        # Validate if selection is within valid range
        if ($selectedOption -gt 0 -and $selectedOption -le $timezoneList.Count) {
            $selectedTimezoneId = $timezoneList[$selectedOption - 1]
            Write-Host "You selected: $selectedTimezoneId" -ForegroundColor Yellow

            Configure-Timezone -timezoneId $selectedTimezoneId
        } else {
            Write-Host "Invalid selection. Option $selectedOption is out of range." -ForegroundColor Red
            Read-Host "Press any key to continue..."
        }
    } else {
        Write-Host "Invalid input. '$selectedOption' is not a valid number." -ForegroundColor Red
        Read-Host "Press any key to continue..."
    }
}
# Main loop
while ($true) {
    Show-Menu
    $selection = Read-Host "Please make a selection"

    switch ($selection.ToUpper()) {
        '1' {
            Install-All -configType T711
            Read-Host "Press any key to continue..."
        }
        '2' {
            Install-All -configType TT511
            Read-Host "Press any key to continue..."
        }
        '3' {
            Install-All -configType T711WithLabac
            Read-Host "Press any key to continue..."
        }
        '4' {
            Install-SystemSoftware
            Read-Host "Press any key to continue..."
        }
        '5' {
            Install-DevelopmentBundle
            Read-Host "Press any key to continue..."
        }
        '6' {
            Copy-eBarcodesAndRawData
            Read-Host "Press any key to continue..."
        }
        '7' {
            Configure-SystemSoftwareForSimulator
            Read-Host "Press any key to continue..."
        }
        '8' {
            Start-SystemSoftware
            Read-Host "Press any key to continue..."
        }
        '9' {
            Restart-SystemSoftware
            Read-Host "Press any key to continue..."
        }
        '10' {
            Create-Shortcuts
            Read-Host "Press any key to continue..."
        }
        '11' {
            Configure-ICSimulator -configType T711
            Read-Host "Press any key to continue..."
        }
        '12' {
            Configure-ICSimulator -configType T711WithLabac
            Read-Host "Press any key to continue..."
        }
        '13' {
            Configure-ICSimulator -configType TT511
            Read-Host "Press any key to continue..."
        }
        '14' {
            Run-ConfigureTimezone
            Read-Host "Press any key to continue..."
        }
        'Q' {
            Exit
        }
        default {
            Write-TimestampMessage "Invalid selection. Please try again." -color Red
        }
    }
}