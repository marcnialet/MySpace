# Global variables for installer paths
$global:installerPathStatusNotifier = "C:\labcoag\Status notifier_1.0.2.3.msi"
$global:installerPathCobasMobileEnabler = "C:\labcoag\cobas_t511_t711_CobasMobileEnabler.msi"
$global:installerPathRecoveryTool = "C:\labcoag\cobas_t511_t711_recoverytool_1.0.0.2.msi"

# Global variables for service names
$global:serviceNameBackendHost = "Roche.C4C.BackendHost"
$global:serviceNameServiceHost = "Roche.C4C.ServiceHost"
$global:serviceNameSystemServiceHost = "Roche.C4C.SystemServiceHost"

# Global path for executable
$global:exePathLabCoag = "C:\Program Files\Roche\c4c\SystemSoftware\Frontend\Roche.C4C.ServiceHosting.Frontend.WPF.exe"


# Function to open Date and Time settings in Control Panel
function Open-DateTimeSettings {
    Start-Process "control.exe" -ArgumentList "timedate.cpl"
}

# Function to run the installer, accepting a parameter but defaulting to the global variable
function Run-Installer {
    param (
        [string]$installerPath
    )

    if (Test-Path $installerPath) {
        Start-Process "msiexec.exe" -ArgumentList "/i $installerPath /quiet /norestart"
        Write-Host "Installer is running..."
    } else {
        Write-Host "Installer not found at $installerPath"
    }
}

# Function to start a Windows service
function Start-WindowsService {
    param (
        [string]$serviceName
    )

    $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

    if ($null -ne $service) {
        if ($service.Status -ne 'Running') {
            Start-Service -Name $serviceName
            Write-Host "Service '$serviceName' started."
        } else {
            Write-Host "Service '$serviceName' is already running."
        }
    } else {
        Write-Host "Service '$serviceName' not found."
    }
}

# Function to start an executable
function Start-Executable {
    param (
        [string]$exePath
    )

    if (Test-Path $exePath) {
        Start-Process -FilePath $exePath
        Write-Host "Executable is starting..."
    } else {
        Write-Host "Executable not found at $exePath"
    }
}
# Function to display the menu
function Show-Menu {
    Clear-Host
    Write-Host "Select option"
    Write-Host "1 - Set Windows date and timezone"
    Write-Host "2 - Install Remote notifier 1.0.2.3"
    Write-Host "3 - Install Cobas Mobile enabler 1.0"
    Write-Host "4 - Install RSR Recovery tool 1.0.0.2"
    Write-Host "5 - Run LabCoag Software"
    Write-Host "Q - Quit."
}

# Main script to handle user input and execute the appropriate actions
do {
    Show-Menu
    $option = Read-Host "Please enter your choice"

    switch ($option) {
        '1' {
            Open-DateTimeSettings
        }
        '2' {
            Run-Installer -installerPath $global:installerPathStatusNotifier
        }
        '3' {
            Run-Installer -installerPath $global:installerPathCobasMobileEnabler
        }
        '4' {
            Run-Installer -installerPath $global:installerPathRecoveryTool
        }
        '5' {
            Start-WindowsService -serviceName $global:serviceNameSystemServiceHost
            Start-WindowsService -serviceName $global:serviceNameServiceHost
            Start-WindowsService -serviceName $global:serviceNameBackendHost
			Start-Executable -exePath $global:exePathLabCoag
        }
        'q' { 
            Write-Host "Quitting the script..."
        }
        'Q' {
            Write-Host "Quitting the script..."
        }
        default {
            Write-Host "Invalid option, please try again."
        }
    }
} while ($option -ne 'q' -and $option -ne 'Q')