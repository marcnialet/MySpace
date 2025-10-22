# Function to display the main menu
function Show-MainMenu {
    cls
    Write-Host "==================== Menu ====================" -ForegroundColor Cyan
    Write-Host "1. Configure timezone"
    Write-Host "2. Quit"
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
    Read-Host "Press any key to continue..."
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

# Main menu loop
while ($true) {
    Show-MainMenu
    $selection = Read-Host "Please make a selection"

    switch ($selection) {
        '1' {
            Run-ConfigureTimezone
        }
        '2' {
            Exit
        }
        default {
            Write-Host "Invalid selection. Option '$selection' is not available." -ForegroundColor Red
            Read-Host "Press any key to continue..."
        }
    }
}