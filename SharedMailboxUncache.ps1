# SharedMailboxUnCache.ps1

# Define the registry path and value to set
$registryPath = "Software\Microsoft\Office\16.0\Outlook"
$cachedModeKey = "Cached Mode"
$registryValueName = "CacheOthersMail"
$registryValueData = 0 # Set the value to 0 (dword:00000000)

# Get all user profiles except the system profiles
$userProfiles = Get-WmiObject -Class Win32_UserProfile | Where-Object { !$_.Special }

foreach ($profile in $userProfiles) {
    $sid = $profile.SID
    $ntuserDatPath = $profile.LocalPath + "\NTUSER.DAT"

    # Load the user's NTUSER.DAT file into the HKEY_USERS hive
    $loadedHive = REG LOAD "HKU\$sid" $ntuserDatPath 2>&1
    if ($loadedHive -match "The operation completed successfully.") {
        # Check if the 'Cached Mode' key exists, if not, create it
        $keyPath = "Registry::HKU\$sid\$registryPath"
        $cachedModeKeyPath = "$keyPath\$cachedModeKey"
        if (-not (Test-Path $cachedModeKeyPath)) {
            New-Item -Path $keyPath -Name $cachedModeKey -Force
        }

        # Set the registry value
        try {
            Set-ItemProperty -Path $cachedModeKeyPath -Name $registryValueName -Value $registryValueData
        } catch {
            Write-Output "Failed to set registry value for SID: $sid"
        }

        # Unload the hive
        REG UNLOAD "HKU\$sid"
    } else {
        Write-Output "Failed to load hive for SID: $sid"
    }
}
