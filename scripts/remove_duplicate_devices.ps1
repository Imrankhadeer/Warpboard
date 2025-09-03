# This script finds and disables duplicate or "ghost" VB-CABLE audio devices.
# It identifies all devices with the name "CABLE Input (VB-Audio Virtual Cable)",
# keeps the single best-functioning one, and disables the rest.

# The target device description to look for.
$deviceName = "CABLE Input (VB-Audio Virtual Cable)"

# Get all audio devices that match the name.
# We look for devices that are present on the system.
$allVBCableDevices = Get-PnpDevice -Class 'MEDIA' -PresentOnly | Where-Object { $_.FriendlyName -eq $deviceName }

if ($null -eq $allVBCableDevices -or $allVBCableDevices.Count -le 1) {
    Write-Host "No duplicate VB-CABLE devices found. No action taken."
    exit 0
}

Write-Host "Found $($allVBCableDevices.Count) devices matching the name '$deviceName'. Analyzing..."

# Try to find the single "best" device to keep. We prioritize one that is running correctly.
# The 'OK' status is the most reliable indicator.
$deviceToKeep = $null
foreach ($device in $allVBCableDevices) {
    if ($device.Status -eq 'OK') {
        $deviceToKeep = $device
        break
    }
}

# If no device has an 'OK' status, we can't safely determine which one to keep.
if ($null -eq $deviceToKeep) {
    Write-Error "Found multiple VB-CABLE devices, but none have a status of 'OK'. Cannot safely determine which device to keep. Please fix the driver installation manually."
    exit 1
}

Write-Host "Keeping device with InstanceId: $($deviceToKeep.InstanceId)"

# Now, disable all other devices that match the name but not the instance ID of the one we're keeping.
$devicesToDisable = $allVBCableDevices | Where-Object { $_.InstanceId -ne $deviceToKeep.InstanceId }

foreach ($device in $devicesToDisable) {
    try {
        Write-Host "Disabling duplicate device with InstanceId: $($device.InstanceId)"
        Disable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false
    }
    catch {
        Write-Warning "Failed to disable device with InstanceId: $($device.InstanceId). Error: $_"
    }
}

Write-Host "Duplicate device cleanup complete."
