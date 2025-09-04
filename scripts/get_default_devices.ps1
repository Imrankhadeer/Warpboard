# This script gets the default audio devices for playback and recording.
# It requires the AudioDeviceCmdlets module to be present in a 'modules' subdirectory.

try {
    # Construct the path to the local module and import it.
    $ScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
    $ModulePath = Join-Path $ScriptRoot "modules/AudioDeviceCmdlets"
    Import-Module -Name $ModulePath -Force
}
catch {
    Write-Error "Could not load the required AudioDeviceCmdlets module from the local 'modules' directory."
    exit 1
}

# Get the default devices for each role.
# We get both 'Console' (speakers) and 'Communications' (headset) as they can be different.
$defaultPlayback = Get-AudioDevice -Playback -Default
$defaultCommPlayback = Get-AudioDevice -Playback -Communications
$defaultRecording = Get-AudioDevice -Recording -Default
$defaultCommRecording = Get-AudioDevice -Recording -Communications

# Output the device information as a JSON object for easy parsing in Python.
# We use the device's unique ID for reliability, as names can change or be duplicated.
$output = @{
    DefaultPlaybackId = $defaultPlayback.ID
    DefaultPlaybackName = $defaultPlayback.FriendlyName
    DefaultCommunicationsPlaybackId = $defaultCommPlayback.ID
    DefaultCommunicationsPlaybackName = $defaultCommPlayback.FriendlyName
    DefaultRecordingId = $defaultRecording.ID
    DefaultRecordingName = $defaultRecording.FriendlyName
    DefaultCommunicationsRecordingId = $defaultCommRecording.ID
    DefaultCommunicationsRecordingName = $defaultCommRecording.FriendlyName
}

# Convert the output object to a JSON string and write it to the console.
$output | ConvertTo-Json -Compress
