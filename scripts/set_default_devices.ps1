# This script sets the default audio devices based on their unique IDs.
# It requires the AudioDeviceCmdlets module to be present in a 'modules' subdirectory.

param (
    [string]$DefaultPlayback,
    [string]$DefaultCommunicationsPlayback,
    [string]$DefaultRecording,
    [string]$DefaultCommunicationsRecording
)

try {
    # Construct the path to the local module and import it.
    $ScriptRootPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
    $ModulePath = Join-Path $ScriptRootPath "modules/AudioDeviceCmdlets"
    Import-Module -Name $ModulePath -Force
}
catch {
    Write-Error "Could not load the required AudioDeviceCmdlets module from the local 'modules' directory."
    exit 1
}

# Set the default devices using the provided IDs.
# The -Force flag is used to ensure the change is applied.
if ($DefaultPlayback) {
    Get-AudioDevice -ID $DefaultPlayback | Set-AudioDevice -Playback -Force
}
if ($DefaultCommunicationsPlayback) {
    Get-AudioDevice -ID $DefaultCommunicationsPlayback | Set-AudioDevice -Playback -Communications -Force
}
if ($DefaultRecording) {
    Get-AudioDevice -ID $DefaultRecording | Set-AudioDevice -Recording -Force
}
if ($DefaultCommunicationsRecording) {
    Get-AudioDevice -ID $DefaultCommunicationsRecording | Set-AudioDevice -Recording -Communications -Force
}

Write-Host "Default audio devices have been restored."
