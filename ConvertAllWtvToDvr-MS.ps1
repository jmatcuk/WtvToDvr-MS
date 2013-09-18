<# ###################################################
## Configuration Variables                          ##
################################################### #>

<#
$DebugPreference = "Continue" PowerShell will show the debug message.
$DebugPreference = "SilentlyContinue" PowerShell will not show the message.
$DebugPreference = "Stop" PowerShell will show the message and then halt.
$DebugPreference = "Inquire" PowerShell will prompt the user.
#>
$DebugPreference = "SilentlyContinue"

# Define the location of Windows WTVConverter.exe.
$wTvConverter = "C:\Windows\ehome\WTVConverter.exe"

# The buffer time in minutes that you want any actions on files to be delayed
#  in order to ensure that they are currently being recorded or written to.
$writeBufferMinutes = 3

# File type extensions configuration.
$recordingsFileExtensionPattern = '*.wtv'
$convertedShowFileExtension = ' - DVRMS.dvr-ms'

# Define the location of recorded ($recordingsFileExtensionPattern) shows.
$recordedTv = "T:\Recorded TV"
# Location to move files to when not deleting them.
$deletedShows = "T:\DeletedShows"
# Move or delete the file?  True to Move to $deletedShows or False to delete.
$moveToDeletedShows = $true

# Is the script done with what it needs to do?
$workComplete = $false

<# ###################################################
## Reusable Methods                                 ##
################################################### #>

# Method to generate the conversion output filename.
Function GetOutputFileName ($file) { ([System.IO.FileInfo] $file).DirectoryName + '\' + $recording.BaseName + $convertedShowFileExtension }

<# ###################################################
## Description                                      ##
######################################################
# This script will loop through all discovered *.wtv #
# in a location where recorded tv is stored.  Any    #
# file that has not been touched in the last buffer  #
# minutes will converted to .dvr-ms if a matching    #
# .dvr-ms file is not found.  After the file has     #
# completed being written to for buffer minutes, it  #
# will be deleted from the system.                   #
################################################### #>

# Retrieve all of the recorded ($recordingsFileExtensionPattern) shows that have not already been converted or touched within the last $writeBufferMinutes.
$showsToConvert = [System.Collections.ArrayList] (Get-ChildItem ($recordedTv + '\' + $recordingsFileExtensionPattern) |
    Where-Object {$_.LastWriteTime -lt (get-date).AddMinutes(-$writeBufferMinutes)} |
    Where-Object {!(Test-Path ($_.DirectoryName + '\' + $_.BaseName + $convertedShowFileExtension))}) # Converted by Right-Click Convert to .dvrms

# Exit if no shows are ready for conversion.
if ($showsToConvert.Count -eq 0) {Write-Debug ("No recordings found."); exit 0}

# At least one show is ready for conversion.
Write-Debug ("Show count: " + $showsToConvert.Count) # Debug

$showsToConvert | ForEach-Object {
    # Casting variable type.
    $recording = [System.IO.FileInfo] $_

    # create output file name.
    $outputFileName = GetOutputFileName $recording
    Write-Debug ("Converting " + $recording.Name + " to " + $recording.BaseName + $convertedShowFileExtension)

    # Start conversion.
    &$wTvConverter $recording.FullName $outputFileName
}

# The executible from Microsoft does not operate synchonously or output an error code
#  so pause 10 seconds to see if a file is created.
Start-Sleep 10 #seconds

$copyProtectedShows = New-Object System.Collections.ArrayList
# Loops through shows and remove 
$showsToConvert | ForEach-Object {
    # Casting variable type.
    $recording = [System.IO.FileInfo] $_

    # create output file name.
    $outputFileName = $recording.DirectoryName + '\' + $recording.BaseName + $convertedShowFileExtension
    
    # Store files that did not create output files.
    if (!(Test-Path $outputFileName)) {Write-Debug ($recording.Name + " must be copy protected."); $null = $copyProtectedShows.Add($recording)}
}
$copyProtectedShows | ForEach-Object {$showsToConvert.Remove($_)} # Remove copy-protected shows.

# Exit if all shows are copy protected.
if ($showsToConvert.Count -eq 0) {Write-Debug ("All shows found were copy protected."); exit 0}

while (!$workComplete) {
    # Check every minute if each conversion has completed by at least #writeBufferMinutes.
    Start-Sleep 60 #seconds

    $finishedShows = New-Object System.Collections.ArrayList
    $showsToConvert | ForEach-Object {
        # Casting variable type.
        $recording = [System.IO.FileInfo] $_
        
        # Create output file name.
        $outputFileName = GetOutputFileName $recording

        # Since the existing variable doesn't contain the latest time, we need to re-query for it.
        $outputFile = Get-ChildItem $outputFileName
        Write-Debug ("OutputFile: " + $outputFile.Name)
        Write-Debug ("Last Write Time: " + $outputFile.LastWriteTime)
        Write-Debug ("Now minus " + $writeBufferMinutes + " minutes: " + (get-date).AddMinutes(-$writeBufferMinutes))
        if ($outputFile.LastWriteTime -le (get-date).AddMinutes(-$writeBufferMinutes)) {
            Write-Debug ("Remove Item: " + $recording.FullName)

            # Remove the original .wtv file once it hasn't been written to in the last $writeBufferMinutes.
            if ($moveToDeletedShows) {
                Move-Item $recording.FullName $deletedShows
            } else {
                Remove-Item $recording.FullName
            }

            # Add show that will be removed to queue of finished shows.
            $null = $finishedShows.Add($recording)
        }
    }
    # Remove finished shows.
    $finishedShows | ForEach-Object {$showsToConvert.Remove($_)}

    if ($showsToConvert.Count -le 0) {
        $workComplete = $true;
    }
    Write-Debug ("Show count: " + $showsToConvert.Count) # Debug
}