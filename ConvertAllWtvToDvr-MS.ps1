######################################################
## Description                                      ##
######################################################
# This script will loop through all discovered *.wtv #
# in a location where recorded tv is stored.  Any    #
# file that has not been touched in the last 5       #
# minutes will converted to .dvr-ms if a matching    #
# .dvr-ms file is not found.  After the file has     #
# completed being written to for 5 minutes, it will  #
# be deleted from the system.                        #
######################################################

######################################################
## Configuration Variables                          ##
######################################################

# Define the location of Windows WTVConverter.exe
$wTvConverter = "C:\Windows\ehome\WTVConverter.exe"

# Define the location of recorded (*.wtv) shows.
$recordedTv = "T:\Recorded TV"

# The buffer time in minutes that you want any actions on files to be delayed
#  in order to ensure that they are currently being recorded or written to.
$writeBufferMinutes = 5
$convertedShowFileExtension = ' - DVRMS.dvr-ms'

# Retrieve all of the recorded (*.wtv) shows that have not already been converted or touched within the last $writeBufferMinutes.
$showsToConvert = [System.Collections.ArrayList] (Get-ChildItem ($recordedTv + '\' + '*.wtv') |
    Where-Object {$_.LastWriteTime -lt (get-date).AddMinutes(-$writeBufferMinutes)} |
    Where-Object {!(Test-Path ($_.DirectoryName + '\' + $_.BaseName + '.dvr-ms'))} | # Converted by WTVConverter.exe
    Where-Object {!(Test-Path ($_.DirectoryName + '\' + $_.BaseName + $convertedShowFileExtension))}) # Converted by Right-Click Convert to .dvrms

if ($showsToConvert.Count -eq 0) {echo ("No recordings found."); exit 0}
("Show count: " + $showsToConvert.Count) # Debug
$workComplete = $false # Is the script done with what it needs to do?

$showsToConvert | ForEach-Object {
    # Casting variable type.
    $recording = [System.IO.FileInfo] $_

    # create output file name.
    $outputFileName = $recording.DirectoryName + '\' + $recording.BaseName + $convertedShowFileExtension
    ("Converting " + $recording.Name + " to " + $recording.BaseName + $convertedShowFileExtension)

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
    if (!(Test-Path $outputFileName)) {($recording.Name + " must be copy protected."); $null = $copyProtectedShows.Add($recording)}
}
$copyProtectedShows | ForEach-Object {$showsToConvert.Remove($_)} # Remove copy-protected shows.

("Show count: " + $showsToConvert.Count) # Debug