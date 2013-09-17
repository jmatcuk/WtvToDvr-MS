<# ###################################################
## Configuration Variables                          ##
################################################### #>

# Define the location of Windows WTVConverter.exe
$wTvConverter = "C:\Windows\ehome\WTVConverter.exe"

# The buffer time in minutes that you want any actions on files to be delayed
#  in order to ensure that they are currently being recorded or written to.
$writeBufferMinutes = 3

# File type extensions configuration.
$recordingsFileExtensionPattern = '*.wtv'
$convertedShowFileExtension = ' - DVRMS.dvr-ms'

# Define the location of recorded ($recordingsFileExtensionPattern) shows.
$recordedTv = "T:\Recorded TV"

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
# file that has not been touched in the last 5       #
# minutes will converted to .dvr-ms if a matching    #
# .dvr-ms file is not found.  After the file has     #
# completed being written to for 5 minutes, it will  #
# be deleted from the system.                        #
################################################### #>

# Retrieve all of the recorded ($recordingsFileExtensionPattern) shows that have not already been converted or touched within the last $writeBufferMinutes.
$showsToConvert = [System.Collections.ArrayList] (Get-ChildItem ($recordedTv + '\' + $recordingsFileExtensionPattern) |
    Where-Object {$_.LastWriteTime -lt (get-date).AddMinutes(-$writeBufferMinutes)} |
    Where-Object {!(Test-Path ($_.DirectoryName + '\' + $_.BaseName + '.dvr-ms'))} | # Converted by WTVConverter.exe
    Where-Object {!(Test-Path ($_.DirectoryName + '\' + $_.BaseName + $convertedShowFileExtension))}) # Converted by Right-Click Convert to .dvrms

# Exit if no shows are ready for conversion.
if ($showsToConvert.Count -eq 0) {echo ("No recordings found."); exit 0}

# At least one show is ready for conversion.
("Show count: " + $showsToConvert.Count) # Debug

$showsToConvert | ForEach-Object {
    # Casting variable type.
    $recording = [System.IO.FileInfo] $_

    # create output file name.
    $outputFileName = GetOutputFileName $recording
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

# For ability to delete to recycle bin.
# http://stackoverflow.com/questions/502002/how-do-i-move-a-file-to-the-recycle-bin-using-powershell
$shell = new-object -comobject "Shell.Application"
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
        echo ("OutputFile: " + $outputFile.Name)
        echo ("Last Write Time: " + $outputFile.LastWriteTime)
        echo ("Now minus " + $writeBufferMinutes + " minutes: " + (get-date).AddMinutes(-$writeBufferMinutes))
        if ($outputFile.LastWriteTime -le (get-date).AddMinutes(-$writeBufferMinutes)) {
            # Remove the original .wtv file once it hasn't been written to in the last $writeBufferMinutes.
            #Remove-Item $_.FullName
            
            echo ("Remove Item: " + $recording.FullName)
            $item = $shell.Namespace(0).ParseName("$_.FullName")
            $item.InvokeVerb("delete")

            # Add show to remove to queue.
            $null = $finishedShows.Add($recording)
            #$showsToConvert.Remove($recording)
        }
    }
    $finishedShows | ForEach-Object {$showsToConvert.Remove($_)} # Remove finished shows.
    if ($showsToConvert.Count -le 0) {
        $workComplete = $true;
    }
    ("Show count: " + $showsToConvert.Count) # Debug
}