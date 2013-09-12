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

# Retrieve all of the recorded (*.wtv) shows and loop through them.
Get-ChildItem ($recordedTv + '\' + '*.wtv') | ForEach-Object {
    echo $_.FullName
    # Check if the file has not been written to in the last $writeBufferMinutes.
    if ($_.LastWriteTime -lt (get-date).AddMinutes(-$writeBufferMinutes))
    {
        $outputFileName = $_.DirectoryName + '\' + $_.BaseName + '.dvr-ms'

        # If the output file does not already exist.
        if (!(Test-Path $outputFileName))
        {
            echo ($outputFileName + " does not already exists.")

            # Start conversion.
            &$wTvConverter $_.FullName $outputFileName
            
            # The executible from Microsoft does not operate synchonously or output an error code
            #  so pause 10 seconds to see if a file is created.
            Start-Sleep 10 #seconds

            # We want to verify that the file exists because if it doesn't then it is likely copy-protected.
            if (Test-Path $outputFileName)
            {
                do
                {
                    # Check every minute if it's finished being touched.
                    Start-Sleep 60 #seconds
                    $outputFile = Get-ChildItem $outputFileName
                    echo ("OutputFile: " + $outputFile.Name)
                    echo ("Last Write Time: " + $outputFile.LastWriteTime)
                    echo ("Now minus " + $writeBufferMinutes + " minutes: " + (get-date).AddMinutes(-$writeBufferMinutes))
                }
                while($outputFile.LastWriteTime -gt (get-date).AddMinutes(-$writeBufferMinutes))
                # Remove the original .wtv file once it hasn't been written to in the last 5 minutes.
                Remove-Item $_.FullName
            }
            else
            {
                echo ($_.BaseName + " must be copy-protected.")
            }
        }
    }
}