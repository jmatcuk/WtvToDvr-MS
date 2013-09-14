# Original script copied from http://wannemacher.us/?p=412
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

Write-Host "Script Starting."

$recordedShows = @(Get-ChildItem ($recordedTv + '\' + '*.wtv') | Where-Object {$_.LastWriteTime -lt (get-date).AddMinutes(-$writeBufferMinutes)})

$waitTime = 10000 # Duration in ms the script waits to check for job status changes.
 
$jobIndex = 0 #Which job is up for running
$workItems = @{} # Items being worked on
$workComplete = $false # Is the script done with what it needs to do?
 
while (!$workComplete) {
	# Process any finished jobs.
	foreach ($key in @() + $workItems.Keys) {
		# Write-Host "Checking job $key."
		if ($workItems[$key].State -eq "Completed") {
			"$key is done."
			$result = Receive-Job $workItems[$key]
			$workItems.Remove($key)
			"Result: $result"
		}
	}
 
	# Start new jobs if there are open slots.
	while (($workItems.Count -lt $recordedShows.Length) -and ($jobIndex -lt $recordedShows.Length)) {
		$job = $recordedShows[$jobIndex]
        #echo ($job.GetType().ToString())

        "Starting job {0}." -f $job.Name
        $workItems[$job] = Start-Job -ArgumentList @($job) -ScriptBlock {
            $show = [System.IO.FileInfo] $args[0]
            
            #Start-Sleep -Milliseconds $args[0]; "{0} processed." -f $args[1]
            $outputFileName = $args[0].DirectoryName + '\' + $args[0].BaseName + '.dvr-ms'
            echo $outputFileName
            # If the output file does not already exist.
            if (!(Test-Path $outputFileName))
            {
                echo ($outputFileName + " does not already exists.")

                # Start conversion.
                &$wTvConverter $args[0].FullName $outputFileName
            
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
                    }
                    while($outputFile.LastWriteTime -gt (get-date).AddMinutes(-$writeBufferMinutes))
                    # Remove the original .wtv file once it hasn't been written to in the last 5 minutes.
                    Remove-Item $args[0].FullName
                }
                else
                {
                    echo ($args[0].BaseName + " must be copy-protected.")
                }
            }
        }
			


        $jobIndex += 1
	}
 
	# If all jobs have been processed we are done.
	if ($jobIndex -eq $jobs.Length -and $workItems.Count -eq 0) {
		$workComplete = $true
	}
 
	# Wait between status checks
	Start-Sleep -Milliseconds $waitTime
}
 
Write-Host "Script Finished."