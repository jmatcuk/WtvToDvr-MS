# Original script copied from http://wannemacher.us/?p=412
Write-Host "Script Starting."
 
$concurrentTasks = 3 # How many tasks can run at once.
$waitTime = 500 # Duration in ms the script waits to check for job status changes.
 
# $jobs holds pertinent job information. This is a test script so these are
# just simple values to track script progress.
$jobs = @("a", "b", "c", "d", "e", "f", "g")
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
	while (($workItems.Count -lt $concurrentTasks) -and 
		   ($jobIndex -lt $jobs.Length)) {
			# These jobs don't do anything other than wait a variable amount of time
			# and print an output message.
			$workTime = Get-Random -Minimum 2000 -Maximum 10000
			$job = $jobs[$jobIndex]
			"Starting job {0}." -f $job
			$workItems[$job] = Start-Job -ArgumentList $workTime, $job -ScriptBlock {Start-Sleep -Milliseconds $args[0]; "{0} processed." -f $args[1]}
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