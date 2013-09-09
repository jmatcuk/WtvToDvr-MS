# Define the location of Windows WTVConverter.exe
$wTvConverter = "C:\Windows\ehome\WTVConverter.exe"
# Define the location of recorded (*.wtv) shows.
$recordedTv = "T:\Recorded TV"
$writeBufferMinutes = 5

# Retrieve all of the recorded (*.wtv) shows and loop through them.
Get-ChildItem ($recordedTv + '\' + '*.wtv') | ForEach-Object {
    echo $_.FullName
    if ($_.LastWriteTime -lt (get-date).AddMinutes(-$writeBufferMinutes))
    {
        $outputFileName = $_.DirectoryName + '\' + $_.BaseName + '.dvr-ms'
        if (!(Test-Path $outputFileName))
        {
            echo ($outputFileName + " does not already exists.")
            #C:\Windows\ehome\WTVConverter.exe $_.FullName $outputFileName
            #Start-Process $wTvConverter -ArgumentList @($_.FullName, $outputFileName, "-ShowUI") -Wait
            #Start-Process $wTvConverter -Wait
            &$wTvConverter $_.FullName $outputFileName
            #echo $LASTEXITCODE
    
            Start-Sleep 10 #seconds

            if (Test-Path $outputFileName)
            {
                do
                {
                    Start-Sleep 60 #seconds
                    $outputFile = Get-ChildItem $outputFileName
                    echo ("OutputFile: " + $outputFile.Name)
                    echo ("Last Write Time: " + $outputFile.LastWriteTime)
                    echo ("Now minus " + $writeBufferMinutes + " minutes: " + (get-date).AddMinutes(-$writeBufferMinutes))
                }
                while($outputFile.LastWriteTime -gt (get-date).AddMinutes(-$writeBufferMinutes))
                Remove-Item $_.FullName
            }
            else
            {
                echo ($_.BaseName + " must be copy-protected.")
            }
        }
    }
}