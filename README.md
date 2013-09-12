# Description
This script will loop through all discovered *.wtv in a location where recorded tv is stored.  Any file that has
not been touched in the last 5 minutes will converted to .dvr-ms if a matching .dvr-ms file is not found.  After
the file has completed being written to for 5 minutes, it will be deleted from the system.

# Tools
Microsoft Windows WTVConverter: "C:\Windows\ehome\WTVConverter.exe"
