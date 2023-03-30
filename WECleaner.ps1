# Create a temporary file to store the output of the "wecutil es" command
$tempFile = [System.IO.Path]::GetTempFileName()

# Run the "wecutil es" command using cmd.exe and save the output to the temporary file
cmd.exe /c "wecutil es > $tempFile"

# Read the contents of the temporary file into the $subscriptions variable
$subscriptions = Get-Content -Path $tempFile

# Delete the temporary file
Remove-Item -Path $tempFile

foreach ($subscription in $subscriptions) {
    $logFile = Join-Path $PSScriptRoot "$subscription.log"

    # Run the "wecutil gr $subscription" command using cmd.exe and save the output to the temporary file
    $tempFile = [System.IO.Path]::GetTempFileName()
    cmd.exe /c "wecutil gr $subscription > $tempFile"

    # Read the contents of the temporary file into the $rawOutput variable
    $rawOutput = Get-Content -Path $tempFile -Raw

    # Delete the temporary file
    Remove-Item -Path $tempFile

    # The rest of the script remains the same
    $pattern = '(?m)^\s{2}(.+)$\n\s{3}RunTimeStatus: .+$\n\s{3}LastError: .+$\n\s{3}LastHeartbeatTime: (.+)$'

    $eventSources = [regex]::Matches($rawOutput, $pattern) | %{
        [PSCustomObject]@{
            ComputerName = $_.Groups[1].Value.Trim()
            LastHeartbeatTime = [datetime]$_.Groups[2].Value.Trim()
        }
    }

    # Filter the computers with LastHeartbeatTime older than 7 days
    $threshold = (Get-Date).AddDays(-7)
    $oldEventSources = $eventSources | Where-Object { $_.LastHeartbeatTime -lt $threshold }

    # Output the filtered computers, delete their registry entries, and write to log file
    $logContent = $oldEventSources | ForEach-Object {
        $output = "{0}: {1}" -f $_.ComputerName, $_.LastHeartbeatTime
        Write-Host $output
        
        $registryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\EventCollector\Subscriptions\$subscription\EventSources\$($_.ComputerName)"
        Remove-Item -Path $registryPath -Force -Verbose

        # Return the output string for the log file
        $output
    }

    # Write the log content to the log file (overwrite if exists)
    Set-Content -Path $logFile -Value ($logContent -join "`r`n")
}
