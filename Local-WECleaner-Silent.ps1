$tempFile = [System.IO.Path]::GetTempFileName()
cmd.exe /c "wecutil es > $tempFile"
$subscriptions = Get-Content -Path $tempFile
Remove-Item -Path $tempFile

foreach ($subscription in $subscriptions) {
    $logFile = Join-Path $PSScriptRoot "$subscription.log"
    $tempFile = [System.IO.Path]::GetTempFileName()
    cmd.exe /c "wecutil gr $subscription > $tempFile"
    $rawOutput = Get-Content -Path $tempFile -Raw
    Remove-Item -Path $tempFile

    $pattern = '(?m)^\s{2}(.+)$\n\s{3}RunTimeStatus: .+$\n\s{3}LastError: .+$\n\s{3}LastHeartbeatTime: (.+)$'
    $eventSources = [regex]::Matches($rawOutput, $pattern) | %{
        [PSCustomObject]@{
            ComputerName = $_.Groups[1].Value.Trim()
            LastHeartbeatTime = [datetime]$_.Groups[2].Value.Trim()
        }
    }

    $threshold = (Get-Date).AddDays(-7)
    $oldEventSources = $eventSources | Where-Object { $_.LastHeartbeatTime -lt $threshold }

    $logContent = $oldEventSources | ForEach-Object {
        "{0}: {1}" -f $_.ComputerName, $_.LastHeartbeatTime
        $registryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\EventCollector\Subscriptions\$subscription\EventSources\$($_.ComputerName)"
        Remove-Item -Path $registryPath -Force -ErrorAction SilentlyContinue
    }

    Set-Content -Path $logFile -Value ($logContent -join "`r`n")
}
