# Get the list of subscriptions
$subscriptions = wecutil es

# Initialize an array to store computer names
$computerNames = @()

# Process each subscription
foreach ($subscription in $subscriptions) {
    # Get subscription details
    $details = wecutil gr $subscription

    # Extract computer names
    $computerNames += ($details -split "`r`n" | Where-Object { $_ -match '^\t{2}\S+' }).Trim()
}

# Define the output file (IPz.csv)
$scriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$serversAndIps = Join-Path -Path $scriptPath -ChildPath "IPz.csv"

# Create an array to store the results
$results = @()

# Create a hash table to keep track of already processed servers
$processedServers = @{}

# Get the total number of servers to resolve
$totalServers = $computerNames.Count

# Initialize a variable to keep track of the current server being processed
$currentServer = 0

# Loop through each server in the list
foreach ($server in $computerNames) {
    # Check if the current server has already been processed
    if ($processedServers.ContainsKey($server)) {
        # If the server has already been processed, skip it and move on to the next one
        Write-Verbose "Skipping duplicate entry for $server"
        continue
    }

    # Increment the current server counter
    $currentServer++

    # Display a progress bar to indicate the progress of the script
    Write-Progress -Activity "Resolving IP addresses" -Status "Processing $server" -PercentComplete (($currentServer / $totalServers) * 100)

    # Create an object to store the server name and IP address
    $result = "" | Select ServerName, IPAddress

    # Attempt to resolve the IP address of the server
    try {
        # Get the IP addresses of the server
        $addresses = [System.Net.Dns]::GetHostAddresses($server)

        # Loop through each IP address
        foreach ($a in $addresses) {
            # Check if the IP address is an IPv4 address
            if ($a.AddressFamily -eq "InterNetwork") {
                # If the IP address is an IPv4 address, store it in the result object
                $result.IPAddress = $a.IPAddressToString
                $result.ServerName = $server
            }
        }
    } catch {
        # Write a verbose message indicating that the IP address could not be resolved
        Write-Verbose "Unable to resolve IP address for $server"
        # Store "NO IP Found" in the result object
        $result.IPAddress = "NO IP Found"
        $result.ServerName = $server
    }

    # If the IP address is not found
    if (!$result.IPAddress) {
        # Store "NO IP Found" in the result object
        $result.IPAddress = "NO IP Found"
        $result.ServerName = $server
    }

    # Add the result object to the results array
    $results += $result
    # Add the server to the hash table of processed servers
    $processedServers.Add($server, $true)
}

# Export the results to a CSV file
$results | Export-Csv -NoTypeInformation $serversAndIps
