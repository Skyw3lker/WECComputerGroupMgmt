# WECComputerGroupMgmt
Contain Multiple WEC related Scripts {each file has another version with no comments or documentations}


#PowerShell-WEC-InActiive-Cleaner::
This PowerShell script automates the process of checking and cleaning up outdated computers from Windows Event Collector (WEC) subscriptions. It retrieves a list of subscriptions, iterates through each, and identifies computers with a LastHeartbeatTime older than 7 days. The script then removes these outdated computers from the registry and logs their information in a subscription-specific log file located in the same directory as the script.


#PowerShell-WEC-IP-Resolver::
The script will generate a CSV file named IPz.csv in the same directory as the script. This file contains the hostnames and IP addresses of the computers in WEC subscriptions, with no duplicate entries.
