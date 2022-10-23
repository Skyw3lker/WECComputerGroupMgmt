param(
     [Parameter()]
     [Int64]$size = '419430400'
 )

$ProgressPreference = 'SilentlyContinue' 
$invocation = $MyInvocation.MyCommand.Path
$startdir = Split-Path -Parent $MyInvocation.MyCommand.Path
cd $Env:WinDir

winrm qc -q
Set-Service -Name WINRM -StartupType Automatic

Start-Sleep -Seconds 1

wecutil qc /q
Start-Sleep -Seconds 2
wevtutil sl ForwardedEvents /ms:$size
#Set-Service -Name Wecsvc -StartupType Automatic
Start-Sleep -Seconds 2

netsh http delete urlacl url=http://+:5985/wsman/
netsh http add urlacl url=http://+:5985/wsman/ "sddl=D:(A;;GX;;;S-1-5-80-569256582-2953403351-2909559716-1301513147-412116970)(A;;GX;;;S-1-5-80-4059739203-877974739-1245631912-527174227-2996563517)"

netsh http delete urlacl url=https://+:5986/wsman/
netsh http add urlacl url=https://+:5986/wsman/ "sddl=D:(A;;GX;;;S-1-5-80-569256582-2953403351-2909559716-1301513147-412116970)(A;;GX;;;S-1-5-80-4059739203-877974739-1245631912-527174227-2996563517)"

cd \tmp\wec-sub
foreach ($file in (Get-ChildItem *.xml)) {wecutil cs $file}

cd $startdir

# Inform of required reboot
write-host "Restart the computer to complete the installation"

