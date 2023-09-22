# Define the output CSV file
$outputFile = "C:\path_to_directory\server_inventory.csv"

# Get basic system info
$sysInfo = Get-WmiObject -Class Win32_ComputerSystem
$osInfo = Get-WmiObject -Class Win32_OperatingSystem
$cpuInfo = Get-WmiObject -Class Win32_Processor
$diskInfo = Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType=3"

# Get event logs
$applicationLogs = Get-EventLog -LogName Application -Newest 100
$systemLogs = Get-EventLog -LogName System -Newest 100
$lyncLogs = Get-EventLog -LogName "Lync Server" -ErrorAction SilentlyContinue -Newest 100

# Get installed applications
$installedApps = Get-WmiObject -Class Win32_Product

# Get installed updates
$installedUpdates = Get-HotFix

# Get network info
$networkInfo = Get-WmiObject -Class Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true }

# Export to CSV
$dataCollection = [PSCustomObject]@{
    "Computer Name" = $sysInfo.Name
    "Operating System" = $osInfo.Caption
    "CPU Model" = $cpuInfo.Name
    "Socket Count" = $cpuInfo.SocketDesignation
    "Processor Count" = $sysInfo.NumberOfProcessors
    "Total Memory (GB)" = [math]::Round($sysInfo.TotalPhysicalMemory / 1GB, 2)
    "Disks" = ($diskInfo | ForEach-Object { "$($_.DeviceID) Total: $([math]::Round($_.Size / 1GB, 2)) GB, Free: $([math]::Round($_.FreeSpace / 1GB, 2)) GB" }) -join "; "
    "Last 100 Application Logs" = ($applicationLogs | ForEach-Object { $_.Message }) -join "; "
    "Last 100 System Logs" = ($systemLogs | ForEach-Object { $_.Message }) -join "; "
    "Last 100 Lync Server Logs" = if ($lyncLogs) { ($lyncLogs | ForEach-Object { $_.Message }) -join "; " } else { "N/A" }
    "Installed Applications" = ($installedApps | ForEach-Object { $_.Name }) -join "; "
    "Installed Updates" = ($installedUpdates | ForEach-Object { $_.Description + ": " + $_.InstalledOn }) -join "; "
    "Network Info" = ($networkInfo | ForEach-Object { "NIC: $($_.Description), IP: $($_.IPAddress -join ', '), DNS: $($_.DNSServerSearchOrder -join ', '), Gateway: $($_.DefaultIPGateway -join ', ')" }) -join "; "
}

$dataCollection | Export-Csv $outputFile -NoTypeInformation

Write-Host "Inventory exported to $outputFile"
