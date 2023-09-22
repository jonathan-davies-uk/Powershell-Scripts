# Define the output directory and files
$outputDir = "C:\temp\"
$systemInfoFile = Join-Path $outputDir "server_system_info.csv"
$applicationLogsFile = Join-Path $outputDir "application_logs.csv"
$systemLogsFile = Join-Path $outputDir "system_logs.csv"
$lyncLogsFile = Join-Path $outputDir "lync_logs.csv"

# Ensure directory exists
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir
}

# Get basic system info
$sysInfo = Get-WmiObject -Class Win32_ComputerSystem
$osInfo = Get-WmiObject -Class Win32_OperatingSystem
$cpuInfo = Get-WmiObject -Class Win32_Processor
$diskInfo = Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType=3"

# Get event logs
$applicationLogs = Get-EventLog -LogName Application -Newest 100 | Select-Object TimeGenerated, EntryType, Source, EventID, Message
$systemLogs = Get-EventLog -LogName System -Newest 100 | Select-Object TimeGenerated, EntryType, Source, EventID, Message
$lyncLogs = Get-EventLog -LogName "Lync Server" -ErrorAction SilentlyContinue -Newest 100 | Select-Object TimeGenerated, EntryType, Source, EventID, Message

# Export event logs to separate CSVs
$applicationLogs | Export-Csv $applicationLogsFile -NoTypeInformation
$systemLogs | Export-Csv $systemLogsFile -NoTypeInformation
$lyncLogs | Export-Csv $lyncLogsFile -NoTypeInformation

# Get installed applications and updates
$installedApps = Get-WmiObject -Class Win32_Product
$installedUpdates = Get-HotFix

# Get network info
$networkInfo = Get-WmiObject -Class Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true }

# Export main system info to CSV
$dataCollection = [PSCustomObject]@{
    "Computer Name" = $sysInfo.Name
    "Operating System" = $osInfo.Caption
    "CPU Model" = $cpuInfo.Name
    "Socket Count" = $cpuInfo.SocketDesignation
    "Processor Cores" = $cpuInfo.NumberOfCores
    "Total Memory (GB)" = [math]::Round($sysInfo.TotalPhysicalMemory / 1GB, 2)
    "Disks" = ($diskInfo | ForEach-Object { "$($_.DeviceID) Total: $([math]::Round($_.Size / 1GB, 2)) GB, Free: $([math]::Round($_.FreeSpace / 1GB, 2)) GB" }) -join "; "
    "Installed Applications" = ($installedApps | ForEach-Object { $_.Name }) -join "; "
    "Installed Updates" = ($installedUpdates | ForEach-Object { $_.Description + ": " + $_.InstalledOn }) -join "; "
    "Network Info" = ($networkInfo | ForEach-Object { "NIC: $($_.Description), IP: $($_.IPAddress -join ', '), DNS: $($_.DNSServerSearchOrder -join ', '), Gateway: $($_.DefaultIPGateway -join ', ')" }) -join "; "
}

$dataCollection | Export-Csv $systemInfoFile -NoTypeInformation

Write-Host "Inventory exported to directory: $outputDir"
