
<#
.SYNOPSIS
    
    PC Hardware Inventory Script>
.DESCRIPTION
    
    This PowerShell script will collect the Date of inventory, IP and MAC address, serial number, model, CPU, RAM, total storage size, GPU(s), OS, OS build, logged in user, attached monitor(s), attached printer(s) and attached scanner(s) of a computer.
    After it collects that information, it is outputted to a CSV file. It will first check the CSV file (if it exists) to see if the hostname already exists in the file. 
    If hostname exists in the CSV file, it will overwrite it with the latest information so that the inventory is up to date and there is no duplicate information.
    It is designed to be run as a login script and/or a scheduled/immediate task run by a domain user. Elevated privileges are not required.

.PARAMETER <Parameter_Name>
   None
.INPUTS
   None
.OUTPUTS
  <Outputs if any, otherwise state None - example: Log file stored in C:\Windows\Temp\<name>.log>
.NOTES
  Version:        1.0
  Author:         Thomas Kuehn TC 26/19 thomas.kuehn@dataport.de
  Creation Date:  16.02.2021
  Purpose/Change: Initial script development
  
.EXAMPLE
  None
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------
#Serverlogpath

$pwd = "\\xxx-fs-01\inventurlog$"


#------------------------------------------------------FROM HERE NO CHANGES------------------------------------------------------------------------------------------------------------
#Region gathering PC information
$csv = "$pwd\Inventory.csv"

## Error log path (Optional but recommended. If this doesn't exist, the script will attempt to create it. Users will need full control of the file.)
$ErrorLogPath = "$pwd\PowerShell-PC-Inventory-Error-Log.log"

Write-Host "Gathering inventory information..."

# Date
$Date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# IP
$IP = (Get-CimInstance -Class Win32_NetworkAdapterConfiguration | Where-Object { $_.DefaultIPGateway -ne $null }).IPAddress | Select-Object -First 1

# MAC address
$MAC = Get-CimInstance -Class Win32_NetworkAdapterConfiguration | Where-Object { $_.DefaultIPGateway -ne $null } | Select-Object -ExpandProperty MACAddress

# Serial Number
$SN = Get-CimInstance -Class Win32_Bios | Select-Object -ExpandProperty SerialNumber

# Model
$Model = Get-CimInstance -Class Win32_ComputerSystem | Select-Object -ExpandProperty Model

#BiosVersion
$BIOS = Get-CimInstance -ClassName Win32_Bios -Property SMBIOSBIOSVersion

#Manufacturer
$Manufacturer = Get-CimInstance -ClassName Win32_BaseBoard -Property Manufacturer

#BaseBoardID
$BoardID = Get-CimInstance -ClassName Win32_BaseBoard -Property Product

# CPU
$CPU = Get-CimInstance -Class win32_processor | Select-Object -ExpandProperty Name

# RAM
$RAM = Get-CimInstance -Class Win32_PhysicalMemory | Measure-Object -Property capacity -Sum | ForEach-Object { [math]::Round(($_.sum / 1GB),2) }

# Storage
$Storage = Get-CimInstance -Class Win32_LogicalDisk -Filter "DeviceID='$env:systemdrive'" | ForEach-Object { [math]::Round($_.Size / 1GB,2) }

#GPU(s)
function GetGPUInfo {
  $GPUs = Get-CimInstance -Class Win32_VideoController
  foreach ($GPU in $GPUs) {
    $GPU | Select-Object -ExpandProperty Description
  }
}

## If some computers have more than two GPUs, you can copy the lines below, but change the variable and index number by counting them up by 1.
$GPU0 = GetGPUInfo | Select-Object -Index 0
$GPU1 = GetGPUInfo | Select-Object -Index 1

# OS
$OS = Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty Caption

# OS Build
$OSBuild = (Get-Item "HKLM:SOFTWARE\Microsoft\Windows NT\CurrentVersion").GetValue('ReleaseID')

# OS Install Date
$OSInstallDate = Get-CimInstance -ClassName Win32_OperatingSystem -Property InstallDate

#EndRegion gathering PC information

#Region gathering user information

$Username = Get-CimInstance Win32_Process -Filter "Name='explorer.exe'" | ForEach-Object { $_.GetOwner() } |Select-Object -Unique -Expand User

#EndRegion gathering user information

# Region gathering Monitor information
function GetMonitorInfo {
  
  $Monitors = Get-CimInstance -Namespace "root\WMI" -Class "WMIMonitorID"
 
  foreach ($Monitor in $Monitors) {
    ([System.Text.Encoding]::ASCII.GetString($Monitor.ManufacturerName)).Replace("$([char]0x0000)","")
    ([System.Text.Encoding]::ASCII.GetString($Monitor.UserFriendlyName)).Replace("$([char]0x0000)","")
    ([System.Text.Encoding]::ASCII.GetString($Monitor.SerialNumberID)).Replace("$([char]0x0000)","")
   ($Monitor.WeekOfManufacture)
    ($Monitor.YearOfManufacture)

  }
}

## If some computers have more than three monitors, you can copy the lines below, but change the variable and index number by counting them up by 1.
$Monitor1 = GetMonitorInfo | Select-Object -Index 0,1
$Monitor1SN = GetMonitorInfo | Select-Object -Index 2
$Monitor1WOP = GetMonitorInfo | Select-Object -Index 3
$Monitor1YOP = GetMonitorInfo | Select-Object -Index 4
$Monitor2 = GetMonitorInfo | Select-Object -Index 5,6
$Monitor2SN = GetMonitorInfo | Select-Object -Index 7
$Monitor2WOP = GetMonitorInfo | Select-Object -Index 8
$Monitor2YOP = GetMonitorInfo | Select-Object -Index 9
$Monitor3 = GetMonitorInfo | Select-Object -Index 10,11
$Monitor3SN = GetMonitorInfo | Select-Object -Index 12
$Monitor3WOP = GetMonitorInfo | Select-Object -Index 13
$Monitor3YOP = GetMonitorInfo | Select-Object -Index 14

$Monitor1 = $Monitor1 -join ' '
$Monitor2 = $Monitor2 -join ' '
$Monitor3 = $Monitor3 -join ' '
# EndRegion gathering Monitor information

#Region gathering ComputerType information 
$Chassis = Get-CimInstance -ClassName Win32_SystemEnclosure -Namespace 'root\CIMV2' -Property ChassisTypes | Select-Object -ExpandProperty ChassisTypes

if ($Chassis -eq "1") {
  $Chassis = "Other"
}
if ($Chassis -eq "2") {
  $Chassis = "Unknown"
}
if ($Chassis -eq "3") {
  $Chassis = "Desktop"
}
if ($Chassis -eq "4") {
  $Chassis = "Low Profile Desktop"
}
if ($Chassis -eq "5") {
  $Chassis = "Pizza Box"
}
if ($Chassis -eq "6") {
  $Chassis = "Mini Tower"
}
if ($Chassis -eq "7") {
  $Chassis = "Tower"
}
if ($Chassis -eq "8") {
  $Chassis = "Portable"
}
if ($Chassis -eq "9") {
  $Chassis = "Laptop"
}
if ($Chassis -eq "10") {
  $Chassis = "Notebook"
}
if ($Chassis -eq "11") {
  $Chassis = "Hand Held"
}
if ($Chassis -eq "12") {
  $Chassis = "Docking Station"
}
if ($Chassis -eq "13") {
  $Chassis = "All in One"
}
if ($Chassis -eq "14") {
  $Chassis = "Sub Notebook"
}
if ($Chassis -eq "15") {
  $Chassis = "Space-Saving"
}
if ($Chassis -eq "16") {
  $Chassis = "Lunch Box"
}
if ($Chassis -eq "17") {
  $Chassis = "Main System Chassis"
}
if ($Chassis -eq "18") {
  $Chassis = "Expansion Chassis"
}
if ($Chassis -eq "19") {
  $Chassis = "SubChassis"
}
if ($Chassis -eq "20") {
  $Chassis = "Bus Expansion Chassis"
}
if ($Chassis -eq "21") {
  $Chassis = "Peripheral Chassis"
}
if ($Chassis -eq "22") {
  $Chassis = "Storage Chassis"
}
if ($Chassis -eq "23") {
  $Chassis = "Rack Mount Chassis"
}
if ($Chassis -eq "24") {
  $Chassis = "Sealed-Case PC"
}

#EndRegion gathering ComputerType information 


#Region gathering Printer information 

$ConnectedPrinter = (Get-CimInstance -query "SELECT Name FROM Win32_Printer WHERE (PortName LIKE 'USB%') AND (WorkOffline = False)").Name

#EndRegion gathering Printer information 

#Region gathering Scanner information 

$ConnectedScanner = ((gwmi Win32_USBControllerDevice |%{[wmi]($_.Dependent)} | Where-Object {($_.Service  -match "usbscan")})).name

#EndRegion gathering Scanner information 




#Region creating .csv File 




# Function to write the inventory to the CSV file
function OutputToCSV {
  # CSV properties
  # Thanks to https://gallery.technet.microsoft.com/scriptcenter/PowerShell-Script-Get-beced710
  Write-Host "Adding inventory information to the CSV file..."
  $infoObject = New-Object PSObject
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "Date Collected" -Value $Date
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "IP Address" -Value $IP
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "Hostname" -Value $env:computername
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "MAC Address" -Value $MAC
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "User" -Value $Username
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "Type" -Value $Chassis
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "Serial Number/Service Tag" -Value $SN
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "Model" -Value $Model
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "CPU" -Value $CPU
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "RAM (GB)" -Value $RAM
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "Storage (GB)" -Value $Storage
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "GPU 0" -Value $GPU0
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "GPU 1" -Value $GPU1
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "OS" -Value $OS
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "OS Version" -Value $OSBuild
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "OS Install Date" -Value $OSInstallDate
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "SMBIOSBIOSVersion" -Value $BIOS
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "Manufacturer" -Value $Manufacturer
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "BaseBoard ID" -Value $BoardID 
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "Monitor 1" -Value $Monitor1
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "Monitor 1 Serial Number" -Value $Monitor1SN
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "Monitor 1 Year of Production"  -Value $Monitor1YOP
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "Monitor 1 Week of Production" -Value $Monitor1WOP
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "Monitor 2" -Value $Monitor2
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "Monitor 2 Year of Production"  -Value $Monitor2YOP
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "Monitor 2 Week of Production" -Value $Monitor2WOP
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "Monitor 2 Serial Number" -Value $Monitor2SN
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "Monitor 3" -Value $Monitor3
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "Monitor 3 Year of Production"  -Value $Monitor3YOP
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "Monitor 3 Week of Production" -Value $Monitor3WOP
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "Monitor 3 Serial Number" -Value $Monitor3SN
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "Connected Printer" -Value $ConnectedPrinter
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "Connected Scanner"-Value $ConnectedScanner
 $infoObject




  $infoColl += $infoObject

  # Output to CSV file
  try {
    $infoColl | Export-Csv -Path $csv -NoTypeInformation -Append
    Write-Host -ForegroundColor Green "Inventory was successfully updated!"
    # Clean up empty rows
    (Get-Content $csv) -notlike ",,,,,,,,,,,,,,,,,,,,*" | Set-Content $csv
    exit 0
  }
  catch {
    if (-not (Test-Path $ErrorLogPath))
    {
      New-Item -ItemType "file" -Path $ErrorLogPath
      icacls $ErrorLogPath /grant Everyone:F
    }
    Add-Content -Path $ErrorLogPath -Value "[$Date] $Username at $env:computername was unable to export to the inventory file at $csv."
    throw "Unable to export to the CSV file. Please check the permissions on the file."
    exit 1
  }
}
#End Region creating .csv File 
#Region Error handling 
# Just in case the inventory CSV file doesn't exist, create the file and run the inventory.
if (-not (Test-Path $csv))
{
  Write-Host "Creating CSV file..."
  try {
    New-Item -ItemType "file" -Path $csv
    icacls $csv /grant Everyone:F
    OutputToCSV
  }
  catch {
    if (-not (Test-Path $ErrorLogPath))
    {
      New-Item -ItemType "file" -Path $ErrorLogPath
      icacls $ErrorLogPath /grant Everyone:F
    }
    Add-Content -Path $ErrorLogPath -Value "[$Date] $Username at $env:computername was unable to create the inventory file at $csv."
    throw "Unable to create the CSV file. Please check the permissions on the file."
    exit 1
  }
}

# Check to see if the CSV file exists then run the script.
function Check-IfCSVExists {
  Write-Host "Checking to see if the CSV file exists..."
  $import = Import-Csv $csv
  if ($import -match $env:computername)
  {
    try {
      (Get-Content $csv) -notmatch $env:computername | Set-Content $csv
      OutputToCSV
    }
    catch {
      if (-not (Test-Path $ErrorLogPath))
      {
        New-Item -ItemType "file" -Path $ErrorLogPath
        icacls $ErrorLogPath /grant Everyone:F
      }
      Add-Content -Path $ErrorLogPath -Value "[$Date] $Username at $env:computername was unable to import and/or modify the inventory file located at $csv."
      throw "Unable to import and/or modify the CSV file. Please check the permissions on the file."
      exit 1
    }
  }
  else
  {
    OutputToCSV
  }
}

Check-IfCSVExists

