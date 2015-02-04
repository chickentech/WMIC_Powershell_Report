###	Machine Report Generator
### Created by Nathan Behe for the Office of Administration Help Desk ###
### 2/3/2015
### natbehe@pa.gov <- E-mail for support or questions.

##Check that script is running in STA mode:
#Validate that Script is launched 

$IsSTAEnabled = $host.Runspace.ApartmentState -eq 'STA'

#Set the name of the console window

$host.UI.RawUI.WindowTitle="Report Generator"

If ($IsSTAEnabled -eq $false) { 

"Script is not running in STA mode. Switching to STA Mode..."

#Get Script path and name

$Script = $MyInvocation.MyCommand.Definition

#Launch script in a separate PowerShell process with STA enabled 

Start-Process powershell.exe -ArgumentList "-sta $Script"

Exit

}
#######
####
#### Start of script
####
#######

####
####  Function to use a Windows Form to select File Save Location
####

Function Get-SaveFile($initialDirectory)
{ 
 [Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
 Out-Null

 $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
 $SaveFileDialog.initialDirectory = $initialDirectory
 $SaveFileDialog.DefaultExt = "html"  # Default file type set to HTML
 $SaveFileDialog.filter = "HTML Files|*.html|All files (*.*)|*.*" # Show HTML Files as the default as well as all files
 $SaveFileDialog.AddExtension = true # Add the HTML extension if not specified
 $SaveFileDialog.ShowDialog() | Out-Null
 $SaveFileDialog.filename
} 
$outputFile = Get-SaveFile -initialDirectory "C:\"

# Do this from the Command Line
#$outputFile = Read-Host 'Please enter the full path and file name to save output to.'

# Ask for the hostname from the command line (change this to a Windows Form box soon)
$targetComputer = Read-Host 'Enter the target computer hostname.'

# HTML & CSS Formatting
$a = "<style>"
$a = $a + "BODY{background-color:white;}" #Background
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;text-align: center;}"
$a = $a + "TH{border-width: 1px;padding: 2px;border-style: solid;border-color: black;}"
$a = $a + "TD{border-width: 1px;padding: 2px;border-style: solid;border-color: black;}"
$a = $a + "TR:nth-child(even){background-color: #CCC;}" # Alternate row colors
$a = $a + "TR:nth-child(odd){background-color: #FFF;}" #Alternate..
$a = $a + "</style>"

# Add JQuery
$a = $a + "<script type='text/javascript' src='http://ajax.googleapis.com/ajax/libs/jquery/1.3/jquery.min.js'></script>"


###
### Meat of the script running from WMIC Commands
###
### PC Model and Serial # information
gwmi win32_computersystem -ComputerName $targetComputer | select Manufacturer, Model, Name | ConvertTo-HTML -head $a | out-file $outputFile -Append

## Serial Number
gwmi win32_bios -ComputerName $targetComputer | select SerialNumber | ConvertTo-HTML -head $a | out-file $outputFile -Append

## Disk Information
gwmi win32_logicaldisk -ComputerName $targetComputer | select DeviceID,Description,FileSystem,FreeSpace,Size,VolumeDirty,VolumeName,VolumeSerialNumber | ConvertTo-HTML -head $a | out-file $outputFile -Append

## IP Address Information
gwmi win32_networkadapterconfiguration -ComputerName $targetComputer -filter "DHCPEnabled = True" | select Description,DHCPEnabled,DHCPLeaseObtained,DHCPServer,DNSDomain,DNSHostName,MACAddress,@{Name='IpAddress';Expression={$_.IpAddress -join '; '}},@{Name='DefaultIPgateway';Expression={$_.DefaultIPgateway -join '; '}} | ConvertTo-HTML -head $a | out-file $outputFile -Append

## Windows Updates
gwmi -cl win32_reliabilityRecords -ComputerName $targetComputer -filter "sourcename = 'Microsoft-Windows-WindowsUpdateClient'" | select @{LABEL = "date";EXPRESSION = {$_.ConvertToDateTime($_.timegenerated)}},user, productname | ConvertTo-HTML -head $a | out-file $outputFile -Append

## Add/Remove Programs
gwmi win32_product -ComputerName $targetComputer | select Name,Vendor,Version | sort name | ConvertTo-HTML -head $a | out-file $outputFile -Append
