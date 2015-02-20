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

Start-Process powershell.exe -ArgumentList "-STA .\report_commands.ps1"

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
 $SaveFileDialog.AddExtension = $true # Add the HTML extension if not specified
 $SaveFileDialog.ShowDialog() | Out-Null
 $SaveFileDialog.filename
}

## Get the user's Desktop Folder to set as initial directory
$fol = New-Object -com Shell.Application
$bfol = ($fol.namespace(0x10)).Self.Path

## Call function to get location to save HTML File - setting default to our Desktop Folder
$outputFile = Get-SaveFile -initialDirectory $bfol

# Do this from the Command Line
#$outputFile = Read-Host 'Please enter the full path and file name to save output to.'

# Ask for the hostname from the command line (change this to a Windows Form box soon)
#$targetComputer = Read-Host 'Enter the target computer hostname.'

####
####
#### Fancy way to ask for the computer name
####
####

# Load bit from Visual Basic we need
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
# Ask for input
$targetComputer = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the computer hostname or IP address.", "Computer Name/IP Address", "")

Write-Host "Setting up HTML File..."

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
Write-Host "Querying Computer..."
Write-Host "Querying Computer's Make, Model, Serial Number..."
### PC Model and Serial # information
gwmi win32_computersystem -ComputerName $targetComputer | select Manufacturer, Model, Name | ConvertTo-HTML -head $a | out-file $outputFile -Append

## Serial Number
gwmi win32_bios -ComputerName $targetComputer | select SerialNumber | ConvertTo-HTML -head $a | out-file $outputFile -Append

Write-Host "Querying Computer's Hard Disk Information..."

## Disk Information
gwmi win32_logicaldisk -ComputerName $targetComputer | select DeviceID,Description,FileSystem,FreeSpace,Size,VolumeDirty,VolumeName,VolumeSerialNumber | ConvertTo-HTML -head $a | out-file $outputFile -Append

Write-Host "Querying Computer's Network Adapter Information..."

## IP Address Information
gwmi win32_networkadapterconfiguration -ComputerName $targetComputer -filter "DHCPEnabled = True" | select Description,DHCPEnabled,DHCPLeaseObtained,DHCPServer,DNSDomain,DNSHostName,MACAddress,@{Name='IpAddress';Expression={$_.IpAddress -join '; '}},@{Name='DefaultIPgateway';Expression={$_.DefaultIPgateway -join '; '}} | ConvertTo-HTML -head $a | out-file $outputFile -Append

Write-Host "Querying Computer's Windows Update Information..."

## Windows Updates
gwmi -cl win32_reliabilityRecords -ComputerName $targetComputer -filter "sourcename = 'Microsoft-Windows-WindowsUpdateClient'" | select @{LABEL = "date";EXPRESSION = {$_.ConvertToDateTime($_.timegenerated)}},user, productname | ConvertTo-HTML -head $a | out-file $outputFile -Append

Write-Host "Querying Computer's Add/Remove Programs List..."

## Add/Remove Programs
gwmi win32_product -ComputerName $targetComputer | select Name,Vendor,Version | sort name | ConvertTo-HTML -head $a | out-file $outputFile -Append

Write-Host "Writing End of File..."

### End of File
