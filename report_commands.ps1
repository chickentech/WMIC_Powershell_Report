#	Machine Report Generator
### Created by Nathan Behe for the Office of Administration Help Desk ###
### 2/3/2015
### natbehe@pa.gov <- E-mail for support or questions.

##Check that script is running in STA mode:
#Validate that Script is launched 

$IsSTAEnabled = $host.Runspace.ApartmentState -eq 'STA'

#Set the name of the console window

$host.UI.RawUI.WindowTitle="Window Title"

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

Function Get-SaveFile($initialDirectory)
{ 
 [Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
 Out-Null

 $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
 $SaveFileDialog.initialDirectory = $initialDirectory
 $SaveFileDialog.filter = "All files (*.*)| *.*"
 $SaveFileDialog.ShowDialog() | Out-Null
 $SaveFileDialog.filename
} 
$outputFile = Get-SaveFile -initialDirectory "C:\"

#$outputFile = Read-Host 'Please enter the full path and file name to save output to.'

#$outputFile = Get-FileName -initialDirectory "c:\users\natbehe\desktop" 

$targetComputer = Read-Host 'Enter the target computer hostname.'
# HTML Formatting
$a = "<style>"
$a = $a + "BODY{background-color:white;}"
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;text-align: center;}"
$a = $a + "TH{border-width: 1px;padding: 2px;border-style: solid;border-color: black;}"
$a = $a + "TD{border-width: 1px;padding: 2px;border-style: solid;border-color: black;}"
$a = $a + "TR:nth-child(even){background-color: #CCC;}"
$a = $a + "TR:nth-child(odd){background-color: #FFF;}"
$a = $a + "</style>"
# Add JQuery
$a = $a + "<script type='text/javascript' src='http://ajax.googleapis.com/ajax/libs/jquery/1.3/jquery.min.js'></script>"
#Color rows even & odd different background colors
#$a = $a + "<script>"
#$a = $a + "$(function(){$('table').each(function(){$('tr:odd',this).addClass('odd').removeClass('even');$('tr:even',this).addClass('even').removeClass('odd');})});"
#$a = $a + "</script>"
#JQuery
#$h = "<script type='text/javascript' src='http://ajax.googleapis.com/ajax/libs/jquery/1.3/jquery.min.js'></script>"
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
