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

$sw = [Diagnostics.Stopwatch]::StartNew() # Start Timer


Write-Host "Setting up HTML File..."

# HTML & CSS Formatting
$a = "<style>"
$a = $a + "BODY{background-color:white;}" #Background
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;text-align: center;margin-left: auto; margin-right: auto;}"
$a = $a + "TH{border-width: 1px;padding: 2px;border-style: solid;border-color: black;}"
$a = $a + "TD{border-width: 1px;padding: 2px;border-style: solid;border-color: black;}"
$a = $a + "TR:nth-child(even){background-color: #CCC;}" # Alternate row colors
$a = $a + "TR:nth-child(odd){background-color: #FFF;}" #Alternate..
$a = $a + "</style>"

# Add JQuery
$a = $a + "<script type='text/javascript' src='http://ajax.googleapis.com/ajax/libs/jquery/1.3/jquery.min.js'></script>"

#Create our output file and write the head info
#This kills any existing file with the same name
#to prevent writing to the end of an existing HTML file (bad!)

## Get Date Object for Header
$dat = Get-Date

Out-File -filePath $outputFile -InputObject $a

## Output Header on the page - including computer name and date information
## This will keep us from working from old information later.

##Generate HTML to inject into page
$e = "<div style='text-align: center;'><h1> Computer:  " + $targetComputer + "</h1>"
$e = $e + "<p><h3>Report Generated:</h3>Date: " + ($dat.ToShortDateString()) + "<br>Time:  " + ($dat.ToShortTimeString()) + "</p></div>"
##Inject Generated HTML into page
Out-File -filePath $outputFile -InputObject $e -Append

###
### Meat of the script running from WMIC Commands
###
Write-Host "Querying Computer..."
Write-Host "Querying Computer's Make, Model, Serial Number..."

############# ~`~`~`~`~`~`~`~`~`~`~ ################
#New section to combine several WMI calls into a single object and write output to the file.
#This is to make it a single table instead of several tables.
############# ~`~`~`~`~`~`~`~`~`~`~ ################


### PC Model and Serial # information
$t1 = gwmi win32_computersystem -ComputerName $targetComputer

## OS Architecture x86 or x64
$t2 = gwmi win32_operatingsystem -ComputerName $targetComputer

## Serial Number
$t3 = gwmi win32_bios -ComputerName $targetComputer

$output2 = New-Object PSObject -Property @{
  Manufacturer = $t1.Manufacturer
  Model = $t1.Model
  'Hostname' = $t1.Name
  'Logged In User' = $t1.UserName
  'OS Architecture' = $t2.OSArchitecture
  'Computer Serial Number' = $t3.SerialNumber
  'System Up Since' = ([Management.ManagementDateTimeConverter]::ToDateTime($t2.LastBootUpTime))
}

ConvertTo-Html -Fragment -inputObject $output2 | Out-File $outputFile -Append

############# ~`~`~`~`~`~`~`~`~`~`~ ################
#############  End of this section  ################
############# ~`~`~`~`~`~`~`~`~`~`~ ################

Write-Host "Querying Computer's Hard Disk Information..."

## Disk Information
gwmi win32_logicaldisk -ComputerName $targetComputer | select DeviceID,Description,FileSystem,FreeSpace,Size,VolumeDirty,VolumeName,VolumeSerialNumber | ConvertTo-HTML -Fragment | out-file $outputFile -Append

Write-Host "Querying Computer's Network Adapter Information..."

## IP Address Information
gwmi win32_networkadapterconfiguration -ComputerName $targetComputer -filter "DHCPEnabled = True" | select Description,DHCPEnabled,DHCPLeaseObtained,DHCPServer,DNSDomain,DNSHostName,MACAddress,@{Name='IpAddress';Expression={$_.IpAddress -join '; '}},@{Name='DefaultIPgateway';Expression={$_.DefaultIPgateway -join '; '}} | ConvertTo-HTML -Fragment | out-file $outputFile -Append

Write-Host "Querying Computer's Windows Update Information..."

## Windows Updates
gwmi -cl win32_reliabilityRecords -ComputerName $targetComputer -filter "sourcename = 'Microsoft-Windows-WindowsUpdateClient'" | select @{LABEL = "date";EXPRESSION = {$_.ConvertToDateTime($_.timegenerated)}},user, productname | ConvertTo-HTML -Fragment | out-file $outputFile -Append

Write-Host "Querying Computer's Add/Remove Programs List..."

## Add/Remove Programs
gwmi win32Reg_AddRemovePrograms -ComputerName $targetComputer -filter "NOT DisplayName LIKE '%Security Update for%' and NOT DisplayName LIKE '%Service Pack 2 for%'" | select DisplayName,Publisher,Version | sort DisplayName | ConvertTo-HTML -Fragment | out-file $outputFile -Append

Write-Host "Writing End of File..."

$sw.Stop()
$formatTime1 = $sw.Elapsed.ToString()
$formatTime = "<p align='right'>Script ran in:  " + $sw.Elapsed.Minutes.ToString() + " Minutes, " + $sw.Elapsed.Seconds.ToString() + " Seconds, and " + $sw.Elapsed.Milliseconds.ToString() + " Milliseconds. </p>"


Out-File -filePath $outputFile -inputObject $formatTime -Append

### End of File

Exit
