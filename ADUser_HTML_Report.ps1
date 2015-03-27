Import-Module ActiveDirectory
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

  Start-Process powershell.exe -ArgumentList "-STA .\ADUser_HTML_Report.ps1"

  Exit

}

Import-Module ActiveDirectory
<#
.Synopsis
    Script by: Nathan Behe
    For: Office of Administration - Commonwealth of Pennsylvania

.Description
    This script takes a username as input and will output a HTML document

#>

## Insert Module to get File Save Location and Username to use

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
$ourUser = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the username.", "CWOPA Username", "")


## Declare Variables
$adServer = "enhbgdc502.pa.lcl"
$dat = Get-Date

### HTML File Info
### Head Section
# HTML & CSS Formatting
$a = "<html><head><style>"
$a = $a + "BODY{background-color:white;}" #Background
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;text-align: center;margin-left: auto; margin-right: auto;}"
$a = $a + "TH{border-width: 1px;padding: 2px;border-style: solid;border-color: black;}"
$a = $a + "TD{border-width: 1px;padding: 2px;border-style: solid;border-color: black;}"
$a = $a + "TR:nth-child(even){background-color: #CCC;}" # Alternate row colors
$a = $a + "TR:nth-child(odd){background-color: #FFF;}" #Alternate..
$a = $a + "</style>"

# Add JQuery
$a = $a + "<script type='text/javascript' src='http://ajax.googleapis.com/ajax/libs/jquery/1.3/jquery.min.js'></script></head>"
Out-File -filePath $outputFile -InputObject $a
## End Head Section
##Generate HTML to inject into page - Heading

$e = "<body><div style='text-align: center;'><h1> User:  " + $ourUser + "</h1>"
$e = $e + "<p><h3>Report Generated:</h3>Date: " + ($dat.ToShortDateString()) + "<br>Time:  " + ($dat.ToShortTimeString()) + "</p></div>"
##Inject Generated HTML into page
Out-File -filePath $outputFile -InputObject $e -Append

## First Block
## 1. Distinguished Name
## 2. When Account was created
## 3. When Account was last changed
get-aduser -server $adServer -identity $ourUser -Properties DistinguishedName,whenCreated,whenChanged | select DistinguishedName,whenCreated,whenChanged | ConvertTo-HTML -Fragment | out-file $outputFile -Append

## Second Block
## 1. Logon Name          4. Description
## 2. Department          5. Telephone 
## 3. City                
get-aduser -server $adServer -Identity $ourUser -Properties CN,Company,City,Department,Description,telephoneNumber | select CN,Company,City,Department,Description,telephoneNumber | ConvertTo-HTML -Fragment | out-file $outputFile -Append

## Third Block
## 1. Account Locked ? 
## 2. Account Enabled ?
Get-ADUser -Server $adServer -Identity $ourUser -Properties LockedOut,Enabled | select @{Name = "Account Locked Out?"; Expression = {$_.LockedOut}},@{Name = "Account Enabled";Expression = {$_.Enabled}} | ConvertTo-HTML -Fragment | out-file $outputFile -Append

## Fourth Block
## 1. Time Pwd Changed    4. Pwd Expires
## 2. Pwd Age             5. Pwd Expired
## 3. Lockout Time        6. AutoUnlock Time

$endTime = ([datetime]::FromFileTime((Get-ADUser -Server $adServer -Identity $ourUser -Properties pwdLastSet | select pwdLastSet).pwdLastSet[0])).addDays(60)

Get-ADUser -Server $adServer -Identity $ourUser -Properties PasswordLastSet, lockoutTime, pwdLastSet | select PasswordLastSet,lockoutTime,@{Name="Password Expires (In Days)";Expression={(New-TimeSpan -Start (Get-Date) -End $endTime).Days}} | ConvertTo-HTML -Fragment | out-file $outputFile -Append

## Fifth Block
## 1. Home Directory    3. Home Drive Letter
## 2. Roaming Profile?  4. Logon Script
Get-ADUser -Server $adServer -Identity $ourUser -Properties HomeDirectory, ProfilePath,HomeDrive,ScriptPath | select HomeDirectory,@{Name="Roaming Profile Path (If Exists)";Expression={$_.ProfilePath}},HomeDrive,ScriptPath | ConvertTo-HTML -Fragment | out-file $outputFile -Append

## Get Groups for user
$mygroups = $((Get-Aduser -server $adServer -identity $ourUser -Properties *).MemberOf -split "," | select-string -SimpleMatch "CN=") -replace "CN=",""
$mygroups = $mygroups -join ", "
$myOutput = "<p align=center>" + $mygroups + "</p></body></html>"
Out-File -Filepath $outputFile -InputObject $myOutput -Append
