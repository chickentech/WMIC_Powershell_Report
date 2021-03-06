<#  
.SYNOPSIS  
    Script by: Nathan Behe
    For: Office of Administration - Commonwealth of Pennsylvania
    Creation Date: 3/25/2015

.NOTES  
    Version                   : 0.1
    Rights Required           : Domain User
    Powershell Version        : 1.0
    Authors                   : Nathan Behe
    Last Update               : 09-April-2015

.VERSION
    0.1 - Initial Version

.DESCRIPTION
    This script creates a GUI that will accept a hostname and a file save location and then output a HTML file (and then open it) which displays the WMI Queried information.
    I worked very hard to make this only require Powershell v1.0 (and have tested it to work on this version).
       
#>

Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
 
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
 
function Hide-Console {
    $consolePtr = [Console.Window]::GetConsoleWindow()
  #0 hide
 [Console.Window]::ShowWindow($consolePtr, 0)
}


##Check that script is running in STA mode:hi
#Validate that Script is launched

$IsSTAEnabled = $host.Runspace.ApartmentState -eq 'STA'

#Set the name of the console window

#$host.UI.RawUI.WindowTitle="HTML Report Generator"

If ($IsSTAEnabled -eq $false) {

  #Launch script in a separate PowerShell process with STA enabled

  Start-Process -FilePath powershell.exe -ArgumentList "-STA .\ADComputer_HTML_Report_GUI.ps1 -WindowStyle Hidden"

  Exit

} elseif ($IsSTAEnabled -eq $true) {
    Hide-Console | Out-Null
}

## Required assemblies to be able to work with Windows Forms
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

Function Get-WMIQuery($outputFile,$targetComputer){
#******     Clear-Host
$isCompEmpty = [bool]$targetComputer # Cast both as Boolean to see if they're blank or not
$isFileEmpty = [bool]$outputFile
$status.Text = "Sending Ping to check if computer is on..."
$objForm.Refresh()
#******     Write-Host -ForegroundColor Blue -BackgroundColor White -Object "Sending Ping to Computer..."
$isAlive = Test-Connection -ComputerName $targetComputer -Quiet
    if($isCompEmpty -and $isFileEmpty -and ($isAlive -eq $true)){ ## Checks to make sure computername and file name are both non-empty values & that computername responds to pings.
        
        ## Setup Variables
        $sw = [Diagnostics.Stopwatch]::StartNew() # Start Timer
        $dat = Get-Date
        # HTML & CSS Formatting
        $status.Text = "Setting Up HTML File..."
        $objForm.refresh()
        #******     Write-Host -ForegroundColor Green -Object "Setting Up HTML File..."
        $a = "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01 Transitional//EN' 'http://www.w3.org/TR/html4/loose.dtd'>`r`n<html>`r`n<head>`r`n<style>`r`n"
        $a = $a + "BODY{background-color:white;}`r`n" #Background
        $a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;text-align: center;margin-left: auto; margin-right: auto;}`r`n"
        $a = $a + "TH{border-width: 1px;padding: 2px;border-style: solid;border-color: black;}`r`n"
        $a = $a + "TD{border-width: 1px;padding: 2px;border-style: solid;border-color: black;}`r`n"
        $a = $a + "TR:nth-child(even){background-color: #CCC;}`r`n" # Alternate row colors
        $a = $a + "TR:nth-child(odd){background-color: #FFF;}`r`n" #Alternate..
        $a = $a + "</style>`r`n"
        $a = $a + "<!--`r`n"
        $a = $a + "Generated Output from ADComputer_HTML_Report_GUI.ps1 script`r`n"
        $a = $a + "Created by: Nathan Behe`r`n"
        $a = $a + "Contact at: natbehe@pa.gov`r`n"
        $a = $a + "For the Office of Administration Help Desk`r`n`r`n`r`n"
        $a = $a + "-->`r`n"

        # Add HTML Page Title with Computer Name
        $a = $a + "<title>WMI Query Report for Computer:  " + $targetComputer + "</title>`r`n"
        $a = $a + "</head>`r`n"

        Out-File -filePath $outputFile -InputObject $a

        $e = "<body><div style='text-align: center;'><h1> Computer:  " + $targetComputer + "</h1>"
        $e = $e + "<p><h3>Report Generated:</h3>Date: " + ($dat.ToShortDateString()) + "<br>Time:  " + ($dat.ToShortTimeString()) + "</p></div>"
        ##Inject Generated HTML into page
        Out-File -filePath $outputFile -InputObject $e -Append

        $status.Text = "Gathering General PC Information..."
        $objForm.refresh()
        #******     Write-Host -ForegroundColor Green -Object "Gathering General PC Information..."
        trap {[System.Windows.Forms.MessageBox]::Show("Unable to connect to RPC Server.","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning);$objTextBox.SelectAll();return}
        ### PC Model and Serial # information
        $t1 = Get-WmiObject -Class win32_computersystem -ComputerName $targetComputer -ErrorAction Stop

        ## OS Architecture x86 or x64
        $t2 = Get-WmiObject -Class win32_operatingsystem -ComputerName $targetComputer

        ## Serial Number
        $t3 = Get-WmiObject -Class win32_bios -ComputerName $targetComputer

        $tim = $t2.ConvertToDateTime($t2.LastBootUpTime)

        $finalElapsed = New-TimeSpan -Start $tim -End $dat

        $output2 = New-Object -TypeName PSObject -Property @{
            Manufacturer = $t1.Manufacturer
            Model = $t1.Model
            'Hostname' = $t1.Name
            'Logged In User' = $t1.UserName
            'OS Architecture' = $t2.OSArchitecture
            'Computer Serial Number' = $t3.SerialNumber
            'System Up Time' = "Days: " + $finalElapsed.Days + " Hours: " + $finalElapsed.Hours + " Minutes: " + $finalElapsed.Minutes
            } | Select-Object -Property Manufacturer,Model,'Computer Serial Number',Hostname,'Logged In User','OS Architecture','System Up Time'
        

        ConvertTo-Html -Fragment -inputObject $output2 | Out-File -FilePath $outputFile -Append

        $status.Text = "Gathering Disk Information..."
        $objForm.refresh()
        #******     Write-Host -ForegroundColor Green -Object "Gathering Disk Information..."
        ## Disk Information
        Get-WmiObject -Class win32_logicaldisk -ComputerName $targetComputer | Select-Object @{Label = "Drive Letter";Expression = {$_.DeviceID}},Description,FileSystem,@{Label="Free Space (GB)";Expression={"{0:N2}" -f ($_.FreeSpace/1GB)}},@{Label="Total Size (GB)";Expression={"{0:N2}" -f ($_.Size/1GB)}},VolumeDirty,VolumeName,VolumeSerialNumber | ConvertTo-HTML -Fragment | out-file -FilePath $outputFile -Append

        $status.Text = "Gathering Network Adapter Information..."
        $objForm.refresh()
        #******     Write-Host -ForegroundColor Green -Object "Gathering Network Adapter Information..."
        ## IP Address Information
        Get-WmiObject -Class win32_networkadapterconfiguration -ComputerName $targetComputer -filter "DHCPEnabled = True and NOT Description like '%Remote%'" | Select-Object @{Label="Network Adapter";Expression = {$_.Description}},DHCPEnabled,@{Label="DHCP Lease Obtained";Expression={$_.ConvertToDateTime($_.DHCPLeaseObtained)}},DHCPServer,DNSDomain,DNSHostName,MACAddress,@{Name='IP Address';Expression={$_.IpAddress -join '; '}},@{Name='DefaultIPgateway';Expression={$_.DefaultIPgateway -join '; '}} | ConvertTo-HTML -Fragment | out-file -FilePath $outputFile -Append

        $status.Text = "Gathering Windows Update Information..."
        $objForm.refresh()
        #******     Write-Host -ForegroundColor Green -Object "Gathering Windows Update Information..."
        ## Windows Updates
        Get-WmiObject -class win32_reliabilityRecords -ComputerName $targetComputer -filter "sourcename = 'Microsoft-Windows-WindowsUpdateClient'" | Select-Object @{LABEL = "Date";EXPRESSION = {$_.ConvertToDateTime($_.timegenerated)}}, @{Label = "Windows Update";Expression = {$_.productname}} | Sort-Object -Property Date -descending | ConvertTo-HTML -Fragment | out-file -FilePath $outputFile -Append

        $status.Text = "Gathering Add/Remove Programs Information..."
        $objForm.refresh()
        #******     Write-Host -ForegroundColor Green -Object "Gathering Add/Remove Programs Information..."
        ## Add/Remove Programs
        Get-WmiObject -Class win32Reg_AddRemovePrograms -ComputerName $targetComputer -filter "NOT DisplayName LIKE '%Security Update for%' and NOT DisplayName LIKE '%Service Pack 2 for%'" | Select-Object @{Label="Software Title";Expression={$_.DisplayName}},Publisher,Version | Sort-Object -Property "Software Title" | ConvertTo-HTML -Fragment | out-file -FilePath $outputFile -Append

        $sw.Stop()
        $formatTime1 = $sw.Elapsed.ToString()
        $formatTime = "<p align='right'>Script ran in:  " + $sw.Elapsed.Minutes.ToString() + " Minutes, " + $sw.Elapsed.Seconds.ToString() + " Seconds, and " + $sw.Elapsed.Milliseconds.ToString() + " Milliseconds. </p></body></html>"

        $status.Text = "Writing End of File Information..."
        $objForm.refresh()
        #******     Write-Host -ForegroundColor Green -Object "Writing End of File Information..."
        Out-File -filePath $outputFile -inputObject $formatTime -Append
        
        $status.Text = ""
        $objForm.refresh()
        #******     Clear-Host
        $objTextBox.SelectAll()
        Invoke-Item -Path $outputFile
        return

    } ElseIf(!$isCompEmpty -or !$isFileEmpty) {
        [System.Windows.Forms.MessageBox]::Show("Please Enter a Hostname and select a file save location.","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
        #******     Clear-Host
        #******     Write-Host -ForegroundColor Red -Object "Computer Hostname or File Save Location is blank!"
        $objTextBox.SelectAll()
        return
    } ElseIf(!$isAlive){
        [System.Windows.Forms.MessageBox]::Show("Computer selected is not responding to pings.","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
        #******     Clear-Host
        #******     Write-Host -ForegroundColor Red -Object "Computer is not responding to pings!"
        $objTextBox.SelectAll()
        return
        }

}

## Create Form
$objForm = New-Object -TypeName System.Windows.Forms.Form 
$objForm.Text = "WMI HTML Computer Report"
$objForm.Size = New-Object System.Drawing.Size(300,280) 
$objForm.StartPosition = "CenterScreen"
$objForm.ShowInTaskbar = $true

# Create ToolStrip Status Label
$status = New-Object -TypeName System.Windows.Forms.ToolStripStatusLabel

# Find User's Desktop Location
$fol = New-Object -ComObject Shell.Application
$bfol = ($fol.namespace(0x10)).Self.Path

# Initialize Save File Dialog Properties
$browse = New-Object -TypeName Windows.Forms.SaveFileDialog
$browse.initialDirectory = $bfol
$browse.DefaultExt = "html"  
$browse.filter = "HTML Files|*.html|All files (*.*)|*.*" 
$browse.AddExtension = $true

# Handle Enter and Escape Keys
$objForm.KeyPreview = $True
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
    {Get-WMIQuery -outputFile ($objFilePath.Text.ToString()) -targetComputer ($objTextBox.Text.ToString());$objForm.Close()}})
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$objForm.Close()}})

#Label for Hostname Box
$objLabel = New-Object -TypeName System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(10,20) 
$objLabel.Size = New-Object System.Drawing.Size(280,20) 
$objLabel.Text = "Please enter the hostname below:"
$objForm.Controls.Add($objLabel) 

# Label for Save File Box
$objSaveLabel = New-Object -TypeName System.Windows.Forms.Label
$objSaveLabel.Location = New-Object System.Drawing.Size(10,117)
$objSaveLabel.Size = New-Object System.Drawing.Size(120,17)
$objSaveLabel.Text = "File Save Location:"
$objForm.Controls.Add($objSaveLabel)

# Textbox for Hostname
$objTextBox = New-Object -TypeName System.Windows.Forms.TextBox 
$objTextBox.Location = New-Object System.Drawing.Size(10,40) 
$objTextBox.Size = New-Object System.Drawing.Size(260,20) 
$objForm.Controls.Add($objTextBox) 

# Button to choose a save location
$PathButton = New-Object -TypeName System.Windows.Forms.Button
$PathButton.Location = New-Object System.Drawing.Size(160,100)
$PathButton.Size = New-Object System.Drawing.Size(80,30)
$PathButton.Text = "Save Location"
$PathButton.Add_Click({$browse.ShowDialog();$objFilePath.Text=$browse.FileName.ToString()})
$objForm.Controls.Add($PathButton)

# Textbox that displays the file save location (may be edited directly)
$objFilePath = New-Object -TypeName System.Windows.Forms.TextBox
$objFilePath.Location = New-Object System.Drawing.Size(10,135)
$objFilePath.Size = New-Object System.Drawing.Size(260,20)
$objFilePath.Text = ""
$objForm.Controls.Add($objFilePath)

## Add a statusbar at the bottom to display what the script is doing while running.
$objStatusBar = New-Object -TypeName System.Windows.Forms.StatusStrip
$status.Text = ""
$objStatusBar.SizingGrip = $false
[void]$objStatusBar.Items.Add($status)
$objForm.Controls.Add($objStatusBar)

# Execute Button
$ExecuteButton = New-Object -TypeName System.Windows.Forms.Button
$ExecuteButton.Location = New-Object System.Drawing.Size(195,180)
$ExecuteButton.Size = New-Object System.Drawing.Size(75,30)
$ExecuteButton.Text = "Execute"
$ExecuteButton.Add_Click({Get-WMIQuery -outputFile ($objFilePath.Text.ToString()) -targetComputer ($objTextBox.Text.ToString())})
$objForm.Controls.Add($ExecuteButton)

#$objForm.Topmost = $True

$objForm.Add_Shown({$objForm.Activate()})
# Set Focus to the Hostname Textbox
$objForm.Add_Load({$objTextBox.Select()})
#Prevent Form Resize and Maximize
$objForm.FormBorderStyle = 'Fixed3D'
$objForm.MaximizeBox = $false
[void] $objForm.ShowDialog()
