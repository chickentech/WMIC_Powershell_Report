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

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

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


Function Get-WMIQuery($outputFile,$targetComputer){
$isEmpty = [bool]$targetComputer
    if($isEmpty){
        ## Setup Variables
        [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
        $sw = [Diagnostics.Stopwatch]::StartNew() # Start Timer
        $dat = Get-Date
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

        Out-File -filePath $outputFile -InputObject $a

        $e = "<div style='text-align: center;'><h1> Computer:  " + $targetComputer + "</h1>"
        $e = $e + "<p><h3>Report Generated:</h3>Date: " + ($dat.ToShortDateString()) + "<br>Time:  " + ($dat.ToShortTimeString()) + "</p></div>"
        ##Inject Generated HTML into page
        Out-File -filePath $outputFile -InputObject $e -Append

        ### PC Model and Serial # information
        $t1 = gwmi win32_computersystem -ComputerName $targetComputer

        ## OS Architecture x86 or x64
        $t2 = gwmi win32_operatingsystem -ComputerName $targetComputer

        ## Serial Number
        $t3 = gwmi win32_bios -ComputerName $targetComputer

        $tim = $t2.ConvertToDateTime($t2.LastBootUpTime)

        $finalElapsed = New-TimeSpan -Start $tim -End $dat

        $output2 = New-Object PSObject -Property @{
            Manufacturer = $t1.Manufacturer
            Model = $t1.Model
            'Hostname' = $t1.Name
            'Logged In User' = $t1.UserName
            'OS Architecture' = $t2.OSArchitecture
            'Computer Serial Number' = $t3.SerialNumber
            'System Up Time' = "Days: " + $finalElapsed.Days + " Hours: " + $finalElapsed.Hours + " Minutes: " + $finalElapsed.Minutes
            }

        ConvertTo-Html -Fragment -inputObject $output2 | Out-File $outputFile -Append

        ## Disk Information
        gwmi win32_logicaldisk -ComputerName $targetComputer | select DeviceID,Description,FileSystem,FreeSpace,Size,VolumeDirty,VolumeName,VolumeSerialNumber | ConvertTo-HTML -Fragment | out-file $outputFile -Append

        ## IP Address Information
        gwmi win32_networkadapterconfiguration -ComputerName $targetComputer -filter "DHCPEnabled = True and NOT Description like '%Remote%'" | select Description,DHCPEnabled,DHCPLeaseObtained,DHCPServer,DNSDomain,DNSHostName,MACAddress,@{Name='IpAddress';Expression={$_.IpAddress -join '; '}},@{Name='DefaultIPgateway';Expression={$_.DefaultIPgateway -join '; '}} | ConvertTo-HTML -Fragment | out-file $outputFile -Append

        ## Windows Updates
        gwmi -cl win32_reliabilityRecords -ComputerName $targetComputer -filter "sourcename = 'Microsoft-Windows-WindowsUpdateClient'" | select @{LABEL = "date";EXPRESSION = {$_.ConvertToDateTime($_.timegenerated)}}, productname | sort date -descending | ConvertTo-HTML -Fragment | out-file $outputFile -Append

        ## Add/Remove Programs
        gwmi win32Reg_AddRemovePrograms -ComputerName $targetComputer -filter "NOT DisplayName LIKE '%Security Update for%' and NOT DisplayName LIKE '%Service Pack 2 for%'" | select DisplayName,Publisher,Version | sort DisplayName | ConvertTo-HTML -Fragment | out-file $outputFile -Append

        $sw.Stop()
        $formatTime1 = $sw.Elapsed.ToString()
        $formatTime = "<p align='right'>Script ran in:  " + $sw.Elapsed.Minutes.ToString() + " Minutes, " + $sw.Elapsed.Seconds.ToString() + " Seconds, and " + $sw.Elapsed.Milliseconds.ToString() + " Milliseconds. </p>"


        Out-File -filePath $outputFile -inputObject $formatTime -Append

        ii $outputFile

    } else {
        [System.Windows.Forms.MessageBox]::Show("Please Enter a Hostname and select a file save location.","Error",[System.Windows.Forms.MessageBoxButtons]::OKCancel,[System.Windows.Forms.MessageBoxIcon]::Warning)
        $objTextBox.Select()
        return
    }

}

$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "WMI HTML Computer Report"
$objForm.Size = New-Object System.Drawing.Size(300,300) 
$objForm.StartPosition = "CenterScreen"

$fol = New-Object -com Shell.Application
$bfol = ($fol.namespace(0x10)).Self.Path

$browse = New-Object Windows.Forms.SaveFileDialog
$browse.initialDirectory = $bfol
$browse.DefaultExt = "html"  
$browse.filter = "HTML Files|*.html|All files (*.*)|*.*" 
$browse.AddExtension = $true 

$objForm.KeyPreview = $True
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
    {$x=$objTextBox.Text;$objForm.Close()}})
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$objForm.Close()}})

<#
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(50,80)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
$OKButton.Add_Click({$x=$objTextBox.Text;$objForm.Close()})
$objForm.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(160,80)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.Add_Click({$objForm.Close()})
$objForm.Controls.Add($CancelButton)

#>

$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(10,20) 
$objLabel.Size = New-Object System.Drawing.Size(280,20) 
$objLabel.Text = "Please enter the hostname below:"
$objForm.Controls.Add($objLabel) 

$objTextBox = New-Object System.Windows.Forms.TextBox 
$objTextBox.Location = New-Object System.Drawing.Size(10,40) 
$objTextBox.Size = New-Object System.Drawing.Size(260,20) 
$objForm.Controls.Add($objTextBox) 

$objLabelPath = New-Object System.Windows.Forms.Label
$objLabelPath.Location = New-Object System.Drawing.Size(10,120) 
$objLabelPath.Size = New-Object System.Drawing.Size(280,20) 
$objLabelPath.Text = "Select a file save location"
$objForm.Controls.Add($objLabelPath) 

$PathButton = New-Object System.Windows.Forms.Button
$PathButton.Location = New-Object System.Drawing.Size(100,150)
$PathButton.Size = New-Object System.Drawing.Size(80,30)
$PathButton.Text = "Click to Choose Path"
$PathButton.Add_Click({$browse.ShowDialog()})
$objForm.Controls.Add($PathButton)

$ExecuteButton = New-Object System.Windows.Forms.Button
$ExecuteButton.Location = New-Object System.Drawing.Size(195,225)
$ExecuteButton.Size = New-Object System.Drawing.Size(75,30)
$ExecuteButton.Text = "Execute"
$ExecuteButton.Add_Click({Get-WMIQuery -outputFile ($browse.FileName.ToString()) -targetComputer ($objTextBox.Text.ToString())})
$objForm.Controls.Add($ExecuteButton)

$objForm.Topmost = $True

$objForm.Add_Shown({$objForm.Activate()})
# Set Focus to the Hostname Textbox
$objForm.Add_Load({$objTextBox.Select()})
#Prevent Form Resize and Maximize
$objForm.FormBorderStyle = 'Fixed3D'
$objForm.MaximizeBox = $false
[void] $objForm.ShowDialog()

$objTextBox.Text
$browse.FileName.ToString()
