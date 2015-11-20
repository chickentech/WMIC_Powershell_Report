##
#
#
##

function is64bit() {
  return ([IntPtr]::Size -eq 8)
}

$key = 'HKLM:\SOFTWARE\Classes\AppID\{CA12B971-8FC0-405F-A772-FCA2E37ECDF2}'
$eestore = ((Get-ItemProperty -Path $key).RunAs).ToString()

#IP Addresses
$QA_IP = "172.22.28.7"
$DEV_IP = "172.22.28.8"
$Train_IP = "172.22.28.6"
$Prod_IP = "172.22.19.190"

while(1){
Write-Host "Select your environment:"
Write-Host "1) QA Telestaff Environment"
Write-Host "2) DEV Telestaff Environment"
Write-Host "3) Training Telestaff Environment"
Write-Host "4) PROD Telestaff Environment"
$Selection = Read-Host "Please enter the number of your selection"

if (($Selection -as [int]) -eq 1 -or 2 -or 3 -or 4){
    break;
} else {
    Write-Host "Invalid Selection.";
    Start-Sleep 2;
    }
}

#Set reg key to chosen environment
Switch (($Selection -as [int])){
    1 {Write-Host "Setting Environment to QA"; Set-ItemProperty -Path $key -Name RemoteServerName -Value $QA_IP; break;}
    2 {Write-Host "Setting Environment to DEV"; Set-ItemProperty -Path $key -Name RemoteServerName -Value $DEV_IP; break;}
    3 {Write-Host "Setting Environment to Training"; Set-ItemProperty -Path $key -Name RemoteServerName -Value $Train_IP; break;}
    4 {Write-Host "Setting Enviornment to Production"; Set-ItemProperty -Path $key -Name RemoteServerName -Value $Prod_IP; break;}
    default { Write-Host "An error has occurred."; exit; break;}
}

#Determine is x86 or x64
function get-programfilesdir() {
  if (is64bit -eq $true) {
    (Get-Item "Env:ProgramFiles(x86)").Value
  }
  else {
    (Get-Item "Env:ProgramFiles").Value
  }
}

$telePath = get-programfilesdir;
$telePath = $telePath.tostring() + "\Telestaff\client\TeleStaff.exe";
Write-Host $telePath;
#Launch Telestaff
Invoke-Item -Path $telePath;
Exit;
