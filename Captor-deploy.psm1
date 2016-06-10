<#
Author:  Nathan Behe
Date: 6/8/2016
Purpose:
    This script is made to automate Blue/Green Deployment on this server.

.Synopsis
    This script is made to automate Blue/Green deployment on this server.

.Description
    The process works like this:
    1.  Copy the code into the folder corresponding to the site you are deploying to.
        i.e. H:\Sites\CaptorForm  or  H:\Sites\CaptorIWA
    2.  Look at what site is active at this time (Check in ARR to see which passes the Health Check)
    3.  Call the function for the site you are deploying to and pass it the name of the color you want to deploy to.
        i.e.  If you are deploying to CaptorForm and the currently active site is Blue, then you would call:
            PS> Import-Module [PATH-TO-MODULE]\Captor-Deploy.psm1
            PS> Publish-CaptorForm Green
        This will move your code from H:\Sites\CaptorForm into the folder H:\Sites\CaptorForm-Green\ then warm up the site.
        Finally it will make the Green site live and bring Blue back offline by changing the contents of the up.html file.

#>


#Import-Module webadministration -ErrorAction Stop

<#
##  ***TODO***
##  ==  Make this into a real Module with optional parameters - make it so if it is run with no parameters, it finds
##      the running site and shuts it down and switches to the non-running site.

##  ==  Pull this together into a single Cmdlet and specify "form" or "iwa" for the site to deploy to.

##  ==  Give option to deploy code from a different location.

##  ==  Finally, (down the road) give an option to roll back the last move and restore the backup and bring it back online.

Parameter sets:

1.  Specify nothing.  Script finds running site and switches to non-running site.  Removes old backup files and writes new ones.
2.  Give option to roll back last move.
3.  Give option to deploy code from a non-standard location.

#>

Function Publish-CaptorForm($color){



if($color.ToLower() -eq "blue"){ ## Deploy to Blue

        Write-Host "Removing old Blue site backup from folder H:\Sites\Backup\CaptorForm-Blue-Backup\*"
        Remove-Item -Path H:\Sites\Backup\CaptorForm-Blue-Backup\* -Recurse -Force

        Write-Host "Copying old files out of folder to H:\Sites\Backup\CaptorForm-Blue-Backup\"
        Move-Item -Path H:\Sites\CaptorForm-Blue\* -Destination H:\Sites\Backup\CaptorForm-Blue-Backup\ -Force


        Write-Host "Copying raw files to H:\Sites\CaptorForm-Blue\..."
        #$destination.CopyHere($deploymentFilesPath + "\*", [System.Int32]1556)
        Copy-Item -Path H:\Sites\CaptorForm\* -Filter *.* -Destination H:\Sites\CaptorForm-Blue\ -Recurse -Force

        Write-Host "Starting Blue Site..."
        Start-Website CaptorForm-Blue

        Write-Host "Warming up Blue site"
        #wake up deployment site
        $url = "http://blue.captorform:8080"
        Write-Host "Starting deployment website $url..."
        $page = (New-Object System.Net.WebClient).DownloadString($url)


        Write-Host "Bringing Blue site up..."
        Set-Content H:\Sites\CaptorForm-Blue\up.html "up"

        Write-Host "Bringing down Green site..."
        Set-Content H:\Sites\CaptorForm-Green\up.html "down"

        Write-Host "Waiting 10 seconds to warm up..."
        Start-Sleep -s 10

        Write-Host "Stopping Green website..."
        Stop-Website CaptorForm-Green

        Write-Host "Done."
    

} elseif($color.ToLower() -eq "green") { ##Deploy to Green

        Write-Host "Removing old Blue site backup from folder H:\Sites\Backup\CaptorForm-Green-Backup\*"
        Remove-Item -Path H:\Sites\Backup\CaptorForm-Green-Backup\* -Recurse -Force

        Write-Host "Copying old files out of folder to H:\Sites\Backup\CaptorForm-Green-Backup\"
        Move-Item -Path H:\Sites\CaptorForm-Green\* -Destination H:\Sites\Backup\CaptorForm-Green-Backup\ -Force

        Write-Host "Copying raw files to H:\Sites\CaptorForm-Green\..."
        #$destination.CopyHere($deploymentFilesPath + "\*", [System.Int32]1556)
        Copy-Item -Path H:\Sites\CaptorForm\* -Filter *.* -Destination H:\Sites\CaptorForm-Green\ -Recurse -Force

        Write-Host "Starting Green Site..."
        Start-Website CaptorForm-Green

        Write-Host "Warming up Green site"
        #wake up deployment site
        $url = "http://green.captorform:8081"
        Write-Host "Starting deployment website $url..."
        $page = (New-Object System.Net.WebClient).DownloadString($url)

        Write-Host "Bringing Green site up..."
        Set-Content H:\Sites\CaptorForm-Green\up.html "up"

        Write-Host "Bringing down Blue site..."
        Set-Content H:\Sites\CaptorForm-Blue\up.html "down"

        Write-Host "Waiting 10 seconds to warm up..."
        Start-Sleep -s 10

        Write-Host "Stopping Blue website..."
        Stop-Website CaptorForm-Blue

        Write-Host "Done."
    

}


} #End Deploy-CaptorForm

Function Publish-CaptorIWA($color){

if($color.ToLower() -eq "blue"){ ## Deploy to Blue

        Write-Host "Removing old Blue site backup from folder H:\Sites\Backup\CaptorIWA-Blue-Backup\*"
        Remove-Item -Path H:\Sites\Backup\CaptorIWA-Blue-Backup\* -Recurse -Force


        Write-Host "Copying old files out of folder to H:\Sites\Backup\CaptorIWA-Blue-Backup\"
        Move-Item -Path H:\Sites\CaptorIWA-Blue\* -Destination H:\Sites\Backup\CaptorIWA-Blue-Backup\ -Force

        Write-Host "Copying raw files to H:\Sites\CaptorIWA-Blue\..."
        #$destination.CopyHere($deploymentFilesPath + "\*", [System.Int32]1556)
        Copy-Item -Path H:\Sites\CaptorIWA\* -Filter *.* -Destination H:\Sites\CaptorIWA-Blue\ -Recurse -Force

        Write-Host "Starting Blue Site..."
        Start-Website CaptorIWA-Blue

        Write-Host "Warming up Blue site"
        #wake up deployment site
        $url = "http://blue.captoriwa:90"
        Write-Host "Starting deployment website $url..."
        $page = (New-Object System.Net.WebClient).DownloadString($url)

        Write-Host "Bringing Blue site up..."
        Set-Content H:\Sites\CaptorIWA-Blue\up.html "up"

        Write-Host "Bringing down Green site..."
        Set-Content H:\Sites\CaptorIWA-Green\up.html "down"

        Write-Host "Waiting 10 seconds to warm up..."
        Start-Sleep -s 10

        Write-Host "Stopping Green website..."
        Stop-Website CaptorIWA-Green

        Write-Host "Done."
    

} elseif($color.ToLower() -eq "green") { ##Deploy to Green

        Write-Host "Removing old Blue site backup from folder H:\Sites\Backup\CaptorIWA-Green-Backup\*"
        Remove-Item -Path H:\Sites\Backup\CaptorIWA-Green-Backup\* -Recurse -Force

        Write-Host "Copying old files out of folder to H:\Sites\Backup\CaptorIWA-Green-Backup\"
        Move-Item -Path H:\Sites\CaptorIWA-Green\* -Destination H:\Sites\Backup\CaptorIWA-Green-Backup\ -Force

        Write-Host "Copying raw files to H:\Sites\CaptorIWA-Green\..."
        #$destination.CopyHere($deploymentFilesPath + "\*", [System.Int32]1556)
        Copy-Item -Path H:\Sites\CaptorIWA\* -Filter *.* -Destination H:\Sites\CaptorIWA-Green\ -Recurse -Force

        Write-Host "Starting Green Site..."
        Start-Website CaptorIWA-Green

        Write-Host "Warming up Green site"
        #wake up deployment site
        $url = "http://green.captoriwa:91"
        Write-Host "Starting deployment website $url..."
        $page = (New-Object System.Net.WebClient).DownloadString($url)

        Write-Host "Bringing Green site up..."
        Set-Content H:\Sites\CaptorIWA-Green\up.html "up"

        Write-Host "Bringing down Blue site..."
        Set-Content H:\Sites\CaptorIWA-Blue\up.html "down"

        Write-Host "Waiting 10 seconds to warm up..."
        Start-Sleep -s 10

        Write-Host "Stopping Blue website..."
        Stop-Website CaptorIWA-Blue

        Write-Host "Done."    

}

} #End Deploy-CaptorIWA







Function Publish-Captor{
[CmdletBinding(DefaultParameterSetName='IWA')]
Param(
[Parameter(ParameterSetName='IWA',Position=1)][Parameter(ParameterSetName='FORM',Position=0)][ValidateSet("iwa","form")][string]$whichSite = "iwa",  # Which App
[Parameter(ParameterSetName='IWA',Position=3)][Parameter(ParameterSetName='FORM',Position=2)][string]$blueAppURL = "blue.captoriwa", #Blue App URL
[Parameter(ParameterSetName='IWA',Position=4)][Parameter(ParameterSetName='FORM',Position=3)][string]$greenAppURL = "green.captoriwa", #Green App URL
[Parameter(ParameterSetName='IWA',Position=5)][Parameter(ParameterSetName='FORM',Position=4)][int]$blueAppPort = 90, #Blue App port number
[Parameter(ParameterSetName='IWA',Position=6)][Parameter(ParameterSetName='FORM',Position=5)][int]$greenAppPort = 91, #Green App port number
[Parameter(ParameterSetName='IWA',Position=7)][Parameter(ParameterSetName='FORM',Position=6)][string]$appFolderPath = "H:\Sites\" #Path to folders
)

if($whichSite.ToLower() -eq "iwa"){
    $baseAppName = "CaptorIWA"

} elseif ($whichSite.ToLower() -eq "form"){
    $baseAppName = "CaptorForm"
    Write-Host "Setting default variables for the form deployment."
    $baseAppName = "CaptorForm"
    $blueAppURL = "blue.captorform"
    $greenAppURL = "green.captorform"
    $blueAppPort = 8080
    $greenAppPort = 8081
}


Write-Host "Got this far"
$blueSiteState = (Get-WebsiteState -Name ($baseAppName + "-Blue")).Value
$greenSiteState = (Get-WebsiteState -Name ($baseAppName + "-Green")).Value

$baseAppFolder = $appFolderPath + $baseAppName + "\"

$blueWebsiteName = $baseAppName + "-Blue"
$greenWebsiteName = $baseAppName + "-Green"

$blueSiteFolder = $appFolderPath + $blueWebsiteName + "\"
$greenSiteFolder = $appFolderPath + $greenWebsiteName + "\"

$blueAppFullURL = "http://" + $blueAppURL + ":" + $blueAppPort
$greenAppFullURL = "http://" + $greenAppURL + ":" + $greenAppPort



if($blueSiteState -eq "Stopped"){ ## Deploy to Blue

        $blueBackupPath = $appFolderPath + "Backup\" + $baseAppName + "-Blue-Backup\"

        Write-Host "Removing old Blue site backup from folder $blueBackupPath"
        Remove-Item -Path ($blueBackupPath + "*") -Recurse -Force

        Write-Host "Copying old files out of folder to $blueBackupPath"
        Move-Item -Path ($blueSiteFolder + "*") -Destination $blueBackupPath -Force


        Write-Host "Copying raw files to $blueSiteFolder..."
        #$destination.CopyHere($deploymentFilesPath + "\*", [System.Int32]1556)
        Copy-Item -Path ($baseAppFolder + "*") -Filter *.* -Destination $blueSiteFolder -Recurse -Force

        Write-Host "Starting Blue Site..."
        Start-Website $blueWebsiteName
        

        Write-Host "Warming up Blue site"
        #wake up deployment site
        Write-Host "Starting deployment website $blueAppFullURL..."
        $page = (New-Object System.Net.WebClient).DownloadString($blueAppFullURL)


        Write-Host "Bringing Blue site up..."
        Set-Content ($blueSiteFolder + "up.html") "up"

        Write-Host "Bringing down Green site..."
        Set-Content ($greenSiteFolder + "up.html") "down"

        Write-Host "Waiting 10 seconds to warm up..."
        Start-Sleep -s 10

        Write-Host "Stopping Green website..."
        Stop-Website $greenWebsiteName

        Write-Host "Done."
    

} elseif($greenSiteState -eq "Stopped") { ##Deploy to Green

        $greenBackupPath = $appFolderPath + "Backup\" + $baseAppName + "-Green-Backup\"

        Write-Host "Removing old Green site backup from folder $greenBackupPath"
        Remove-Item -Path ($greenBackupPath + "*") -Recurse -Force

        Write-Host "Copying old files out of folder to $greenBackupPath"
        Move-Item -Path ($greenSiteFolder + "*") -Destination $greenBackupPath -Force


        Write-Host "Copying raw files to $greenSiteFolder..."
        #$destination.CopyHere($deploymentFilesPath + "\*", [System.Int32]1556)
        Copy-Item -Path ($baseAppFolder + "*") -Filter *.* -Destination $greenSiteFolder -Recurse -Force

        Write-Host "Starting Green Site..."
        Start-Website $greenWebsiteName
        

        Write-Host "Warming up Green site"
        #wake up deployment site
        Write-Host "Starting deployment website $greenAppFullURL..."
        $page = (New-Object System.Net.WebClient).DownloadString($greenAppFullURL)


        Write-Host "Bringing Green site up..."
        Set-Content ($greenSiteFolder + "up.html") "up"

        Write-Host "Bringing down Blue site..."
        Set-Content ($blueSiteFolder + "up.html") "down"

        Write-Host "Waiting 10 seconds to warm up..."
        Start-Sleep -s 10

        Write-Host "Stopping Blue website..."
        Stop-Website $blueWebsiteName

        Write-Host "Done."
    

}


} #End Deploy-CaptorApp