<#
Author:  Nathan Behe
Date: 6/10/2016
Version: 1.14
Purpose:
    This script is made to automate Blue/Green Deployment on this server.

.Synopsis
    This script is made to automate Blue/Green deployment on this server.

.Description
    The process works like this:
    1.  Copy the code into the folder corresponding to the site you are deploying to.
        i.e. H:\Sites\CaptorApp
    2.  Look at what site is active at this time (Check in ARR to see which passes the Health Check)
    3.  Call the function
        i.e.
            PS> Publish-CaptorApp
        This will move your code from H:\Sites\CaptorApp into the folder H:\Sites\CaptorApp-*Previously Stopped Site*\ then warm up the site.
        Finally it will make the offline site live and bring the online site back offline by changing the contents of the up.html file.

#>


#Import-Module webadministration -ErrorAction Stop

<#
##  ***TODO***
##  ==  Give option to deploy code from a different location.

Parameter sets:

1.  Specify nothing.  Script finds running site and switches to non-running site.  Removes old backup files and writes new ones.
2.  Give option to roll back last move.
3.  Give option to deploy code from a non-standard location.

#>

Function Publish-CaptorApp{
[CmdletBinding()]
Param(
[Parameter(Mandatory=$False,Position=1)][string]$baseAppName = "CaptorApp",  # Name of the application
[Parameter(Mandatory=$False,Position=2)][string]$blueAppURL = "blue.captorapp", #Blue App URL
[Parameter(Mandatory=$False,Position=3)][string]$greenAppURL = "green.captorapp", #Green App URL
[Parameter(Mandatory=$False,Position=4)][int]$blueAppPort = 8080, #Blue App port number
[Parameter(Mandatory=$False,Position=5)][int]$greenAppPort = 8081, #Green App port number
[Parameter(Mandatory=$False,Position=6)][string]$appFolderPath = "H:\Sites\", #Path to folders
[Parameter(Mandatory=$False,Position=6)][string]$staticHTMLSubPath = "infraapi\wwwroot\" #Path to static HTML
)

Write-Host "Got this far"
$blueSiteState = (Get-WebsiteState -Name ($baseAppName + "-Blue")).Value
$greenSiteState = (Get-WebsiteState -Name ($baseAppName + "-Green")).Value

$baseAppFolder = $appFolderPath + $baseAppName + "\"

$blueWebsiteName = $baseAppName + "-Blue"
$greenWebsiteName = $baseAppName + "-Green"

$blueSiteFolder = $appFolderPath + $blueWebsiteName + "\"
$greenSiteFolder = $appFolderPath + $greenWebsiteName + "\"

$blueSiteUpFolder = $blueSiteFolder + $staticHTMLSubPath + "up.html"
$greenSiteUpFolder = $greenSiteFolder + $staticHTMLSubPath + "up.html"

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
        Set-Content $blueSiteUpFolder "up"

        Write-Host "Bringing down Green site..."
        Set-Content $greenSiteUpFolder "down"

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
        Set-Content $greenSiteUpFolder "up"

        Write-Host "Bringing down Blue site..."
        Set-Content $blueSiteUpFolder "down"

        Write-Host "Waiting 10 seconds to warm up..."
        Start-Sleep -s 10

        Write-Host "Stopping Blue website..."
        Stop-Website $blueWebsiteName

        Write-Host "Done."
    

}


} #End Deploy-CaptorApp

Export-ModuleMember -Function Publish-CaptorApp