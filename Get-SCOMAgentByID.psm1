#Import-Module OperationsManager
<#
By:  Nathan Behe for PA Department of Corrections
Date:  11/20/2015
Get-SCOMAgentByID
#>

Function Get-SCOMAgentByID{


<#
.SYNOPSIS

.DESCRIPTION
This script will take a GUID and return the computer name that it equates with.

.PARAMETER ID
This is the GUID of the HealthService Object you would like to get the hostname from.

.EXAMPLE
Get-SCOMAgentByID -Id d3081202-3d67-995a-66a7-eb6c2f1fc6df

#>


[CmdletBinding(DefaultParameterSetName="Id",SupportsShouldProcess=$False,ConfirmImpact='Low')]
#gather parameters and support pipelining
Param(
[parameter(Mandatory=$True,Position=0,HelpMessage="GUID ID of Health Service.")]
[STRING[]]$Id="test"
)

$ReturnName = Get-SCOMAgent | where {$_.hostedHealthService.Id -eq $Id[0]} | select name

return $ReturnName

}