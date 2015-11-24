# WMI Powershell Scripts
These are several powershell scripts for work.  It's starting to grow more lately to include many more projects.  This was started originally for:

Used to generate a report similar to an old SCCM 2007 report that will be lost in our upgrade to 2012 SCCM.  Generates a HTML file as output that contains the computer Name, manufacturer, serial number, ip address information, Windows Updates installed, and anything from the Add/Remove Programs (listed in tables in that order).  Tables are formatted so alternating rows have different background color for readability (using CSS).

I might try to replicate this in something other than Powershell so that it won't require people to change their Set-ExecutionPolicy (my fellow techs might not be able to figure out that one-liner).

To run these on your computer, run from an elevated PowerShell Command Prompt:
```sh
Set-ExecutionPolicy Unrestricted
```
Then download and run these and they should work for you.

## Scripts

* **ADComputer_HTML_Report_GUI.ps1** - This is a Powershell script that should work on all clients that have PS V1 on them.  Uses Windows Forms to build a GUI that takes a hostname/ip and file save location as input and then generates and displays a HTML formatted report.
* **ADUser_HTML_Report.ps1** - This is a Powershell script (again written for PS V1) that takes an AD username and file save location as input and displays an Active Directory report in HTML format.  Using Windows Forms yet again.
* **report_commands.ps1** - This is a precursor to the first script that doesn't use Windows Forms and runs through a single time.  Feed it a hostname/ip address and file save location.
*  **Telestaff.ps1** - This is a simple script to fix a registry key for a program we use.  Enables end users to change their environments to DEV, PROD, etc.  To simplify things, the permissions on the registry key must be changed to allow the end user to modify.  Didn't want to hard-code credentials into this script.
*  **Get-SCOMAgentByID.ps1** - This one I'm happy with.  Made a Powershell Cmdlet to allow me to get a scom agent hostname by using the Health Service ID only.  Just making a nice Cmdlet front-end to a simple PS one-liner I use a lot.

----------

Developed by Nathan Behe
