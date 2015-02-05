# WMIC_Powershell_Report
This is a powershell report for work.

Used to generate a report similar to an old SCCM 2007 report that will be lost in our upgrade to 2012 SCCM.
Generates a HTML file as output that contains the computer Name, manufacturer, serial number, ip address information,
Windows Updates installed, and anything from the Add/Remove Programs (listed in tables in that order).
Tables are formatted so alternating rows have different background color for readability (using JQuery & CSS).

I might try to replicate this in something other than Powershell so that it won't require people to change their 
Set-ExecutionPolicy (my fellow techs might not be able to figure out that one-liner).

Developed by Nathan Behe