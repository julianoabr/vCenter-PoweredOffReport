# VCenterPoweredOffReport
This Script verify all powered of in one or more VCenters and send an e-mail

**You need to put the file ContentReportPowerOffVMs_Git.html in a specific folder**

In script the content of this file is loaded in line 132:

$tmpHTML = Get-Content "$env:systemdrive\SCRIPTS\BOX\Process\Mail\contentReportPowerOffVMs.html"


**You have to create a folder to copy the file generated by report. I suggest the following folder:**

This is in line 104 of Script

Set-Location "$env:systemdrive\Scripts\Box\Output\Vmware\VM\PoweredOff"

