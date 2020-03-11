Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Unrestricted -Force -ErrorAction SilentlyContinue

Import-Module -Name Vmware.VimAutomation.Core -WarningAction SilentlyContinue -ErrorAction Stop

$Script_Parent = Split-Path -Parent $MyInvocation.MyCommand.Definition  
 
#************** Remove old files *************************** 

Remove-item ($Script_Parent + "\Report\P*.html") -Force -Verbose
 
########### Connect VCs from VC_List.txt ############ 
$vCenterList = @()

$vCenterList= (Get-Content  -Path ".\vcenterList.txt")   

$D = get-date -uformat "%d-%m-%Y-%H:%M" # To get a current date. 

$OutDate = get-date -uformat "%d%m%Y-%H%M" 

Write-Host "Connecting to VC" -foregroundcolor yellow 

  $HTML = '<style type="text/css"> 
   #Header{font-family:"Trebuchet MS", Arial, Helvetica, sans-serif;width:100%;border-collapse:collapse;} 
   #Header td, #Header th {font-size:14px;border:1px solid #98bf21;padding:3px 7px 2px 7px;} 
   #Header th {font-size:14px;text-align:center;padding-top:5px;padding-bottom:4px;background-color:#cccccc;color:#000000;} 
   #Header tr.alt td {color:#000;background-color:#EAF2D3;} 
   </Style>' 


foreach ($vCenter in $vCenterList)
{

    Connect-VIServer -Server $vCenter -WarningAction 0 
 
    $outputfile = ($SCRIPT_PARENT + "\Report\PoweredOffVMs_$($OutDate).html") #".\Report\$($VC).html" 

    Write-Host "" 

    Write-Host "Collecting PoweredOff VMs from $vCenter" -foregroundcolor green 

   $Result = @()

         
    $HTML += "<HTML><BODY><Table border=1 cellpadding=0 cellspacing=0 id=Header><caption><font size=3 color=green><h1 align=""center"">Powered Off VMs Report - Vcenter: $vCenter </h1></font> 
            <h4 align=""Right""><font size=3 color=""#00008B"">Date: $D </font></h4></caption>" 
            
            
            
     $clusterNames = @()
     $clusterNames = Get-Cluster -Server $vCenter | Sort-Object -Property Name | Select-Object -ExpandProperty Name

     
     foreach ($clName in $clusterNames){
        
        $vmList = @()

        $vmList = Get-Vm -Location $clName | Where-Object -FilterScript {($_.PowerState -eq "PoweredOff") -or ($_.PowerState -eq "Suspended")} 

        
        $HTML += "<HTML><BODY><Table border=1 cellpadding=0 cellspacing=0 id=Header><caption><font size=3 color=green><h1 align=""center"">PoweredOff VMs: $clName</h1></font> 
                
            <TR> 
                  <TH><B>VM Name</B></TH> 
                  <TH><B>Power State</B></TH> 
                  <TH><B>Notes</B></TH> 
            </TR>" 

       foreach ($vm in $vmList){
       
             $HTML += " 
                            <TR>
                                    <TR bgColor=White>
                                    <TD>$($vm.Name)</TD> 
                                    <TD>$($vm.PowerState)</TD> 
                                    <TD>$($vm.Notes)</TD>
                            </TR>" 

       
       
       }#END FOREACH VM
        

     
     }#END FOREACH CLUSTER
     
     
        $HTML += "</Table></BODY></HTML>" 
        $HTML | Out-File $OutputFile        
        
        Disconnect-VIServer -Server $vCenter -Force -ErrorAction SilentlyContinue -WarningAction Continue  -Confirm:$false   
    
}#END FOREACH VCENTER



#MOVE AND SEND E-MAIL
Move-Item -Path ($Script_Parent + "\Report\P*.html") -Destination "$env:systemdrive\Scripts\Box\Output\Vmware\VM\PoweredOff" -Force

Start-Sleep -Seconds 2

#Send Mail Daily
Clear-Host

Set-Location "$env:systemdrive\Scripts\Box\Output\Vmware\VM\PoweredOff"

$tmpfileFromToday = Get-Date -Format ddMMyyyy

$fileFromToday = $tmpfileFromToday.ToString()

$fileLocation = "$env:systemdrive\Scripts\Box\Output\Vmware\VM\PoweredOff"

#Code to get attachments
<#$tmpAttachment = @()

$attachments = @()

$tmpAttachment = Get-ChildItem | Where-Object -FilterScript {$_.Name -like "*$fileFromToday*"} | Select-Object -ExpandProperty Name

foreach ($tmpAttach in $tmpAttachment){


$attachment = $fileLocation + '\' + $tmpAttach

Write-Output $attachment

$attachments += $attachment

}
#>


$tmpHTML = Get-Content "$env:systemdrive\SCRIPTS\BOX\Process\Mail\contentReportPowerOffVMs.html"

$finalHTML = $tmpHTML | Out-String


###########Define Variables########

$fromaddress = "powershellrobot@yourcompany.com"
$toaddress = "l-microsoft-team@yourcompany.com"
$CCaddress = "yourboss@yourcompany.com"
$HCVMSubject = "[VCENTERS] Powered Off VM Report - VCenters 6.5 and 5.5"
$HCVMattachment = $attachments
$smtpserver = "yourserver.local"

####################################
Send-MailMessage -SmtpServer $smtpserver -From $fromaddress -To $toaddress -Cc $CCaddress -Subject $HCVMSubject -Body $finalHTML -BodyAsHtml -Attachments @(Get-ChildItem -File | Where-Object -FilterScript {$_.Name -like "*$fileFromToday*"}) -Priority Normal -Encoding UTF8