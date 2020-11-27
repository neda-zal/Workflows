Import-Module VMware.VimAutomation.Core
Import-Module VMware.VimAutomation.HorizonView
Import-Module VMware.HV.Helper

$hvServer = connect-hvserver lab-vcs01.lab.loginvsi.com -user neda -pass VSIlab!1 -Domain lab.loginvsi.com 
$Global:hvServices = $hvServer.ExtensionData

$spechelper=New-Object VMware.Hv.DesktopService+DesktopRefreshSpecHelper
$specbase=New-Object VMware.Hv.DesktopRefreshSpec

$PoolVMs = get-hvmachinesummary -PoolName 'WIN10-2004-OSOT'
foreach ($SingleVM in $PoolVMs) {

$VM = Get-VM -Name $SingleVM.Base.Name

Write-Host $SingleVM.Base.Name "is being restarted" -foreground green

Shutdown-VMGuest -VM $VM -Confirm:$false -ErrorAction SilentlyContinue | Wait-Tools
do {

      Start-Sleep -s 5

      $status = $VM.PowerState

   } until($status -eq "PoweredOn")
   
}

Write-Output "Machines are ready" -foreground green
Write-Output "Disconnected from Connection Server."
Disconnect-HVServer -Server $hvServer -Confirm:$false
