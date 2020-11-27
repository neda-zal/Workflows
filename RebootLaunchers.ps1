Set-PowerCLIConfiguration -ProxyPolicy NoProxy -Confirm:$false
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false

$config = Get-Content -Path "$PSScriptRoot\config.json" -Raw | ConvertFrom-Json

$ServersTXTFile = $config.ServersTXTFile 

Write-Host "Connecting to vCenter Server" -foreground green
Connect-VIServer -Server $config.vCenter -Protocol https -User $config.vCenterUser -Password $config.vCenterUserPassword -WarningAction 0

$DomainUser = '$config.DomainUser'
$DomainPWord = ConvertTo-SecureString -String '$config.DomainUserPassword' -AsPlainText -Force
$DomainCredential = New-Object -Typename System.Management.Automation.PSCredential -ArgumentList $DomainUser, $DomainPWord

if (Get-Module -ListAvailable -Name 'VMware.VimAutomation.Core') {
    Write-Host "Module exists"
} 
else {
    Write-Host "Module does not exist"
    Install-Module -Name VMware.VimAutomation.Core
    Import-Module VMware.VimAutomation.Core
}

For ($i = 1; $i -le $config.VMCount; $i++) {

    $y = "{0:d3}" -f $i
    $VM_name = $config.VMPrefix + "-LS" + $y

    write-host "Restart of Launcher $VM_name initiated"  -foreground green
    Restart-VMGuest -VM $VM_name -Confirm:$false
      
}

if(!(Test-Path -Path $PSScriptRoot\$ServersTXTfile)) {
    New-Item -Path $PSScriptroot -Name "Launchers.txt" -ItemType "file"
} else {
    Clear-Content -Path $PSScriptRoot\$ServersTXTfile
}

For ($i = 1; $i -le $config.VMCount; $i++) {

     $y = "{0:d3}" -f $i
     $VMName = $config.VMPrefix + "-LS" + $y

     if ((VMware.vimautomation.core\Get-VM -Name $VMName).State -eq "PoweredOff") {
         vmware.vimautomation.core\Get-VM -Name $VMName | vmware.vimautomation.core\Start-VM -Confirm:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
     }
     
     Add-Content -Path $PSScriptroot\$ServersTXTfile -Value $VMName

     $max = 20
     $v = 0
     while ($v -ne $max) {
         try {
             $process = Invoke-Command -ComputerName $VMName -ScriptBlock {Get-Process -Name "Agent"} -ErrorAction SilentlyContinue
             if ($process) {
                 break
             }
             Start-Sleep -Seconds 10
                
         }
         catch {
             Start-Sleep -Seconds 10
         }

         $v++
     }

     if ($v -gt $max) {
         Write-Error -Message "Launcher $VMName is not running!"
     }
 }

 write-host "Launchers are ready"  -foreground green
