$config = Get-Content -Path "$PSScriptRoot\config.json" -Raw | ConvertFrom-Json


$ApiSecret = $config.PublicAPISecret
$TestID = $config.TestID 
$ServersTXTFile = $config.ServersTXTFile 
$OutPutFile = $config.PoolName + " " + [string](Get-Date -format "yyyy-MM-d hh:mm:ss") + ".csv"

New-Item -Path $PSScriptRoot -Name "$OutPutFile" -ItemType "file"

Write-Host $PSScriptRoot\$OutPutFile

$baseUrl = $config.BaseApplianceURL + "/publicApi"


function Get-CounterStatsPlus {

[CmdLetBinding(DefaultParameterSetName = "NormalRun")]
Param(
    [Parameter(Mandatory = $False, Position = 1, ParameterSetName = "NormalRun")][string]$ServersTXTfile,
    [Parameter(Mandatory = $False, Position = 2, ParameterSetName = "NormalRun")][int]$NumberOfSamples = 5,
    [Parameter(Mandatory = $False, Position = 3, ParameterSetName = "NormalRun")][string]$OutputFile,
    [Parameter(Mandatory = $False, Position = 4, ParameterSetName = "NormalRun")][switch]$IncludeFullCounterPath,
    [Parameter(Mandatory = $false, Position = 5, ParameterSetName = "CheckOnly")][switch]$CheckVersion
)

$stopwatch = [system.diagnostics.stopwatch]::StartNew()

$DebugPreference = "Continue"

$ErrorActionPreference = "SilentlyContinue"

$ScriptVersion = "1.0"
$ScriptName = $MyInvocation.MyCommand.Name
If ($CheckVersion) {Write-Host "SCRIPT NAME     : $ScriptName `nSCRIPT VERSION  : $ScriptVersion";exit}
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$OutputReport = "$ScriptPath\$($ScriptName)_$(get-date -f yyyy-MM-dd-hh-mm-ss).csv"

$Answer = ""

$MyCounters = @"
Processor(_total)\% processor time 
System\Processor Queue Length
Memory\Available MBytes
Memory\Pages/sec
Memory\% Committed Bytes In Use
Network Interface(_total)\Bytes Total/sec
Network Interface(*)\Bytes Received/sec
Network Interface(*)\Bytes Sent/sec
Terminal Services Session(*)\% Processor Time
Terminal Services(*)\Active Sessions
Terminal Services(*)\Inactive Sessions
Terminal Services(*)\Total Sessions
LogicalDisk(*)\% Disk Time
LogicalDisk(*)\Avg. Disk Queue Length
LogicalDisk(*)\% Free Space
LogicalDisk(*)\Avg. Disk sec/Read
LogicalDisk(*)\Avg. Disk sec/Write
LogicalDisk(*)\Disk Reads/sec
LogicalDisk(*)\Disk Transfers/sec
LogicalDisk(*)\Disk Writes/sec
LogicalDisk(*)\Free Megabytes
Processor(_Total)\% Processor Time
Network Adapter(*)\Bytes Received/sec
Network Adapter(*)\Bytes Sent/sec
PhysicalDisk(*)\Avg. Disk Bytes/Read
PhysicalDisk(*)\Avg. Disk Bytes/Write
PhysicalDisk(*)\Avg. Disk sec/Write
PhysicalDisk(*)\Avg. Disk sec/Read
PhysicalDisk(*)\Avg. Disk Bytes/Transfer
PhysicalDisk(*)\Avg. Disk sec/Transfer
"@

<# -------------------------- FUNCTIONS -------------------------- #>
#region Functions region
#Function to have the customized output in CSV format

function Global:Convert-HString {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)] [String]$HString
        )

    Begin 
    {Write-Verbose "Converting Here-String to Array"}
    Process 
    {
        $HString -split "`n" | ForEach-Object {
            $Item = $_.trim()
            #NOTE: below is to enable the use of hashtag to comment aka ignore #lines in your txt file...
            if ($Item -notmatch "#")
            {$Item}    
        }
    }#Process
    End 
    {
        # Nothing to do here.
    }
}#Convert-HString

#Performance counters declaration
function Get-CounterStats { 
    param(
        [String[]]$ComputerName = $Env:ComputerName
    ) 

    (Get-Counter -ComputerName $ComputerName -Counter $(Convert-HString $MyCounters)).counterSamples | ForEach-Object {
        $path = $_.path
        $PropertyHash=@{
                WholeCounter = $path;
                ComputerName=($Path -split "\\")[2];
                Instance = $_.InstanceName ;
                Value = [Math]::Round($_.CookedValue,2) 
                DateTime=(Get-Date -format "yyyy-MM-d hh:mm:ss")
        }

        $PropertyHash.Add('CounterCategory',$(($path  -split "\\")[3]))
        $PropertyHash.Add('CounterName',$(($path  -split "\\")[4]))

    New-Object PSObject -Property $PropertyHash
    }
}

function IsEmpty($Param){
    If ($Param -eq "All" -or $Param -eq "" -or $Param -eq $Null -or $Param -eq 0) {
        Return $True
    } Else {
        Return $False
    }
}

function IsPSV3 {

    $PowerShellMajorVersion = $PSVersionTable.PSVersion.Major
    $msgPowershellMajorVersion = "You're running Powershell v$PowerShellMajorVersion"
    Write-Host $msgPowershellMajorVersion -BackgroundColor blue -ForegroundColor yellow
    If($PowerShellMajorVersion -le 2){
        Return $false
    } Else {
        Return $true
        }
}

<# -------------------------- EXECUTIONS -------------------------- #>

If (!(IsPSV3)){
    $errMsg = "Sorry, you need PSV3 or more recent to run this script."
    $errMsg += "`nBecause we use Export-CSV with the -APPEND property, which exist only starting Powershell V3."
    Write-host $errMsg
    exit
}

If (IsEmpty $OutputFile){$OutputFile = $OutputReport}

If (IsEmpty $PSScriptRoot\$ServersTXTfile){
    $MsgErrInputFile = "No ServersTXTfile specified - collecting counters on local machine.`nCollect counters from the local machine ? (Y/N)"
    while ($Answer -ne "Y" -AND $Answer -ne "N") {
        cls
        Write-Host $MsgErrInputFile -BackgroundColor Yellow -ForegroundColor Red
        $Answer = Read-host
        If($Answer -eq "N"){Exit} Else {[array]$Servers = $($Env:COMPUTERNAME)}
    }
} Else {
    If (!(Test-Path -Path $PSScriptRoot\$ServersTXTfile)){
        $MsgErrInputFile = "Input file with server names doesn't exist.`n"
        $MsgErrInputFile += "Please specify a valid file path and name with -ServersTXTfile parameter.`n"
        $MsgErrInputFile += "Or don't specify the -ServersTXTfile parameter to collect counters on local machine."
        Write-Host $MsgErrInputFile -BackgroundColor yellow -ForegroundColor red
        exit
    }
    [string[]]$servers = get-content $PSScriptRoot\$ServersTXTFile
    $FinServers = @()
    $Servers | Foreach {
        #Regular expression inside the IF to ignore Blank lines or
        #lines with Spaces or TABs characters on beginning and/or on end 
        If ($_ -notmatch "^\s*$"){
            $FinServers += $_.trim()
        }
    $Servers = $FinServers
    }
    $FinServers = $null # a little bit of variable cleanup cannot harm
}

Write-Host "Gathering performance counters for $($Servers -Join ", ")"
Write-Host "That's a total of $($Servers.count) servers"

#Collecting counter information for target servers

$Expression = "Get-CounterStats -ComputerName `$Servers -Counter `$MyCounters | Select-Object ComputerName,DateTime,"
If ($IncludeFullCounterPath) {$expression += "WholeCounter,"}
$Expression += "CounterCategory,CounterName,Instance,Value | Export-Csv -Path `$OutputFile -Append -NoTypeInformation"

while ( $ltStatus = (Get-ltStatus -ErrorAction SilentlyContinue ) -eq "Running" ) {

    if(Get-ltStatus -ErrorAction) {
        SendErrorMessage
    }

    invoke-expression $Expression
 
    Write-Host '.' -NoNewline
    Start-Sleep -Seconds 10
}

Write-Host "File exported to: $outputFile at" (get-date)


<# -------------------------- CLEANUP VARIABLES -------------------------- #>
$OutputFile = $null
$Expression = $null

$stopwatch.Stop()
Write-Host "`n`nThe script took $($StopWatch.Elapsed.TotalSeconds) seconds to execute..."
$StopWatch = $null

}

function Get-ltStatus {
[string]$publicApiSecret = $ApiSecret

$environmentId = $testID

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11

[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { return $true; }

$requestHeaders = @{ 'Content-Type' = 'application/json' }

$authRequest = ConvertTo-Json -InputObject @{ secret = $publicApiSecret }
$authResponse = Invoke-RestMethod -Method Post -Uri "${baseUrl}/v2/authentication/token" -Headers $requestHeaders -Body $authRequest

$requestHeaders.Add('Authorization', "Bearer $($authResponse.accessToken)")

$requestBody = ConvertTo-Json -InputObject $requestObject

    try {
        $response = Invoke-RestMethod -Method Get -Uri "${baseUrl}/v4/tests/${environmentId}" -Headers $requestHeaders

        } catch {
            [int]$statusCode = $_.Exception.Response.StatusCode;

            switch ($statusCode) {
                403 { Write-Host -Object 'Forbidden' }
                404 { Write-Host -Object 'Environment not found' }
                409 { Write-Host -Object 'Unable to start test' }
                401 { Write-Host -Object 'Unauthorized' }
                default { throw }
            }
        }
$response.state
}

function Waitfor-LoadTest{
    Write-Host "Waiting for Load test to finish"
    while ( [array]$ltStatus = (Get-ltStatus -ErrorAction SilentlyContinue ) -eq "Running" )
        {
        Write-Host 'Test is still running'
        Start-Sleep -Seconds 10
        }
}

function SendErrorMessage {

    $JSONBody = [PSCustomObject][Ordered]@{
      "@type"      = "MessageCard"
      "@context"   = "http://schema.org/extensions"
      "summary"    = "Incoming Alert Message!"
      "themeColor" = '0078D7'
      "sections"   = @(
         @{
     "activityTitle"    = "Test initiated"
            "activitySubtitle" = "By user: @Neda"
     "activityImage"    = "https://www.prchecker.info/free-icons/128x128/rocket_128_px.png"
            "facts"            = @(
     @{
     "name"  = "Test:"
     "value" = $config.PoolName
     },
     @{
     "name"  = "Users"
     "value" = "250"
     }
     )
     "markdown" = $true
     }
      )
    }

    $TeamMessageBody = ConvertTo-Json $JSONBody -Depth 100

    $parameters = @{
        "URI"         = $config.WebHoohURI
        "Method"      = 'POST'
        "Body"        = $TeamMessageBody
        "ContentType" = 'application/json'
    }

    Invoke-RestMethod @parameters | Out-Null

}

Get-CounterStatsPlus -ServersTXTfile $ServersTextFile -OutputFile $PSScriptRoot\$OutPutFile