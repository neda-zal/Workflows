$config = Get-Content -Path "$PSScriptRoot\config.json" -Raw | ConvertFrom-Json


$ApiSecret = $config.PublicAPISecret
$TestID = $config.TestID 


$baseUrl = $config.BaseApplianceURL + "/publicApi"

function Start-LoadTest{
[string]$publicApiSecret = $ApiSecret

$environmentId = $testID
$requestObject = @{
    # Comment
    'comment' = '...'
}

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { return $true; }

$requestHeaders = @{ 'Content-Type' = 'application/json' }

$authRequest = ConvertTo-Json -InputObject @{ secret = $publicApiSecret }

$authResponse = Invoke-RestMethod -Method Post -Uri "${baseUrl}/v2/authentication/token" -Headers $requestHeaders -Body $authRequest

$requestHeaders.Add('Authorization', "Bearer $($authResponse.accessToken)")

$requestBody = ConvertTo-Json -InputObject $requestObject

    try {
        $response = Invoke-RestMethod -Method Put -Uri "${baseUrl}/v2/environments/${environmentId}/start" -Headers $requestHeaders -Body $authRequest
        Write-Host -Object 'Load Test Started'

    } catch {

        Write-Host $_.Exception
        [int]$statusCode = $_.Exception.Response.StatusCode;

        switch ($statusCode) {
            403 { Write-Host -Object 'Forbidden' }
            404 { Write-Host -Object 'Environment not found' }
            409 { Write-Host -Object 'Unable to start test' }
            401 { Write-Host -Object 'Unauthorized' }
            default { throw }
        }
    }
}

Start-LoadTest

Start-Sleep -Seconds 10

