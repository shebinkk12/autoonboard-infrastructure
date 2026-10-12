$payload = @{
    name = "Rohan Verma"
    email = "rohan.v@shoufilavgmail.onmicrosoft.com"
    department = "Engineering"
} | ConvertTo-Json

try {
    $res = Invoke-RestMethod -Uri "https://func-autoonboard-2026.azurewebsites.net/api/ProcessOnboarding" -Method Post -ContentType "application/json" -Body $payload
    Write-Host "SUCCESS:"
    $res | ConvertTo-Json -Depth 5
} catch {
    Write-Host "BACKEND RESPONSE:"
    if ($_.Exception.Response) {
        $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
        $reader.ReadToEnd()
    } else {
        $_
    }
}
