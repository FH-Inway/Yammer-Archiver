# Update-LatestMessageId.ps1
# Updates groups-config.json with the latest message id for each group

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$configPath = Join-Path $repoRoot 'Dynamics 365 and Power Platform Preview Programs\groups-config.json'

# Load the group config
$config = Get-Content $configPath | ConvertFrom-Json

foreach ($group in $config.groups) {
    $groupName = $group.name
    Write-Host  "Processing group: $groupName"
    $groupDir = Join-Path $repoRoot "Dynamics 365 and Power Platform Preview Programs\$groupName"
    if (Test-Path $groupDir) {
        $jsonFiles = Get-ChildItem -Path $groupDir -Filter "*Messages*.json" -File
        $maxId = 0
        foreach ($jsonFile in $jsonFiles) {
            $json = Get-Content $jsonFile.FullName -Raw | ConvertFrom-Json
            if ($json.body.value) {
                $ids = $json.body.value | Where-Object { $_.id } | ForEach-Object { [int64]$_.id }
                if ($ids.Count -gt 0) {
                    $fileMax = ($ids | Measure-Object -Maximum).Maximum
                    if ($fileMax -gt $maxId) { $maxId = $fileMax }
                }
            }
        }
        $group.lastMessageId = $maxId
    } else {
        $group.lastMessageId = 0
    }
}

# Save the updated config with integer lastMessageId values
$json = $config | ConvertTo-Json -Depth 5
# Replace any ".0" at the end of lastMessageId values with nothing
$json = $json -replace '("lastMessageId"\s*:\s*)(\d+)\.0', '$1$2'
$json | Set-Content $configPath -Encoding UTF8

Write-Host "groups-config.json updated with latest message ids."
