# Merge-GroupJsonFiles.ps1
# Merges numbered Messages/References JSON files into the main Messages/References JSON file for each group

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$groupsConfigPath = Join-Path $repoRoot 'Dynamics 365 and Power Platform Preview Programs\groups-config.json'

# Load group config
$config = Get-Content $groupsConfigPath | ConvertFrom-Json

# Replace foreach with for loops for groups
for ($g = 0; $g -lt $config.groups.Count; $g++) {
    $group = $config.groups[$g]
    $groupName = $group.name
    Write-Host "Processing group: $groupName"
    $groupDir = Join-Path $repoRoot "Dynamics 365 and Power Platform Preview Programs\$groupName"
    if (-not (Test-Path $groupDir)) { Write-Host "  Group directory not found: $groupDir"; continue }

    $types = @('Messages', 'References')
    for ($t = 0; $t -lt $types.Count; $t++) {
        $type = $types[$t]
        $mainFile = Join-Path $groupDir ("$groupName $type.json")
        # Improved regex: match files with a space, then digits, then .json (e.g. Messages 123456.json)
        $numberedFiles = Get-ChildItem -Path $groupDir -Filter "$groupName $type *.json" -File | Where-Object { $_.Name -match "^$([regex]::Escape($groupName)) $type \d+\.json$" }
        Write-Host "    Numbered files matching regex: $($numberedFiles.Count)"
        if ($numberedFiles.Count -eq 0) { continue }

        # Load main file if it exists, else create a new structure
        $mainJson = @{ body = @{ value = @() } }
        if (Test-Path $mainFile) {
            $mainJson = Get-Content $mainFile -Raw | ConvertFrom-Json
            if (-not $mainJson.body.value) { $mainJson.body.value = @() }
        }

        # Merge all numbered files
        $json = $null
        for ($f = 0; $f -lt $numberedFiles.Count; $f++) {
            $file = $numberedFiles[$f]
            $json = Get-Content $file.FullName -Raw | ConvertFrom-Json
            if ($json) {
                $mainJson.body.value += $json
            } 
        }

        $beforeCount = $mainJson.body.value.Count
        if ($type -eq 'Messages') {
            $mainJson.body.value = $mainJson.body.value | Sort-Object id -Descending
        }
        elseif ($type -eq 'References') {
            $mainJson.body.value = $mainJson.body.value | Sort-Object type,id -Descending
        }
        $afterCount = $mainJson.body.value.Count

        # Save merged file
        $mainJson | ConvertTo-Json -Depth 10 | Set-Content $mainFile -Encoding UTF8

        # Optionally, remove numbered files after merging
        $numberedFiles | Remove-Item
        Write-Host "    Removed $($numberedFiles.Count) numbered $type files."
        Write-Host "  Merged $($numberedFiles.Count) $type files for group '$groupName' into $mainFile."
    }
}

Write-Host "Merging complete."
