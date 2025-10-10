# Export-GroupConversationHistory.ps1
# Creates a markdown file with the conversation history of a specific Yammer group
# Based on the logic from script.js

<#
.SYNOPSIS
    Exports Yammer group conversation history to a markdown file.

.DESCRIPTION
    This script loads messages and user references from JSON files for a specific Yammer group,
    builds the conversation hierarchy (threads and replies), and exports it to a markdown file
    with proper formatting including user names, dates, and threaded conversations.

.PARAMETER GroupName
    The name of the Yammer group to export. Must match a directory name under 
    "Dynamics 365 and Power Platform Preview Programs".

.PARAMETER OutputPath
    The path where the markdown file will be created. The directory will be created if it doesn't exist.

.EXAMPLE
    .\Export-GroupConversationHistory.ps1 -GroupName "Dev ALM" -OutputPath "C:\temp\dev-alm-history.md"
    
    Exports the conversation history for the "Dev ALM" group to the specified markdown file.

.EXAMPLE
    .\Export-GroupConversationHistory.ps1 -GroupName "Self-Service Database Movement _ DataALM" -OutputPath ".\output\database-movement.md"
    
    Exports conversation history using a group name with special characters.

.NOTES
    - The script expects JSON files in the format: "{GroupName} Messages.json" and "{GroupName} References.json"
    - Messages are organized into threads with replies properly indented
    - User names are resolved from the References file when available
    - Dates are formatted as "yyyy-MM-dd HH:mm:ss UTC"
    - HTML content is converted to plain text for markdown compatibility
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$GroupName,
    
    [Parameter(Mandatory=$true)]
    [string]$OutputPath
)

function Format-YammerDate {
    param([string]$DateString)
    
    if ([string]::IsNullOrEmpty($DateString)) {
        return "No date"
    }
    
    try {
        # Parse date format: 2025/06/11 15:12:40 +0000
        # Convert to ISO format: 2025-06-11T15:12:40+00:00
        $isoDateString = $DateString -replace '(\d{4})/(\d{2})/(\d{2})\s(\d{2}):(\d{2}):(\d{2})\s\+(\d{2})(\d{2})', '$1-$2-$3T$4:$5:$6+$7:$8'
        $dateObj = [DateTime]::Parse($isoDateString)
        return $dateObj.ToString("yyyy-MM-dd HH:mm:ss UTC")
    }
    catch {
        return $DateString
    }
}

function Get-UserName {
    param(
        [string]$UserId,
        [hashtable]$UserMap
    )
    
    if ($UserMap.ContainsKey($UserId)) {
        return $UserMap[$UserId]
    }
    return "User ID: $UserId"
}

function Get-MessageContent {
    param($Message)
    
    if ($Message.body.plain) {
        return $Message.body.plain
    }
    elseif ($Message.body.parsed) {
        return $Message.body.parsed
    }
    elseif ($Message.body.rich) {
        # Strip HTML tags for markdown and convert common HTML entities
        $content = $Message.body.rich -replace '<br\s*/?>', "`n"
        $content = $content -replace '<[^>]+>', ''
        $content = $content -replace '&lt;', '<'
        $content = $content -replace '&gt;', '>'
        $content = $content -replace '&amp;', '&'
        $content = $content -replace '&quot;', '"'
        return $content.Trim()
    }
    elseif ($Message.content_excerpt) {
        return $Message.content_excerpt
    }
    else {
        return "[No content]"
    }
}

function Write-MessageToMarkdown {
    param(
        $Message,
        [hashtable]$UserMap,
        [int]$IndentLevel = 0
    )
    
    $indent = "  " * $IndentLevel
    $userName = Get-UserName -UserId $Message.sender_id -UserMap $UserMap
    $formattedDate = Format-YammerDate -DateString $Message.created_at
    $content = Get-MessageContent -Message $Message
    
    # Escape markdown special characters in content
    $content = $content -replace '([*_`\[\]\\])', '\$1'
    
    # Format message as markdown
    $markdown = @"
$indent**Author:** $userName  
$indent**Date:** $formattedDate  
$indent**Message ID:** $($Message.id)

$indent$content

"@
    
    return $markdown
}

function Write-ThreadToMarkdown {
    param(
        $Messages,
        [hashtable]$UserMap,
        [int]$IndentLevel = 0
    )
    
    $markdown = ""
    
    foreach ($message in $Messages) {
        $markdown += Write-MessageToMarkdown -Message $message -UserMap $UserMap -IndentLevel $IndentLevel
        $markdown += "`n"
        if ($message.children -and $message.children.Count -gt 0) {
            $indent = "  " * $IndentLevel
            $markdown += "$indent**Replies ($($message.children.Count)):**`n`n"
            $markdown += Write-ThreadToMarkdown -Messages $message.children -UserMap $UserMap -IndentLevel ($IndentLevel + 1)
        }
    }
    
    return $markdown
}

# Main script logic
try {
    $repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
    $groupDir = Join-Path $repoRoot "Dynamics 365 and Power Platform Preview Programs\$GroupName"
    
    if (-not (Test-Path $groupDir)) {
        throw "Group directory not found: $groupDir"
    }
    
    # Load messages
    $messagesFile = Join-Path $groupDir "$GroupName Messages.json"
    if (-not (Test-Path $messagesFile)) {
        throw "Messages file not found: $messagesFile"
    }
    
    Write-Host "Loading messages from: $messagesFile"
    $messagesFileContentRaw = Get-Content $messagesFile -Raw
    $messagesFileContentUTF8 = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::Default.GetBytes($messagesFileContentRaw))
    $messagesData = $messagesFileContentUTF8 | ConvertFrom-Json
    $messages = $messagesData.body.value
    
    # Load user references
    $userMap = @{}
    $referencesFile = Join-Path $groupDir "$GroupName References.json"
    if (Test-Path $referencesFile) {
        Write-Host "Loading user references from: $referencesFile"
        $referencesFileContentRaw = Get-Content $referencesFile -Raw
        $referencesFileContentUTF8 = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::Default.GetBytes($referencesFileContentRaw))
        $referencesData = $referencesFileContentUTF8 | ConvertFrom-Json
        foreach ($ref in $referencesData.body.value) {
            if ($ref.id -and $ref.full_name) {
                $userMap[$ref.id.ToString()] = $ref.full_name
            }
        }
        Write-Host "Loaded $($userMap.Count) user references"
    }
    else {
        Write-Warning "References file not found: $referencesFile"
    }
    
    # Filter out invalid messages and build message map
    $validMessages = @()
    $messageMap = @{}
    
    foreach ($msg in $messages) {
        # Skip entries that don't have an ID or are just placeholders
        # Check if message only has group_created_id property (matching JavaScript logic)
        $properties = @($msg.PSObject.Properties | Where-Object { $_.Name -ne "group_created_id" })
        if (-not $msg.id -or ($properties.Count -eq 0 -and $msg.group_created_id)) {
            continue
        }
        
        $msg | Add-Member -NotePropertyName "children" -NotePropertyValue @() -Force
        $messageMap[$msg.id.ToString()] = $msg
        $validMessages += $msg
    }
    
    Write-Host "Processing $($validMessages.Count) valid messages"
    
    # Sort messages by created_at in ascending order (oldest first)
    $validMessages = $validMessages | Sort-Object { 
        try {
            $isoDateString = $_.created_at -replace '(\d{4})/(\d{2})/(\d{2})\s(\d{2}):(\d{2}):(\d{2})\s\+(\d{2})(\d{2})', '$1-$2-$3T$4:$5:$6+$7:$8'
            [DateTime]::Parse($isoDateString)
        }
        catch {
            [DateTime]::MinValue
        }
    }
    
    # Build hierarchy
    $roots = @()
    $missingParentCount = 0
    
    foreach ($msg in $validMessages) {
        if ($msg.replied_to_id) {
            $parentKey = $msg.replied_to_id.ToString()
            if ($messageMap.ContainsKey($parentKey)) {
                $messageMap[$parentKey].children += $msg
            }
            else {
                $roots += $msg
                $missingParentCount++
            }
        }
        else {
            $roots += $msg
        }
    }
    
    # Sort root messages by date (newest first) and replies by date (oldest first)
    $roots = $roots | Sort-Object { 
        try {
            $isoDateString = $_.created_at -replace '(\d{4})/(\d{2})/(\d{2})\s(\d{2}):(\d{2}):(\d{2})\s\+(\d{2})(\d{2})', '$1-$2-$3T$4:$5:$6+$7:$8'
            [DateTime]::Parse($isoDateString)
        }
        catch {
            [DateTime]::MinValue
        }
    } -Descending
    
    foreach ($msg in $messageMap.Values) {
        if ($msg.children.Count -gt 0) {
            $msg.children = $msg.children | Sort-Object { 
                try {
                    $isoDateString = $_.created_at -replace '(\d{4})/(\d{2})/(\d{2})\s(\d{2}):(\d{2}):(\d{2})\s\+(\d{2})(\d{2})', '$1-$2-$3T$4:$5:$6+$7:$8'
                    [DateTime]::Parse($isoDateString)
                }
                catch {
                    [DateTime]::MinValue
                }
            }
        }
    }
    
    Write-Host "Found $($roots.Count) root threads with $($validMessages.Count) total messages"
    if ($missingParentCount -gt 0) {
        Write-Host "Note: $missingParentCount messages have missing parents and are treated as root messages"
    }
    
    # Generate markdown content
    $markdown = @"
# $GroupName - Conversation History

Generated on: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")  
Total threads: $($roots.Count)  
Total messages: $($validMessages.Count)  
$(if ($missingParentCount -gt 0) { "Messages with missing parents: $missingParentCount" })

---

"@
    
    # Add each thread to markdown
    for ($i = 0; $i -lt $roots.Count; $i++) {
        if ($i % 100 -eq 0) {
            Write-Progress -Activity "Processing Threads" -Status "Thread $($i + 1) of $($roots.Count)" -CurrentOperation "Writing thread $($i + 1)"
        }

        $thread = $roots[$i]
        $threadNumber = $i + 1
        
        $markdown += "## Thread $threadNumber`n`n"
        $markdown += Write-ThreadToMarkdown -Messages @($thread) -UserMap $userMap
        $markdown += "`n---`n`n"
    }
    
    # Ensure output directory exists
    $outputDir = Split-Path -Parent $OutputPath
    if ($outputDir -and -not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }
    
    # Write to file
    # $markdown | Set-Content -Path $OutputPath -Encoding UTF8
    New-Item -Path $OutputPath -ItemType File -Force -Value $markdown
    
    Write-Host "Successfully exported conversation history to: $OutputPath"
    Write-Host "File size: $((Get-Item $OutputPath).Length) bytes"
    
}
catch {
    Write-Error "Error: $($_.Exception.Message)"
    exit 1
}