# Export-ThreadConversation.ps1
# Creates a markdown file with the conversation history of a specific Yammer thread

<#
.SYNOPSIS
    Exports a single Yammer thread conversation to a markdown file.

.DESCRIPTION
    This script loads a thread JSON file, builds the conversation hierarchy (main message and replies), 
    and exports it to a markdown file with proper formatting including user names, dates, and threaded conversations.
    All user information and message references are extracted from the thread JSON file itself.

.PARAMETER ThreadJsonPath
    The path to the thread JSON file (e.g., "thread_3049071028559872.json").

.PARAMETER OutputPath
    The path where the markdown file will be created. The directory will be created if it doesn't exist.

.EXAMPLE
    .\Export-ThreadConversation.ps1 -ThreadJsonPath ".\thread_3049071028559872.json" -OutputPath ".\output\thread-discussion.md"
    
    Exports the thread conversation to the specified markdown file.

.NOTES
    - The script expects the thread JSON to have a "messages" array and "references" section
    - User names and message references are resolved from the thread JSON file
    - Dates are formatted as "yyyy-MM-dd HH:mm:ss UTC"
    - HTML content is converted to plain text for markdown compatibility
    - All messages from both the messages array and references are included
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$ThreadJsonPath,
    
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
        [int]$MessageNumber
    )
    
    $userName = Get-UserName -UserId $Message.sender_id -UserMap $UserMap
    $formattedDate = Format-YammerDate -DateString $Message.created_at
    $content = Get-MessageContent -Message $Message
    
    # Escape markdown special characters in content
    $content = $content -replace '([*_`\[\]\\])', '\$1'
    
    # Add attachment information if present
    $attachmentInfo = ""
    if ($Message.attachments -and $Message.attachments.Count -gt 0) {
        $attachmentInfo = "`n**Attachments:** $($Message.attachments.Count) file(s)"
    }
    
    # Format message as markdown
    $markdown = @"
## Message $MessageNumber

**Author:** $userName  
**Date:** $formattedDate  
**Message ID:** $($Message.id)$attachmentInfo

$content

***

"@
    
    return $markdown
}

# Main script logic
try {
    # Validate input files exist
    if (-not (Test-Path $ThreadJsonPath)) {
        throw "Thread JSON file not found: $ThreadJsonPath"
    }
    
    # Load thread data
    Write-Host "Loading thread from: $ThreadJsonPath"
    $threadFileContentRaw = Get-Content $ThreadJsonPath -Raw
    $threadFileContentUTF8 = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::Default.GetBytes($threadFileContentRaw))
    $threadData = $threadFileContentUTF8 | ConvertFrom-Json
    $messages = $threadData.messages
    
    # Load references from the thread data itself
    $userMap = @{}
    $referenceMessages = @()
    
    if ($threadData.references) {
        foreach ($ref in $threadData.references) {
            if ($ref.type -eq "user" -and $ref.id -and $ref.full_name) {
                $userMap[$ref.id.ToString()] = $ref.full_name
            }
            elseif ($ref.type -eq "message" -and $ref.id) {
                # Add all message references to our collection
                $referenceMessages += $ref
                Write-Host "Found message reference: ID $($ref.id)"
            }
            # Handle messages without explicit type field but with message properties
            elseif (-not $ref.type -and $ref.id -and ($ref.body -or $ref.content_excerpt -or $ref.sender_id)) {
                $referenceMessages += $ref
                Write-Host "Found message in references without type: ID $($ref.id)"
            }
        }
    }
    Write-Host "Loaded $($userMap.Count) user references and $($referenceMessages.Count) message references"
    
    # Filter valid messages and build message map
    $validMessages = @()
    $messageMap = @{
    }
    
    # Add all reference messages first (including the thread starter and any others)
    foreach ($refMsg in $referenceMessages) {
        if ($refMsg.id) {
            $refMsg | Add-Member -NotePropertyName "children" -NotePropertyValue @() -Force
            $messageMap[$refMsg.id.ToString()] = $refMsg
            $validMessages += $refMsg
            Write-Host "Added reference message: ID $($refMsg.id)"
        }
    }
    Write-Host "Added $($referenceMessages.Count) messages from references"
    
    # Add messages from the main messages array
    $addedFromMessages = 0
    foreach ($msg in $messages) {
        if (-not $msg.id) {
            continue
        }
        
        # Check if message already exists from references
        if ($messageMap.ContainsKey($msg.id.ToString())) {
            $existingMsg = $messageMap[$msg.id.ToString()]
            
            # Compare message content to determine if it's truly a duplicate
            $existingContent = ""
            $newContent = ""
            
            # Get content from existing message
            if ($existingMsg.body.plain) {
                $existingContent = $existingMsg.body.plain
            } elseif ($existingMsg.body.parsed) {
                $existingContent = $existingMsg.body.parsed
            } elseif ($existingMsg.content_excerpt) {
                $existingContent = $existingMsg.content_excerpt
            }
            
            # Get content from new message
            if ($msg.body.plain) {
                $newContent = $msg.body.plain
            } elseif ($msg.body.parsed) {
                $newContent = $msg.body.parsed
            } elseif ($msg.content_excerpt) {
                $newContent = $msg.content_excerpt
            }
            
            # If content is significantly different or existing message has minimal data, replace it
            $existingHasFullData = ($existingMsg.body -and ($existingMsg.body.plain -or $existingMsg.body.parsed -or $existingMsg.body.rich))
            $newHasFullData = ($msg.body -and ($msg.body.plain -or $msg.body.parsed -or $msg.body.rich))
            
            if (-not $existingHasFullData -and $newHasFullData) {
                # Replace reference stub with full message
                Write-Host "Replacing reference stub with full message: ID $($msg.id)"
                $msg | Add-Member -NotePropertyName "children" -NotePropertyValue $existingMsg.children -Force
                $messageMap[$msg.id.ToString()] = $msg
                
                # Update the message in validMessages array
                for ($i = 0; $i -lt $validMessages.Count; $i++) {
                    if ($validMessages[$i].id -eq $msg.id) {
                        $validMessages[$i] = $msg
                        break
                    }
                }
                $addedFromMessages++
            }
            elseif ($existingContent -ne $newContent -and $newContent.Length -gt $existingContent.Length) {
                # Content is different and new message has more content, replace it
                Write-Host "Replacing message with more complete content: ID $($msg.id)"
                $msg | Add-Member -NotePropertyName "children" -NotePropertyValue $existingMsg.children -Force
                $messageMap[$msg.id.ToString()] = $msg
                
                # Update the message in validMessages array
                for ($i = 0; $i -lt $validMessages.Count; $i++) {
                    if ($validMessages[$i].id -eq $msg.id) {
                        $validMessages[$i] = $msg
                        break
                    }
                }
                $addedFromMessages++
            }
            else {
                Write-Host "Skipping duplicate message from main array: ID $($msg.id)"
            }
        } 
        else {
            # New message, add it
            $msg | Add-Member -NotePropertyName "children" -NotePropertyValue @() -Force
            $messageMap[$msg.id.ToString()] = $msg
            $validMessages += $msg
            $addedFromMessages++
        }
    }
    Write-Host "Added/updated $addedFromMessages messages from main messages array"
    
    Write-Host "Processing $($validMessages.Count) total valid messages"
    
    # Debug: List all message IDs
    Write-Host "All message IDs found:"
    foreach ($msg in $validMessages | Sort-Object id) {
        $sender = Get-UserName -UserId $msg.sender_id -UserMap $userMap
        $date = Format-YammerDate -DateString $msg.created_at
        Write-Host "  ID: $($msg.id), Sender: $sender, Date: $date"
    }
    
    # Sort all messages by created_at in ascending order (oldest first)
    $validMessages = $validMessages | Sort-Object { 
        try {
            $isoDateString = $_.created_at -replace '(\d{4})/(\d{2})/(\d{2})\s(\d{2}):(\d{2}):(\d{2})\s\+(\d{2})(\d{2})', '$1-$2-$3T$4:$5:$6+$7:$8'
            [DateTime]::Parse($isoDateString)
        }
        catch {
            [DateTime]::MinValue
        }
    }
    
    # Extract thread info from the first (oldest) message
    $firstMessage = $validMessages[0]
    $threadIdForDisplay = if ($threadId) { $threadId } else { $firstMessage.id }
    $threadStartDate = if ($firstMessage) { Format-YammerDate -DateString $firstMessage.created_at } else { "Unknown" }
    $originalAuthor = if ($firstMessage) { Get-UserName -UserId $firstMessage.sender_id -UserMap $userMap } else { "Unknown" }
    
    Write-Host "Found thread with $($validMessages.Count) total messages"
    Write-Host "Thread starter: ID $threadIdForDisplay by $originalAuthor"
    
    # Generate markdown content
    $markdown = @"
***
title: "Thread Discussion - $threadIdForDisplay"
date: $(Get-Date -Format "yyyy-MM-dd")
thread_id: $threadIdForDisplay
original_author: $originalAuthor
thread_start_date: $threadStartDate
***

# Thread Discussion

**Thread ID:** $threadIdForDisplay  
**Original Author:** $originalAuthor  
**Thread Started:** $threadStartDate  
**Total Messages:** $($validMessages.Count)  
**Generated on:** $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

***

"@
    
    # Add all messages in chronological order
    for ($i = 0; $i -lt $validMessages.Count; $i++) {
        $message = $validMessages[$i]
        $messageNumber = $i + 1
        $markdown += Write-MessageToMarkdown -Message $message -UserMap $userMap -MessageNumber $messageNumber
    }
    
    # Ensure output directory exists
    $outputDir = Split-Path -Parent $OutputPath
    if ($outputDir -and -not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }
    
    # Write to file
    New-Item -Path $OutputPath -ItemType File -Force -Value $markdown
    
    Write-Host "Successfully exported thread conversation to: $OutputPath"
    Write-Host "File size: $((Get-Item $OutputPath).Length) bytes"
    
}
catch {
    Write-Error "Error: $($_.Exception.Message)"
    exit 1
}
