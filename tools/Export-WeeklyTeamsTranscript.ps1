[CmdletBinding()]
param(
    [string]$ConfigPath,
    [string]$ExpectedAccount = "",
    [ValidateRange(1, 30)]
    [int]$DaysBack = 7,
    [string]$OutputDirectory = "",
    [string]$TimeZoneId = "",
    [string[]]$ExcludedChatIds = @(),
    [string[]]$ExcludedEmailDomains = @(),
    [string[]]$ExcludedTopicPatterns = @(),
    [ValidateRange(1, 20)]
    [int]$MaxRetries = 8
)
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"


if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Join-Path (Split-Path -Parent $PSScriptRoot) "config.psd1"
}


$configData = @{}
if (Test-Path -LiteralPath $ConfigPath) { $configData = Import-PowerShellDataFile -LiteralPath $ConfigPath }


foreach ($settingName in @(
    "ExpectedAccount",
    "DaysBack",
    "OutputDirectory",
    "TimeZoneId",
    "ExcludedChatIds",
    "ExcludedEmailDomains",
    "ExcludedTopicPatterns",
    "MaxRetries"
)) {
    if (-not $PSBoundParameters.ContainsKey($settingName) -and $configData.ContainsKey($settingName)) {
        Set-Variable -Name $settingName -Value $configData[$settingName] -Scope Script
    }
}


if ($DaysBack -lt 1 -or $DaysBack -gt 30) { throw "DaysBack must be between 1 and 30." }
if ($MaxRetries -lt 1 -or $MaxRetries -gt 20) { throw "MaxRetries must be between 1 and 20." }


$script:GraphRequestCount = 0
$script:RetryWaits = [System.Collections.Generic.List[int]]::new()
$script:ExpectedUserId = $null
$script:ExpectedUserDisplayName = $null
$script:UnresolvedReactionCount = 0
$script:ExpandedReplyReferenceCount = 0
$script:IncompleteReplyReferenceCount = 0
$script:ReplyAttachmentContentMissingCount = 0
$script:ReplyAttachmentJsonErrorCount = 0
$script:ReplySourceBodyMissingCount = 0
$script:ReplySenderNameUnavailableCount = 0
$script:ReplyTimestampUnavailableCount = 0
$script:FetchedReplySourceMessageCount = 0
$script:KnownIdentityNames = [System.Collections.Generic.Dictionary[string,string]]::new([StringComparer]::OrdinalIgnoreCase)
$script:MeetingChatsWithoutRoster = [System.Collections.Generic.List[string]]::new()


function Get-ReportingTimeZone {
    if ([string]::IsNullOrWhiteSpace($TimeZoneId)) { return [System.TimeZoneInfo]::Local }
    try { return [System.TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneId) }
    catch { throw "The configured TimeZoneId '$TimeZoneId' is not available on this computer." }
}
function Get-OptionalPropertyValue {
    param(
        $InputObject,
        [Parameter(Mandatory)]
        [string]$Name
    )


    if ($null -eq $InputObject) {
        return $null
    }


    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }


    return $property.Value
}


function Resolve-OutputDirectory {
    param([string]$RequestedDirectory)
    if (-not [string]::IsNullOrWhiteSpace($RequestedDirectory)) { return [System.IO.Path]::GetFullPath($RequestedDirectory) }
    $documentsRoot = [Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments)
    if ([string]::IsNullOrWhiteSpace($documentsRoot)) { $documentsRoot = Join-Path $env:USERPROFILE "Documents" }
    if ([string]::IsNullOrWhiteSpace($documentsRoot)) { throw "The Documents folder could not be resolved. Set OutputDirectory in config.psd1." }
    return (Join-Path $documentsRoot "Teams Transcripts")
}
function Get-HttpStatusCode {
    param($CaughtError)


    try {
        if ($null -ne $CaughtError.Exception.Response.StatusCode) {
            return [int]$CaughtError.Exception.Response.StatusCode
        }
    }
    catch {}


    try {
        if ($null -ne $CaughtError.Exception.ResponseStatusCode) {
            return [int]$CaughtError.Exception.ResponseStatusCode
        }
    }
    catch {}


    if ($CaughtError.Exception.Message -match "\b(429|500|502|503|504)\b") {
        return [int]$Matches[1]
    }


    return $null
}


function Get-RetryAfterSeconds {
    param(
        $CaughtError,
        [int]$Attempt
    )


    try {
        $retryAfter = $CaughtError.Exception.Response.Headers.RetryAfter
        if ($null -ne $retryAfter) {
            if ($null -ne $retryAfter.Delta) {
                return [Math]::Max(1, [int][Math]::Ceiling($retryAfter.Delta.TotalSeconds))
            }
            if ($null -ne $retryAfter.Date) {
                $seconds = ($retryAfter.Date - [DateTimeOffset]::UtcNow).TotalSeconds
                return [Math]::Max(1, [int][Math]::Ceiling($seconds))
            }
        }
    }
    catch {}


    if ($CaughtError.Exception.Message -match "(?i)retry.?after[^0-9]*(\d+)") {
        return [Math]::Max(1, [int]$Matches[1])
    }


    return [Math]::Min(65, [int][Math]::Pow(2, [Math]::Min($Attempt, 6)))
}


function Invoke-GraphGetWithRetry {
    param(
        [Parameter(Mandatory)]
        [string]$Uri
    )


    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            $script:GraphRequestCount++
            return Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject
        }
        catch {
            $statusCode = Get-HttpStatusCode -CaughtError $_
            $isRetriable = $statusCode -in @(429, 500, 502, 503, 504)


            if (-not $isRetriable -or $attempt -eq $MaxRetries) {
                $statusText = if ($null -eq $statusCode) { "unknown status" } else { "HTTP $statusCode" }
                throw "Microsoft Graph GET failed for '$Uri' with $statusText. $($_.Exception.Message)"
            }


            $waitSeconds = Get-RetryAfterSeconds -CaughtError $_ -Attempt $attempt
            $script:RetryWaits.Add($waitSeconds)
            Write-Warning "Microsoft Graph returned HTTP $statusCode. Waiting $waitSeconds seconds before retry $($attempt + 1) of $MaxRetries."
            Start-Sleep -Seconds $waitSeconds
        }
    }
}


function Get-AllGraphCollectionItems {
    param(
        [Parameter(Mandatory)]
        [string]$InitialUri
    )


    $items = [System.Collections.Generic.List[object]]::new()
    $nextUri = $InitialUri


    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        $response = Invoke-GraphGetWithRetry -Uri $nextUri
        foreach ($item in @($response.value)) {
            $items.Add($item)
        }
        $nextLinkProperty = $response.PSObject.Properties['@odata.nextLink']
        if ($null -eq $nextLinkProperty) {
            $nextUri = $null
        }
        else {
            $nextUri = [string]$nextLinkProperty.Value
        }
    }


    return $items
}


function Test-IsMeetingChat {
    param($Chat)


    $chatType = [string](Get-OptionalPropertyValue -InputObject $Chat -Name "chatType")
    if ([string]::Equals($chatType, "meeting", [StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }


    # Some /me/chats responses omit or inconsistently deserialize chatType.
    # Also recognize the 19:meeting_ form returned for Teams meeting chats.
    $chatId = [string](Get-OptionalPropertyValue -InputObject $Chat -Name "id")
    return $chatId -match "(?i)^19(?::|%3A)meeting_"
}


function Get-ChatMembers {
    param($Chat)


    $expandedMembers = Get-OptionalPropertyValue -InputObject $Chat -Name "members"
    if ($null -ne $expandedMembers -and @($expandedMembers).Count -gt 0) {
        return @($expandedMembers)
    }


    # A meeting can remain in /me/chats even when Graph denies its separate
    # roster endpoint. The roster is metadata, not message content, and is not
    # needed for optional 1:1 domain exclusions. Include such meetings and
    # still require their message endpoint to succeed.
    if (Test-IsMeetingChat -Chat $Chat) {
        $script:MeetingChatsWithoutRoster.Add([string]$Chat.id)
        return @()
    }


    $encodedChatId = [Uri]::EscapeDataString([string]$Chat.id)
    # This endpoint does not support OData query parameters; adding $top returns HTTP 400.
    $uri = "https://graph.microsoft.com/v1.0/me/chats/$encodedChatId/members"
    return @(Get-AllGraphCollectionItems -InitialUri $uri)
}


function Test-IsCurrentUserMember {
    param(
        $Member,
        [string]$Account
    )


    $memberUserId = [string](Get-OptionalPropertyValue -InputObject $Member -Name "userId")
    if (-not [string]::IsNullOrWhiteSpace($script:ExpectedUserId) -and $memberUserId.Equals($script:ExpectedUserId, [StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }


    $memberEmail = [string](Get-OptionalPropertyValue -InputObject $Member -Name "email")
    if (-not [string]::IsNullOrWhiteSpace($memberEmail) -and $memberEmail.Equals($Account, [StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }


    return $false
}


function Test-IsExcludedChat {
    param($Chat, [object[]]$Members)
    $otherMembers = @($Members | Where-Object { -not (Test-IsCurrentUserMember -Member $_ -Account $ExpectedAccount) })
    if ($Chat.chatType -eq "oneOnOne" -and $otherMembers.Count -eq 1) {
        $otherEmail = [string]$otherMembers[0].email
        foreach ($domain in @($ExcludedEmailDomains)) {
            $cleanDomain = ([string]$domain).Trim().TrimStart("@")
            if (-not [string]::IsNullOrWhiteSpace($cleanDomain) -and $otherEmail.EndsWith("@$cleanDomain", [StringComparison]::OrdinalIgnoreCase)) { return $true }
        }
    }
    $topic = [string]$Chat.topic
    foreach ($pattern in @($ExcludedTopicPatterns)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$pattern) -and $topic -match [string]$pattern) { return $true }
    }
    return $false
}
function Get-ChatLabel {
    param(
        $Chat,
        [object[]]$Members
    )


    $topic = ([string]$Chat.topic).Trim()
    if (-not [string]::IsNullOrWhiteSpace($topic)) {
        return ($topic -replace "[\r\n]+", " ")
    }


    $otherNames = @(
        $Members |
            Where-Object { -not (Test-IsCurrentUserMember -Member $_ -Account $ExpectedAccount) } |
            ForEach-Object { ([string]$_.displayName).Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -Unique
    )


    if ($otherNames.Count -gt 0) {
        return ($otherNames -join ", ")
    }


    if (Test-IsMeetingChat -Chat $Chat) {
        $createdDateTime = [string](Get-OptionalPropertyValue -InputObject $Chat -Name "createdDateTime")
        if (-not [string]::IsNullOrWhiteSpace($createdDateTime)) {
            try {
                return "Meeting chat ($([DateTimeOffset]::Parse($createdDateTime).ToString('yyyy-MM-dd')))"
            }
            catch {}
        }
        return "Meeting chat"
    }


    return "Unnamed Teams chat"
}


function Test-ChatMayHaveMessagesInWindow {
    param(
        $Chat,
        [Parameter(Mandatory)]
        [DateTimeOffset]$StartUtc
    )


    $previewProperty = $Chat.PSObject.Properties["lastMessagePreview"]
    if ($null -eq $previewProperty) {
        # If Graph did not return the requested expansion, do not assume the
        # chat is inactive. Query its messages and retain the hard-stop rule.
        return $true
    }


    $preview = $previewProperty.Value
    if ($null -eq $preview) {
        # Microsoft documents null as meaning that no messages were sent.
        return $false
    }


    $lastMessageDateTime = [string](Get-OptionalPropertyValue -InputObject $preview -Name "createdDateTime")
    if ([string]::IsNullOrWhiteSpace($lastMessageDateTime)) {
        return $true
    }


    try {
        return [DateTimeOffset]::Parse($lastMessageDateTime).ToUniversalTime() -ge $StartUtc
    }
    catch {
        return $true
    }
}


function Get-ChatMessagesInWindow {
    param(
        [Parameter(Mandatory)]
        [string]$ChatId,
        [Parameter(Mandatory)]
        [DateTimeOffset]$StartUtc,
        [Parameter(Mandatory)]
        [DateTimeOffset]$EndUtc
    )


    $messages = [System.Collections.Generic.List[object]]::new()
    $encodedChatId = [Uri]::EscapeDataString($ChatId)
    $endIso = $EndUtc.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    $filter = [Uri]::EscapeDataString("createdDateTime lt $endIso")
    $nextUri = "https://graph.microsoft.com/v1.0/chats/$encodedChatId/messages?`$top=50&`$orderby=createdDateTime%20desc&`$filter=$filter"
    $reachedStart = $false


    while (-not [string]::IsNullOrWhiteSpace($nextUri) -and -not $reachedStart) {
        $response = Invoke-GraphGetWithRetry -Uri $nextUri
        $pageMessages = @($response.value)


        foreach ($message in $pageMessages) {
            if ([string]::IsNullOrWhiteSpace([string]$message.createdDateTime)) {
                continue
            }


            $messageType = [string](Get-OptionalPropertyValue -InputObject $message -Name "messageType")
            if (-not [string]::IsNullOrWhiteSpace($messageType) -and -not [string]::Equals($messageType, "message", [StringComparison]::OrdinalIgnoreCase)) {
                continue
            }


            $createdUtc = [DateTimeOffset]::Parse([string]$message.createdDateTime).ToUniversalTime()
            if ($createdUtc -ge $StartUtc -and $createdUtc -lt $EndUtc) {
                $messages.Add($message)
            }
        }


        if ($pageMessages.Count -gt 0) {
            $datedMessages = @($pageMessages | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.createdDateTime) })
            if ($datedMessages.Count -gt 0) {
                $oldestOnPage = $datedMessages |
                    ForEach-Object { [DateTimeOffset]::Parse([string]$_.createdDateTime).ToUniversalTime() } |
                    Sort-Object |
                    Select-Object -First 1
                if ($oldestOnPage -lt $StartUtc) {
                    $reachedStart = $true
                }
            }
        }


        if (-not $reachedStart) {
            $nextLinkProperty = $response.PSObject.Properties['@odata.nextLink']
            if ($null -eq $nextLinkProperty) {
                $nextUri = $null
            }
            else {
                $nextUri = [string]$nextLinkProperty.Value
            }
        }
    }


    return $messages
}


function Get-MessageSender {
    param($Message)


    $from = Get-OptionalPropertyValue -InputObject $Message -Name "from"
    if ($null -ne $from) {
        $fromUser = Get-OptionalPropertyValue -InputObject $from -Name "user"
        if ($null -ne $fromUser -and -not [string]::IsNullOrWhiteSpace([string]$fromUser.displayName)) {
            return [string]$fromUser.displayName
        }
        $fromApplication = Get-OptionalPropertyValue -InputObject $from -Name "application"
        if ($null -ne $fromApplication -and -not [string]::IsNullOrWhiteSpace([string]$fromApplication.displayName)) {
            return [string]$fromApplication.displayName
        }
        $fromDevice = Get-OptionalPropertyValue -InputObject $from -Name "device"
        if ($null -ne $fromDevice -and -not [string]::IsNullOrWhiteSpace([string]$fromDevice.displayName)) {
            return [string]$fromDevice.displayName
        }
    }


    return "System"
}


function Get-IdentityIdCandidates {
    param([string]$IdentityId)


    if ([string]::IsNullOrWhiteSpace($IdentityId)) {
        return @()
    }


    $candidates = [System.Collections.Generic.List[string]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $trimmedId = $IdentityId.Trim()
    if ($seen.Add($trimmedId)) {
        $candidates.Add($trimmedId)
    }


    $guidMatch = [Regex]::Match($trimmedId, '(?i)([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})$')
    if ($guidMatch.Success -and $seen.Add($guidMatch.Groups[1].Value)) {
        $candidates.Add($guidMatch.Groups[1].Value)
    }


    return $candidates
}


function Add-IdentityToMap {
    param(        [System.Collections.Generic.Dictionary[string,string]]$Map,
        $Identity
    )


    if ($null -eq $Identity) {
        return
    }


    $identityId = [string](Get-OptionalPropertyValue -InputObject $Identity -Name "id")
    $displayName = [string](Get-OptionalPropertyValue -InputObject $Identity -Name "displayName")
    if (-not [string]::IsNullOrWhiteSpace($identityId) -and -not [string]::IsNullOrWhiteSpace($displayName)) {
        foreach ($candidateId in @(Get-IdentityIdCandidates -IdentityId $identityId)) {
            $Map[$candidateId] = $displayName.Trim()
            $script:KnownIdentityNames[$candidateId] = $displayName.Trim()
        }
    }
}


function New-ChatIdentityMap {
    param(
        [object[]]$Members,
        [object[]]$Messages
    )


    $map = [System.Collections.Generic.Dictionary[string,string]]::new([StringComparer]::OrdinalIgnoreCase)
    if (-not [string]::IsNullOrWhiteSpace($script:ExpectedUserId) -and -not [string]::IsNullOrWhiteSpace($script:ExpectedUserDisplayName)) {
        $map[$script:ExpectedUserId] = $script:ExpectedUserDisplayName
    }


    foreach ($member in @($Members)) {
        $displayName = [string](Get-OptionalPropertyValue -InputObject $member -Name "displayName")
        if ([string]::IsNullOrWhiteSpace($displayName)) {
            continue
        }
        foreach ($memberIdProperty in @("userId", "id")) {
            $memberId = [string](Get-OptionalPropertyValue -InputObject $member -Name $memberIdProperty)
            if (-not [string]::IsNullOrWhiteSpace($memberId)) {
                foreach ($candidateId in @(Get-IdentityIdCandidates -IdentityId $memberId)) {
                    $map[$candidateId] = $displayName.Trim()
                    $script:KnownIdentityNames[$candidateId] = $displayName.Trim()
                }
            }
        }
    }


    foreach ($message in @($Messages)) {
        $from = Get-OptionalPropertyValue -InputObject $message -Name "from"
        if ($null -eq $from) {
            continue
        }
        Add-IdentityToMap -Map $map -Identity (Get-OptionalPropertyValue -InputObject $from -Name "user")
        Add-IdentityToMap -Map $map -Identity (Get-OptionalPropertyValue -InputObject $from -Name "application")
        Add-IdentityToMap -Map $map -Identity (Get-OptionalPropertyValue -InputObject $from -Name "device")
    }


    return ,$map
}


function Resolve-IdentityDisplayName {
    param(
        $Identity,
        [System.Collections.Generic.Dictionary[string,string]]$IdentityMap
    )


    if ($null -eq $Identity) {
        return $null
    }


    $displayName = [string](Get-OptionalPropertyValue -InputObject $Identity -Name "displayName")
    if (-not [string]::IsNullOrWhiteSpace($displayName)) {
        return $displayName.Trim()
    }


    $identityId = [string](Get-OptionalPropertyValue -InputObject $Identity -Name "id")
    foreach ($candidateId in @(Get-IdentityIdCandidates -IdentityId $identityId)) {
        if ($null -ne $IdentityMap -and $IdentityMap.ContainsKey($candidateId)) {
            return $IdentityMap[$candidateId]
        }
        if ($script:KnownIdentityNames.ContainsKey($candidateId)) {
            return $script:KnownIdentityNames[$candidateId]
        }
    }


    return $null
}


function Resolve-ChatLabelFromMessages {
    param(
        [string]$CurrentLabel,
        [object[]]$Messages,
        [System.Collections.Generic.Dictionary[string,string]]$IdentityMap
    )


    if (-not [string]::Equals($CurrentLabel, "Unnamed Teams chat", [StringComparison]::OrdinalIgnoreCase)) {
        return $CurrentLabel
    }


    $names = [System.Collections.Generic.List[string]]::new()
    $seenNames = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($message in @($Messages)) {
        $from = Get-OptionalPropertyValue -InputObject $message -Name "from"
        $fromUser = Get-OptionalPropertyValue -InputObject $from -Name "user"
        if ($null -eq $fromUser) {
            continue
        }


        $senderId = [string](Get-OptionalPropertyValue -InputObject $fromUser -Name "id")
        if (-not [string]::IsNullOrWhiteSpace($senderId) -and [string]::Equals($senderId, $script:ExpectedUserId, [StringComparison]::OrdinalIgnoreCase)) {
            continue
        }


        $senderName = Resolve-IdentityDisplayName -Identity $fromUser -IdentityMap $IdentityMap
        if (-not [string]::IsNullOrWhiteSpace($senderName) -and $seenNames.Add($senderName)) {
            $names.Add($senderName)
        }
    }


    if ($names.Count -gt 0) {
        return ($names -join ", ")
    }


    return $CurrentLabel
}


function Convert-HtmlInlineToText {
    param([string]$Html)


    if ([string]::IsNullOrWhiteSpace($Html)) {
        return ""
    }


    $value = [Regex]::Replace($Html, "<br\s*/?>", " ", [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $value = [Regex]::Replace($value, "<[^>]+>", " ", [Text.RegularExpressions.RegexOptions]::Singleline)
    $value = [Net.WebUtility]::HtmlDecode($value).Replace([char]0x00A0, " ")
    return ([Regex]::Replace($value, "\s+", " ")).Trim()
}


function Convert-HtmlTableToMarkdown {
    param([Text.RegularExpressions.Match]$TableMatch)


    $options = [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::Singleline
    $rows = [System.Collections.Generic.List[object]]::new()
    $maximumColumns = 0


    foreach ($rowMatch in [Regex]::Matches($TableMatch.Value, "<tr\b[^>]*>(.*?)</tr>", $options)) {
        $cells = [System.Collections.Generic.List[string]]::new()
        foreach ($cellMatch in [Regex]::Matches($rowMatch.Groups[1].Value, "<t[hd]\b[^>]*>(.*?)</t[hd]>", $options)) {
            $cellText = Convert-HtmlInlineToText -Html $cellMatch.Groups[1].Value
            $cells.Add(($cellText -replace "\|", "\|"))
        }
        if ($cells.Count -gt 0) {
            $rows.Add(@($cells))
            $maximumColumns = [Math]::Max($maximumColumns, $cells.Count)
        }
    }


    if ($rows.Count -eq 0 -or $maximumColumns -eq 0) {
        return Convert-HtmlInlineToText -Html $TableMatch.Value
    }


    $lines = [System.Collections.Generic.List[string]]::new()
    for ($rowIndex = 0; $rowIndex -lt $rows.Count; $rowIndex++) {
        $paddedCells = [System.Collections.Generic.List[string]]::new()
        foreach ($cell in @($rows[$rowIndex])) {
            $paddedCells.Add([string]$cell)
        }
        while ($paddedCells.Count -lt $maximumColumns) {
            $paddedCells.Add("")
        }


        $lines.Add("| " + ($paddedCells -join " | ") + " |")
        if ($rowIndex -eq 0) {
            $separatorCells = @(for ($columnIndex = 0; $columnIndex -lt $maximumColumns; $columnIndex++) { "---" })
            $lines.Add("| " + ($separatorCells -join " | ") + " |")
        }
    }


    return "`n`n" + ($lines -join "`n") + "`n`n"
}


function Convert-HtmlListToMarkdown {
    param(
        [Text.RegularExpressions.Match]$ListMatch,
        [switch]$Ordered
    )


    $options = [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::Singleline
    $lines = [System.Collections.Generic.List[string]]::new()
    $index = 1
    foreach ($itemMatch in [Regex]::Matches($ListMatch.Value, "<li\b[^>]*>(.*?)</li>", $options)) {
        $itemText = Convert-HtmlInlineToText -Html $itemMatch.Groups[1].Value
        $prefix = if ($Ordered) { "$index. " } else { "- " }
        $lines.Add($prefix + $itemText)
        $index++
    }


    return "`n`n" + ($lines -join "`n") + "`n`n"
}


function Convert-ReplyReferenceToMarkdown {
    param(
        $Attachment,
        [System.Collections.Generic.Dictionary[string,string]]$IdentityMap,
        [TimeZoneInfo]$TimeZone,
        [string]$ChatId,
        [System.Collections.Generic.Dictionary[string,object]]$MessageMap
    )


    $script:ExpandedReplyReferenceCount++
    $attachmentContentType = [string](Get-OptionalPropertyValue -InputObject $Attachment -Name "contentType")
    $rawContent = [string](Get-OptionalPropertyValue -InputObject $Attachment -Name "content")
    if ([string]::IsNullOrWhiteSpace($rawContent)) {
        $script:IncompleteReplyReferenceCount++
        $script:ReplyAttachmentContentMissingCount++
        return "[Reply context unavailable from Microsoft Graph]"
    }


    try {
        $reference = $rawContent | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        $script:IncompleteReplyReferenceCount++
        $script:ReplyAttachmentJsonErrorCount++
        return "[Reply context could not be decoded from Microsoft Graph]"
    }


    $senderSet = $null
    $originalSentDateTime = $null
    $originalMessage = $null


    if ([string]::Equals($attachmentContentType, "messageReference", [StringComparison]::OrdinalIgnoreCase)) {
        $referenceMessageId = [string](Get-OptionalPropertyValue -InputObject $reference -Name "messageId")
        if ([string]::IsNullOrWhiteSpace($referenceMessageId)) {
            $referenceMessageId = [string](Get-OptionalPropertyValue -InputObject $Attachment -Name "id")
        }


        if ([string]::IsNullOrWhiteSpace($referenceMessageId) -or [string]::IsNullOrWhiteSpace($ChatId)) {
            $script:IncompleteReplyReferenceCount++
            $script:ReplySourceBodyMissingCount++
            return "[Referenced message ID unavailable from Microsoft Graph]"
        }


        if ($null -ne $MessageMap -and $MessageMap.ContainsKey($referenceMessageId)) {
            $originalMessage = $MessageMap[$referenceMessageId]
        }
        else {
            $encodedChatId = [Uri]::EscapeDataString($ChatId)
            $encodedMessageId = [Uri]::EscapeDataString($referenceMessageId)
            $originalMessage = Invoke-GraphGetWithRetry -Uri "https://graph.microsoft.com/v1.0/chats/$encodedChatId/messages/$encodedMessageId"
            $script:FetchedReplySourceMessageCount++
            if ($null -ne $MessageMap) {
                $MessageMap[$referenceMessageId] = $originalMessage
            }
        }


        $senderSet = Get-OptionalPropertyValue -InputObject $originalMessage -Name "from"
        $originalSentDateTime = [string](Get-OptionalPropertyValue -InputObject $originalMessage -Name "createdDateTime")
        foreach ($identityProperty in @("user", "application", "device")) {
            Add-IdentityToMap -Map $IdentityMap -Identity (Get-OptionalPropertyValue -InputObject $senderSet -Name $identityProperty)
        }


        $body = Get-OptionalPropertyValue -InputObject $originalMessage -Name "body"
        $originalContent = [string](Get-OptionalPropertyValue -InputObject $body -Name "content")
        if ([string]::IsNullOrWhiteSpace($originalContent)) {
            $script:IncompleteReplyReferenceCount++
            $script:ReplySourceBodyMissingCount++
            $originalMarkdown = "[Original message content unavailable from Microsoft Graph]"
        }
        else {
            $originalMarkdown = Convert-TeamsBodyToMarkdown -Message $originalMessage -IdentityMap $IdentityMap -TimeZone $TimeZone -ChatId $ChatId -MessageMap $MessageMap
        }
    }
    else {
        $senderSet = Get-OptionalPropertyValue -InputObject $reference -Name "originalMessageSender"
        $originalSentDateTime = [string](Get-OptionalPropertyValue -InputObject $reference -Name "originalSentDateTime")
        $originalContent = [string](Get-OptionalPropertyValue -InputObject $reference -Name "originalMessageContent")
        if ([string]::IsNullOrWhiteSpace($originalContent)) {
            $script:IncompleteReplyReferenceCount++
            $script:ReplySourceBodyMissingCount++
            $originalMarkdown = "[Original message content unavailable from Microsoft Graph]"
        }
        else {
            $originalMessage = [PSCustomObject]@{
                body = [PSCustomObject]@{
                    contentType = "html"
                    content = $originalContent
                }
            }
            $originalMarkdown = Convert-TeamsBodyToMarkdown -Message $originalMessage -IdentityMap $IdentityMap -TimeZone $TimeZone -ChatId $ChatId -MessageMap $MessageMap
        }
    }


    $senderName = "unknown sender"
    foreach ($identityProperty in @("user", "application", "device")) {
        $identity = Get-OptionalPropertyValue -InputObject $senderSet -Name $identityProperty
        $resolvedName = Resolve-IdentityDisplayName -Identity $identity -IdentityMap $IdentityMap
        if (-not [string]::IsNullOrWhiteSpace($resolvedName)) {
            $senderName = $resolvedName
            break
        }
    }
    if ([string]::Equals($senderName, "unknown sender", [StringComparison]::Ordinal)) {
        $script:ReplySenderNameUnavailableCount++
    }


    $sentText = "unknown time"
    if (-not [string]::IsNullOrWhiteSpace($originalSentDateTime)) {
        try {
            $sentUtc = [DateTimeOffset]::Parse($originalSentDateTime)
            if ($null -ne $TimeZone) {
                $sentLocal = [TimeZoneInfo]::ConvertTime($sentUtc, $TimeZone)
                $sentText = $sentLocal.ToString("yyyy-MM-dd HH:mm:ss zzz")
            }
            else {
                $sentText = $sentUtc.ToString("yyyy-MM-dd HH:mm:ss zzz")
            }
        }
        catch {
            $sentText = $originalSentDateTime
        }
    }
    else {
        $script:ReplyTimestampUnavailableCount++
    }


    $quotedBody = (($originalMarkdown -replace "\r\n?", "`n") -split "`n" | ForEach-Object {
        if ([string]::IsNullOrWhiteSpace($_)) { ">" } else { "> $_" }
    }) -join "`n"


    return "`n`n> **Replying to $senderName | $sentText**`n>`n$quotedBody`n`n"
}


function Convert-TeamsBodyToMarkdown {
    param(
        $Message,
        [System.Collections.Generic.Dictionary[string,string]]$IdentityMap,
        [TimeZoneInfo]$TimeZone,
        [string]$ChatId,
        [System.Collections.Generic.Dictionary[string,object]]$MessageMap
    )


    $bodyProperty = $Message.PSObject.Properties['body']
    if ($null -eq $bodyProperty -or $null -eq $bodyProperty.Value) {
        $eventProperty = $Message.PSObject.Properties['eventDetail']
        if ($null -ne $eventProperty -and $null -ne $eventProperty.Value) {
            return "[Teams system event]"
        }
        return "[empty, deleted, or system message body]"
    }


    $content = [string]$bodyProperty.Value.content
    $contentType = [string]$bodyProperty.Value.contentType


    if ([string]::IsNullOrWhiteSpace($content)) {
        $eventProperty = $Message.PSObject.Properties['eventDetail']
        if ($null -ne $eventProperty -and $null -ne $eventProperty.Value) {
            return "[Teams system event]"
        }
        return "[empty, deleted, or system message body]"
    }


    if (-not [string]::Equals($contentType, "html", [StringComparison]::OrdinalIgnoreCase)) {
        return ($content -replace "\r\n?", "`n").Trim()
    }


    $options = [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::Singleline
    $text = $content


    $text = [Regex]::Replace($text, "<table\b[^>]*>.*?</table>", [Text.RegularExpressions.MatchEvaluator]{
        param($match)
        Convert-HtmlTableToMarkdown -TableMatch $match
    }, $options)


    $text = [Regex]::Replace($text, "<ol\b[^>]*>.*?</ol>", [Text.RegularExpressions.MatchEvaluator]{
        param($match)
        Convert-HtmlListToMarkdown -ListMatch $match -Ordered
    }, $options)


    $text = [Regex]::Replace($text, "<ul\b[^>]*>.*?</ul>", [Text.RegularExpressions.MatchEvaluator]{
        param($match)
        Convert-HtmlListToMarkdown -ListMatch $match
    }, $options)


    $text = [Regex]::Replace($text, "<emoji\b(?<attrs>[^>]*)>.*?</emoji>", [Text.RegularExpressions.MatchEvaluator]{
        param($match)
        $altMatch = [Regex]::Match($match.Groups["attrs"].Value, 'alt\s*=\s*["''](?<value>.*?)["'']', $options)
        if ($altMatch.Success) {
            return [Net.WebUtility]::HtmlDecode($altMatch.Groups["value"].Value)
        }
        return ""
    }, $options)


    $text = [Regex]::Replace($text, "<img\b[^>]*>", "[inline image - not retrieved]", $options)
    $messageAttachments = @(Get-OptionalPropertyValue -InputObject $Message -Name "attachments")
    $text = [Regex]::Replace($text, "<attachment\b(?<attrs>[^>]*)>.*?</attachment>", [Text.RegularExpressions.MatchEvaluator]{
        param($match)
        $idMatch = [Regex]::Match($match.Groups["attrs"].Value, 'id\s*=\s*["''](?<value>.*?)["'']', $options)
        if ($idMatch.Success) {
            $attachmentId = [Net.WebUtility]::HtmlDecode($idMatch.Groups["value"].Value)
            $attachment = $messageAttachments | Where-Object {
                [string]::Equals([string](Get-OptionalPropertyValue -InputObject $_ -Name "id"), $attachmentId, [StringComparison]::OrdinalIgnoreCase)
            } | Select-Object -First 1
            $contentType = [string](Get-OptionalPropertyValue -InputObject $attachment -Name "contentType")
            if (-not [string]::IsNullOrWhiteSpace($contentType) -and $contentType -match "(?i)messageReference") {
                return (Convert-ReplyReferenceToMarkdown -Attachment $attachment -IdentityMap $IdentityMap -TimeZone $TimeZone -ChatId $ChatId -MessageMap $MessageMap)
            }
        }
        return "[attachment - open in Teams]"
    }, $options)


    $text = [Regex]::Replace($text, "<at\b[^>]*>(.*?)</at>", [Text.RegularExpressions.MatchEvaluator]{
        param($match)
        "@" + (Convert-HtmlInlineToText -Html $match.Groups[1].Value)
    }, $options)


    $text = [Regex]::Replace($text, "<a\b(?<attrs>[^>]*)>(?<label>.*?)</a>", [Text.RegularExpressions.MatchEvaluator]{
        param($match)
        $label = Convert-HtmlInlineToText -Html $match.Groups["label"].Value
        $hrefMatch = [Regex]::Match($match.Groups["attrs"].Value, 'href\s*=\s*["''](?<value>.*?)["'']', $options)
        if (-not $hrefMatch.Success) {
            return $label
        }
        $href = [Net.WebUtility]::HtmlDecode($hrefMatch.Groups["value"].Value)
        if ($href -match "(?i)^mailto:(.+)$" -and [string]::Equals($label, $Matches[1], [StringComparison]::OrdinalIgnoreCase)) {
            return $label
        }
        return "[$label]($href)"
    }, $options)


    $text = [Regex]::Replace($text, "<(strong|b)\b[^>]*>(.*?)</\1>", '**$2**', $options)
    $text = [Regex]::Replace($text, "<(em|i)\b[^>]*>(.*?)</\1>", '*$2*', $options)
    $text = [Regex]::Replace($text, "<code\b[^>]*>(.*?)</code>", '`$1`', $options)
    $text = [Regex]::Replace($text, "<hr\b[^>]*>", "`n`n---`n`n", $options)
    $text = [Regex]::Replace($text, "<br\s*/?>", "`n", $options)
    $text = [Regex]::Replace($text, "</?(p|div|pre)\b[^>]*>", "`n`n", $options)
    $text = [Regex]::Replace($text, "<blockquote\b[^>]*>", "`n`n> ", $options)
    $text = [Regex]::Replace($text, "</blockquote>", "`n`n", $options)
    $text = [Regex]::Replace($text, "<li\b[^>]*>", "`n- ", $options)
    $text = [Regex]::Replace($text, "</li>", "`n", $options)
    $text = [Regex]::Replace($text, "</?(ul|ol)\b[^>]*>", "`n", $options)
    $text = [Regex]::Replace($text, "<[^>]+>", "", $options)
    $text = [Net.WebUtility]::HtmlDecode($text).Replace([char]0x00A0, " ")
    $text = $text -replace "\r\n?", "`n"


    $cleanLines = foreach ($line in ($text -split "`n")) {
        ([Regex]::Replace($line, "[ \t]+", " ")).Trim()
    }
    $text = $cleanLines -join "`n"
    $text = [Regex]::Replace($text, "`n{3,}", "`n`n")
    return $text.Trim()}


function Add-MessageAnnotations {
    param(
        [Text.StringBuilder]$Builder,
        $Message,
        [TimeZoneInfo]$TimeZone,
        [System.Collections.Generic.Dictionary[string,string]]$IdentityMap,
        [string]$ChatId,
        [System.Collections.Generic.Dictionary[string,object]]$MessageMap
    )


    $lastEditedDateTime = [string](Get-OptionalPropertyValue -InputObject $Message -Name "lastEditedDateTime")
    if (-not [string]::IsNullOrWhiteSpace($lastEditedDateTime)) {
        $editedUtc = [DateTimeOffset]::Parse($lastEditedDateTime)
        $editedLocal = [TimeZoneInfo]::ConvertTime($editedUtc, $TimeZone)
        [void]$Builder.AppendLine("[Edited: $($editedLocal.ToString('yyyy-MM-dd HH:mm:ss zzz'))]")
    }


    $deletedDateTime = [string](Get-OptionalPropertyValue -InputObject $Message -Name "deletedDateTime")
    if (-not [string]::IsNullOrWhiteSpace($deletedDateTime)) {
        $deletedUtc = [DateTimeOffset]::Parse($deletedDateTime)
        $deletedLocal = [TimeZoneInfo]::ConvertTime($deletedUtc, $TimeZone)
        [void]$Builder.AppendLine("[Deleted: $($deletedLocal.ToString('yyyy-MM-dd HH:mm:ss zzz'))]")
    }


    $attachments = @(Get-OptionalPropertyValue -InputObject $Message -Name "attachments")
    foreach ($attachment in $attachments) {
        $attachmentName = [string]$attachment.name
        if ([string]::IsNullOrWhiteSpace($attachmentName)) {
            $attachmentName = "unnamed attachment"
        }


        if ([string]$attachment.contentType -match "(?i)messageReference") {
            $body = Get-OptionalPropertyValue -InputObject $Message -Name "body"
            $bodyContent = [string](Get-OptionalPropertyValue -InputObject $body -Name "content")
            $attachmentId = [string](Get-OptionalPropertyValue -InputObject $attachment -Name "id")
            $attachmentTagPattern = '(?i)<attachment\b[^>]*\bid\s*=\s*["'']{0}["'']' -f [Regex]::Escape($attachmentId)
            $bodyContainsReference = -not [string]::IsNullOrWhiteSpace($attachmentId) -and $bodyContent -match $attachmentTagPattern
            if (-not $bodyContainsReference) {
                [void]$Builder.AppendLine((Convert-ReplyReferenceToMarkdown -Attachment $attachment -IdentityMap $IdentityMap -TimeZone $TimeZone -ChatId $ChatId -MessageMap $MessageMap))
            }
        }
        else {
            [void]$Builder.AppendLine("[Attachment: $attachmentName - open the file in Teams]")
        }
    }


    $reactions = @(Get-OptionalPropertyValue -InputObject $Message -Name "reactions")
    foreach ($reaction in $reactions) {
        $reactionType = [string]$reaction.reactionType
        $reactor = "unknown user"
        $reactionIdentitySet = Get-OptionalPropertyValue -InputObject $reaction -Name "user"
        $reactionUser = Get-OptionalPropertyValue -InputObject $reactionIdentitySet -Name "user"
        $resolvedReactor = Resolve-IdentityDisplayName -Identity $reactionUser -IdentityMap $IdentityMap
        if (-not [string]::IsNullOrWhiteSpace($resolvedReactor)) {
            $reactor = $resolvedReactor
        }
        else {
            $script:UnresolvedReactionCount++
        }
        [void]$Builder.AppendLine("[Reaction: $reactionType by $reactor]")
    }
}


function Write-FailureLog {
    param(
        [string]$Directory,
        [string]$FailureMessage
    )


    try {
        if (-not [string]::IsNullOrWhiteSpace($Directory)) {
            [IO.Directory]::CreateDirectory($Directory) | Out-Null
            $logPath = Join-Path $Directory "Teams-Chat-Transcript-LastRun.log"
            $line = "$(Get-Date -Format o) FAILED: $FailureMessage"
            [IO.File]::WriteAllText($logPath, $line + [Environment]::NewLine, [Text.UTF8Encoding]::new($false))
        }
    }
    catch {
        Write-Warning "The failure log could not be written: $($_.Exception.Message)"
    }
}


$resolvedOutputDirectory = $null
$temporaryPath = $null


try {
    $resolvedOutputDirectory = Resolve-OutputDirectory -RequestedDirectory $OutputDirectory
    [IO.Directory]::CreateDirectory($resolvedOutputDirectory) | Out-Null


    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw "Microsoft.Graph.Authentication is not installed. Run Setup-WeeklyTeamsTranscript.ps1 interactively first."
    }


    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop


    Connect-MgGraph -Scopes "Chat.Read", "User.Read" -ContextScope CurrentUser -NoWelcome -ClientTimeout 45
    $context = Get-MgContext
    if ($null -eq $context -or [string]::IsNullOrWhiteSpace([string]$context.Account)) {
        throw "Microsoft Graph authentication did not return a signed-in account."
    }


    if ([string]::IsNullOrWhiteSpace($ExpectedAccount)) {
        $ExpectedAccount = [string]$context.Account
    }
    elseif (-not ([string]$context.Account).Equals($ExpectedAccount, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Microsoft Graph authenticated as '$($context.Account)', but this export requires '$ExpectedAccount'."
    }
    if (@($context.Scopes) -notcontains "Chat.Read") {
        throw "The authenticated Microsoft Graph session does not include the delegated Chat.Read permission."
    }


    if (@($context.Scopes) -notcontains "User.Read") {
        throw "The authenticated Microsoft Graph session does not include the delegated User.Read permission."
    }


    $me = Invoke-GraphGetWithRetry -Uri 'https://graph.microsoft.com/v1.0/me?$select=id,displayName,userPrincipalName,mail'
    $script:ExpectedUserId = [string]$me.id
    $script:ExpectedUserDisplayName = [string]$me.displayName
    if ([string]::IsNullOrWhiteSpace($script:ExpectedUserId)) {
        throw "Microsoft Graph did not return the signed-in user's ID."
    }
    if (-not [string]::IsNullOrWhiteSpace($script:ExpectedUserDisplayName)) {
        foreach ($candidateId in @(Get-IdentityIdCandidates -IdentityId $script:ExpectedUserId)) {
            $script:KnownIdentityNames[$candidateId] = $script:ExpectedUserDisplayName
        }
    }


    $reportingTimeZone = Get-ReportingTimeZone
    $nowLocal = [TimeZoneInfo]::ConvertTime([DateTimeOffset]::UtcNow, $reportingTimeZone)
    $startLocal = [DateTime]::SpecifyKind($nowLocal.Date.AddDays(-$DaysBack), [DateTimeKind]::Unspecified)
    $endLocal = [DateTime]::SpecifyKind($nowLocal.Date.AddDays(1), [DateTimeKind]::Unspecified)
    $startUtc = [DateTimeOffset]([TimeZoneInfo]::ConvertTimeToUtc($startLocal, $reportingTimeZone))
    $endUtc = [DateTimeOffset]([TimeZoneInfo]::ConvertTimeToUtc($endLocal, $reportingTimeZone))


    Write-Host "Retrieving the complete Teams chat inventory for $ExpectedAccount..."
    $allChats = @(Get-AllGraphCollectionItems -InitialUri 'https://graph.microsoft.com/v1.0/me/chats?$top=50&$expand=members,lastMessagePreview')
    if ($allChats.Count -eq 0) {
        throw "Microsoft Graph returned zero chats. The export was stopped because chat coverage could not be validated."
    }


    $chatResults = [System.Collections.Generic.List[object]]::new()
    $excludedChats = [System.Collections.Generic.List[string]]::new()
    $explicitlyExcludedChats = [System.Collections.Generic.List[string]]::new()
    $messageKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $datesWithMessages = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $completedChats = 0
    $inactiveChats = 0


    foreach ($chat in $allChats) {
        $chatId = [string](Get-OptionalPropertyValue -InputObject $chat -Name "id")
        if (@($ExcludedChatIds) -contains $chatId) {
            $chatLabel = Get-ChatLabel -Chat $chat -Members @()
            $excludedChats.Add($chatLabel)
            $explicitlyExcludedChats.Add("${chatLabel} (chatId: $chatId)")
            continue
        }


        if (-not (Test-ChatMayHaveMessagesInWindow -Chat $chat -StartUtc $startUtc)) {
            $inactiveChats++
            continue
        }


        $members = @(Get-ChatMembers -Chat $chat)
        $chatLabel = Get-ChatLabel -Chat $chat -Members $members


        if (Test-IsExcludedChat -Chat $chat -Members $members) {
            $excludedChats.Add($chatLabel)
            continue
        }


        Write-Host "Reading chat $($completedChats + 1): $chatLabel"
        $messages = @(Get-ChatMessagesInWindow -ChatId ([string]$chat.id) -StartUtc $startUtc -EndUtc $endUtc)
        $completedChats++


        $uniqueMessages = [System.Collections.Generic.List[object]]::new()
        foreach ($message in $messages) {
            $messageKey = "$($chat.id)|$($message.id)"
            if ($messageKeys.Add($messageKey)) {
                $uniqueMessages.Add($message)
                $createdUtc = [DateTimeOffset]::Parse([string]$message.createdDateTime)
                $createdLocal = [TimeZoneInfo]::ConvertTime($createdUtc, $reportingTimeZone)
                [void]$datesWithMessages.Add($createdLocal.ToString("yyyy-MM-dd"))
            }
        }


        if ($uniqueMessages.Count -gt 0) {
            $sortedMessages = @($uniqueMessages | Sort-Object { [DateTimeOffset]::Parse([string]$_.createdDateTime) })
            $messageMap = [System.Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($sortedMessage in $sortedMessages) {
                $sortedMessageId = [string](Get-OptionalPropertyValue -InputObject $sortedMessage -Name "id")
                if (-not [string]::IsNullOrWhiteSpace($sortedMessageId)) {
                    $messageMap[$sortedMessageId] = $sortedMessage
                }
            }
            $identityMap = New-ChatIdentityMap -Members $members -Messages $sortedMessages
            $chatLabel = Resolve-ChatLabelFromMessages -CurrentLabel $chatLabel -Messages $sortedMessages -IdentityMap $identityMap
            $latestActivity = $sortedMessages |
                ForEach-Object { [DateTimeOffset]::Parse([string]$_.createdDateTime) } |
                Sort-Object -Descending |
                Select-Object -First 1
            $chatResults.Add([PSCustomObject]@{
                Chat = $chat
                Label = $chatLabel
                Members = $members
                Messages = $sortedMessages
                MessageMap = $messageMap
                IdentityMap = $identityMap
                LatestActivity = $latestActivity
            })
        }
    }


    if ($completedChats -ne ($allChats.Count - $excludedChats.Count - $inactiveChats)) {
        throw "Chat coverage verification failed. Inventory: $($allChats.Count); no activity in window: $inactiveChats; excluded: $($excludedChats.Count); completed: $completedChats."
    }


    $builder = [Text.StringBuilder]::new()
    [void]$builder.AppendLine("# Teams Chat Transcript")
    [void]$builder.AppendLine()
    [void]$builder.AppendLine("**Reporting period:** $($startLocal.ToString('yyyy-MM-dd HH:mm:ss')) $($reportingTimeZone.Id) through $($endLocal.ToString('yyyy-MM-dd HH:mm:ss')) $($reportingTimeZone.Id) (end exclusive)")
    [void]$builder.AppendLine()


    $orderedChatResults = @($chatResults | Sort-Object LatestActivity -Descending)
    foreach ($chatResult in $orderedChatResults) {
        $safeLabel = ([string]$chatResult.Label -replace "[\r\n]+", " ").Trim()
        [void]$builder.AppendLine("## $safeLabel")
        [void]$builder.AppendLine()


        foreach ($message in $chatResult.Messages) {
            $createdUtc = [DateTimeOffset]::Parse([string]$message.createdDateTime)
            $createdLocal = [TimeZoneInfo]::ConvertTime($createdUtc, $reportingTimeZone)
            $sender = (Get-MessageSender -Message $message) -replace "[\r\n]+", " "
            [void]$builder.AppendLine("**$($createdLocal.ToString('yyyy-MM-dd HH:mm:ss zzz')) | $sender**")
            [void]$builder.AppendLine()
            [void]$builder.AppendLine((Convert-TeamsBodyToMarkdown -Message $message -IdentityMap $chatResult.IdentityMap -TimeZone $reportingTimeZone -ChatId ([string]$chatResult.Chat.id) -MessageMap $chatResult.MessageMap))
            [void]$builder.AppendLine()
            Add-MessageAnnotations -Builder $builder -Message $message -TimeZone $reportingTimeZone -IdentityMap $chatResult.IdentityMap -ChatId ([string]$chatResult.Chat.id) -MessageMap $chatResult.MessageMap
            [void]$builder.AppendLine()
        }
    }


    $zeroMessageDates = [System.Collections.Generic.List[string]]::new()
    for ($date = $startLocal.Date; $date -lt $endLocal.Date; $date = $date.AddDays(1)) {
        $dateText = $date.ToString("yyyy-MM-dd")
        if (-not $datesWithMessages.Contains($dateText)) {
            $zeroMessageDates.Add($dateText)
        }
    }


    $totalWaitSeconds = ($script:RetryWaits | Measure-Object -Sum).Sum
    if ($null -eq $totalWaitSeconds) {
        $totalWaitSeconds = 0
    }


    $auditBuilder = [Text.StringBuilder]::new()
    [void]$auditBuilder.AppendLine("Reporting period: $($startLocal.ToString('yyyy-MM-dd HH:mm:ss')) $($reportingTimeZone.Id) through $($endLocal.ToString('yyyy-MM-dd HH:mm:ss')) $($reportingTimeZone.Id) (end exclusive)")
    [void]$auditBuilder.AppendLine("Source: Microsoft Graph delegated access using Chat.Read.")
    [void]$auditBuilder.AppendLine("Discovery: fully paginated /me/chats inventory and /chats/{chat-id}/messages retrieval for every active, non-excluded chat.")
    [void]$auditBuilder.AppendLine("UTC query window: $($startUtc.UtcDateTime.ToString('o')) through $($endUtc.UtcDateTime.ToString('o')) (end exclusive).")
    [void]$auditBuilder.AppendLine("Total chats in inventory: $($allChats.Count).")
    [void]$auditBuilder.AppendLine("Chats proven inactive by lastMessagePreview: $inactiveChats.")
    [void]$auditBuilder.AppendLine("Chats excluded by scope: $($excludedChats.Count).")
    [void]$auditBuilder.AppendLine("Active non-excluded chats queried successfully: $completedChats.")
    [void]$auditBuilder.AppendLine("Chats represented in transcript: $($orderedChatResults.Count).")
    [void]$auditBuilder.AppendLine("Unique conversation messages: $($messageKeys.Count).")
    [void]$auditBuilder.AppendLine("Microsoft Graph requests: $script:GraphRequestCount.")
    [void]$auditBuilder.AppendLine("Rate-limit or transient-error waits: $($script:RetryWaits.Count), totaling $totalWaitSeconds seconds.")
    [void]$auditBuilder.AppendLine("Dates with zero retrieved messages: $(if ($zeroMessageDates.Count -eq 0) { 'none' } else { $zeroMessageDates -join ', ' }).")
    [void]$auditBuilder.AppendLine("Message formatting: Teams HTML converted to clean Markdown; tables and lists converted structurally.")
    [void]$auditBuilder.AppendLine("Reply references expanded with original sender, timestamp, and message content: $script:ExpandedReplyReferenceCount.")
    [void]$auditBuilder.AppendLine("Referenced reply-source messages fetched individually because they were not already in the reporting-window result set: $script:FetchedReplySourceMessageCount.")
    [void]$auditBuilder.AppendLine("Reply references lacking original message content in returned Graph data: $script:IncompleteReplyReferenceCount.")
    [void]$auditBuilder.AppendLine("Reply sender names unavailable after identity matching: $script:ReplySenderNameUnavailableCount.")
    [void]$auditBuilder.AppendLine("Reply timestamps unavailable from returned Graph data: $script:ReplyTimestampUnavailableCount.")
    [void]$auditBuilder.AppendLine("Inline images: represented as [inline image - not retrieved].")
    [void]$auditBuilder.AppendLine("Attachments, edits, and reactions: recorded when returned by Microsoft Graph; reaction user IDs resolved through the chat roster and message senders.")
    [void]$auditBuilder.AppendLine("Reaction user identities not resolvable from returned Graph data: $script:UnresolvedReactionCount.")
    [void]$auditBuilder.AppendLine("Meeting chats without an expanded member roster: $($script:MeetingChatsWithoutRoster.Count).")
    [void]$auditBuilder.AppendLine("Explicit chat-ID exclusions: $(if ($explicitlyExcludedChats.Count -eq 0) { 'none' } else { $explicitlyExcludedChats -join '; ' }).")
    [void]$auditBuilder.AppendLine("Coverage: every chat with possible reporting-window activity was either excluded by scope or queried successfully; zero possible active chats failed.")


    $startFileDate = $startLocal.ToString("yyyy-MM-dd")
    $endFileDate = $endLocal.ToString("yyyy-MM-dd")
    $baseName = "Teams-Chat-Transcript-$startFileDate`_to_$endFileDate"
    $finalPath = Join-Path $resolvedOutputDirectory "$baseName.md"
    if (Test-Path -LiteralPath $finalPath) {
        $rerunSuffix = Get-Date -Format "yyyyMMdd-HHmmss"
        $finalPath = Join-Path $resolvedOutputDirectory "$baseName`_rerun-$rerunSuffix.md"
    }


    $temporaryPath = "$finalPath.incomplete"
    [IO.File]::WriteAllText($temporaryPath, $builder.ToString(), [Text.UTF8Encoding]::new($false))


    $verificationText = [IO.File]::ReadAllText($temporaryPath, [Text.UTF8Encoding]::new($false))
    if ([string]::IsNullOrWhiteSpace($verificationText)) {
        throw "The generated Markdown file is empty."
    }


    $messageHeaderPattern = '(?m)^\*\*\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} [+-]\d{2}:\d{2} \| .+\*\*\r?$'
    $messageHeaderCount = [Regex]::Matches($verificationText, $messageHeaderPattern).Count
    if ($messageHeaderCount -ne $messageKeys.Count) {
        throw "Saved-file verification failed. Expected $($messageKeys.Count) message headers but found $messageHeaderCount."
    }


    if ($verificationText -match "(?m)^## How this transcript was compiled$") {
        throw "Saved-file verification failed because audit content appeared in the reader-facing transcript."
    }


    Move-Item -LiteralPath $temporaryPath -Destination $finalPath


    try {
        $successLogPath = Join-Path $resolvedOutputDirectory "Teams-Chat-Transcript-LastRun.log"
        $successLine = "$(Get-Date -Format o) SUCCESS: $($messageKeys.Count) messages across $($orderedChatResults.Count) chats. File: $finalPath"
        $successLog = $successLine + [Environment]::NewLine + $auditBuilder.ToString()
        [IO.File]::WriteAllText($successLogPath, $successLog, [Text.UTF8Encoding]::new($false))
    }
    catch {
        Write-Warning "The transcript was completed, but the last-run log could not be written: $($_.Exception.Message)"
    }


    Write-Host "Complete transcript created: $finalPath"
    Write-Host "$($messageKeys.Count) messages across $($orderedChatResults.Count) chats; $completedChats active non-excluded chats queried; $inactiveChats inactive chats skipped by last-message timestamp; zero active chat failures."
    exit 0
}
catch {
    $failureMessage = $_.Exception.Message
    if (-not [string]::IsNullOrWhiteSpace($temporaryPath) -and (Test-Path -LiteralPath $temporaryPath)) {
        Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue
    }
    Write-FailureLog -Directory $resolvedOutputDirectory -FailureMessage $failureMessage
    Write-Error "Weekly Teams transcript FAILED. No completed transcript was created. $failureMessage"
    exit 1
}