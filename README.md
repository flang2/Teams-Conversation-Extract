# Weekly Teams Transcript

Export the signed-in user's Microsoft Teams chats to a chronological Markdown transcript using delegated Microsoft Graph access.

This package is designed for **manual weekly use**. Scheduling is optional.

## Installation location

After downloading or cloning it, you can store the repository in any permanent local folder, a OneDrive-synchronized folder, another synchronized drive, or a normal Git working directory.

The repository folder may be renamed or moved. Keep its internal files and subfolders together, and keep `Run-WeeklyTeamsTranscript.bat` in the repository root. Transcript output is controlled separately by `OutputDirectory` in `config.psd1`.

## What it does

The exporter:

- Retrieves one-to-one, group, and meeting chats for the signed-in user.  
- Does not include Teams channels.  
- Enumerates the complete chat inventory instead of relying on keyword searches.  
- Retrieves every message in the reporting window from every active, non-excluded chat.  
- Preserves message bodies, paragraphs, lists, links, tables, mentions, emojis, edits, attachments, reactions, and available reply context.  
- Writes a clean Markdown transcript and a separate detailed run log.  
- Stops without creating a completed transcript if it can't retrieve an active, non-excluded chat.

By default, the reporting window begins seven local calendar days before today and ends at the start of tomorrow.

## 

## Required directory structure

The repository folder may be renamed or moved, but the files inside it must retain this structure:

Weekly-Teams-Transcript/

├── tools/  
│   ├── Export-WeeklyTeamsTranscript.ps1  
│   ├── Install-WeeklyTeamsTranscriptTask.ps1  
│   ├── Invoke-WeeklyTeamsTranscript.ps1  
│   └── Setup-WeeklyTeamsTranscript.ps1  
│── .github/  
│   └── workflows/  
│       └── powershell.yml  
├── .gitignore  
├── config.example.psd1  
├── LICENSE  
├── README.md  
└── Run-WeeklyTeamsTranscript.bat

After setup, this additional local file appears in the repository root:

config.psd1

Do not move `Run-WeeklyTeamsTranscript.bat` away from the repository root. It finds the runner under `tools` using a relative path. To run it from the Windows desktop, create a shortcut to the BAT file instead of moving the BAT file.

The `.github` folder, workflow, license, and installer are not needed for a manual export, but they should remain in the public GitHub repository.

## Output location

Unless changed in `config.psd1`, transcripts are written to:

C:\\Users\<Windows-user\>\\Documents\\Teams Transcripts

Output is deliberately stored outside the repository so private Teams messages are not accidentally committed to GitHub.

Each successful run creates:

Teams-Chat-Transcript-YYYY-MM-DD\_to\_YYYY-MM-DD.md

Teams-Chat-Transcript-LastRun.log

The exporter first writes an `.incomplete` file. It becomes a completed `.md` file only after saved-file verification succeeds.

## Requirements

- Windows 10 or Windows 11  
- PowerShell 5.1 or PowerShell 7  
- A Microsoft 365 account containing Teams chats  
- Permission to grant delegated Microsoft Graph `Chat.Read` and `User.Read`  
- Internet access during authentication and export

The package does not request tenant-wide `Chat.Read.All`.

## Installation

### 1\. Download or clone the repository

Place the extracted repository in a permanent location, for example:

C:\\Users\<Windows-user\>\\Documents\\Weekly-Teams-Transcript

Do not run the scripts from inside a ZIP file.

### 2\. Open PowerShell in the repository folder

For example:

Set-Location \-LiteralPath "$env:USERPROFILE\\Documents\\Weekly-Teams-Transcript"

Substitute the actual folder name if it differs.

### 3\. Unblock the downloaded files

Windows may mark scripts downloaded from the internet as blocked. Run:

Get-ChildItem \-LiteralPath . \-Recurse \-File | Unblock-File

If Windows still says that a script is not digitally signed, apply a temporary execution-policy override to the current PowerShell window:

Set-ExecutionPolicy \-Scope Process \-ExecutionPolicy Bypass \-Force

This setting ends when that PowerShell window is closed. It does not change the permanent computer or user policy.

### 4\. Optional syntax check

$syntaxErrors \\= @(  
    Get-ChildItem \-LiteralPath .\\tools \-Filter \*.ps1 |  
        ForEach-Object {  
            $tokens \\= $null  
            $parseErrors \\= $null  
            \[System.Management.Automation.Language.Parser\]::ParseFile(  
                $\_.FullName,  
                \[ref\]$tokens,  
                \[ref\]$parseErrors  
            ) | Out-Null  
            $parseErrors  
        }  
)

if ($syntaxErrors.Count \-gt 0\) {  
    $syntaxErrors | Format-List  
    throw "PowerShell syntax validation failed."  
}

Write-Host "All PowerShell scripts passed syntax validation."

### 5\. Authenticate and run the initial test

Run the setup script from PowerShell:

.\\tools\\Setup-WeeklyTeamsTranscript.ps1

Do not double-click the `.ps1` file. If Windows asks which app should open it, close that dialog and run the command above from PowerShell.

Setup will:

1. Create `config.psd1` from `config.example.psd1` if it does not already exist.  
2. Install `Microsoft.Graph.Authentication` for the current Windows user if necessary.  
3. Open a Microsoft sign-in window.  
4. Request delegated `Chat.Read` and `User.Read`.  
5. Run a short one-day test export.

Sign in with the Microsoft 365 account whose Teams chats should be exported.

A successful test ends with messages similar to:

Complete transcript created: ...

... messages across ... chats; ... active non-excluded chats queried; ... inactive chats skipped; zero active chat failures.

Setup and the one-day test export completed successfully.

## Configure the exporter

Open the locally created configuration:

notepad .\\config.psd1

Its structure is:

@{  
    ExpectedAccount \\= ""  
    DaysBack \\= 7  
    OutputDirectory \\= ""  
    TimeZoneId \\= ""  
    ExcludedChatIds \\= @()  
    ExcludedEmailDomains \\= @()  
    ExcludedTopicPatterns \\= @()  
    MaxRetries \\= 8  
}

### ExpectedAccount

Leave blank to accept whichever Microsoft 365 account is selected during sign-in:

ExpectedAccount \\= ""

Set an email address to prevent accidentally exporting the wrong account:

ExpectedAccount \\= "person@example.com"

### DaysBack

The normal weekly value is:

DaysBack \\= 7

Valid values are 1 through 30\.

### OutputDirectory

Leave blank to use `Documents\Teams Transcripts`:

OutputDirectory \\= ""

Or provide a complete Windows path:

OutputDirectory \\= "C:\\Users\\ExampleUser\\Documents\\My Teams Transcripts"

Do not use variables such as `$env:USERPROFILE` inside `config.psd1`. Use a literal complete path.

### TimeZoneId

Leave blank to use the Windows computer's local time zone:

TimeZoneId \\= ""

A specific Windows time-zone ID may be used:

TimeZoneId \\= "Central Standard Time"

### ExcludedChatIds

Use exact Microsoft Graph chat IDs for chats that should never be queried:

ExcludedChatIds \\= @(

    "exact-chat-id-here"

)

Only exclude a chat after confirming it is irrelevant. This is particularly useful for stale meeting chats that Graph lists but returns HTTP 403 because the user is no longer a roster member.

### ExcludedEmailDomains

For one-to-one chats, exclude the conversation when the only other participant uses a configured email domain:

ExcludedEmailDomains \\= @(

    "example.com"

)

This setting is not applied to group or meeting chats.

### ExcludedTopicPatterns

Exclude chats whose topics match PowerShell regular expressions:

ExcludedTopicPatterns \\= @(

    '\\bpersonal\\b',

    '\\btest chat\\b'

)

### MaxRetries

The default is:

MaxRetries \\= 8

The exporter retries Microsoft Graph HTTP 429 rate limits and transient HTTP 5xx failures. It does not silently ignore authorization or coverage failures.

## Run the weekly transcript manually

The easiest method is to double-click:

Run-WeeklyTeamsTranscript.bat

The BAT file:

1. Locates the PowerShell runner under `tools`.  
2. Uses PowerShell 7 when available, otherwise Windows PowerShell.  
3. Loads `config.psd1`.  
4. Runs the export.  
5. Displays either `SUCCESS` or `FAILED`.  
6. Keeps the window open so the result can be read.

The same export can be started directly from PowerShell:

.\\tools\\Invoke-WeeklyTeamsTranscript.ps1

If the computer again blocks the scripts in a new PowerShell window, use:

Set-ExecutionPolicy \-Scope Process \-ExecutionPolicy Bypass \-Force

.\\tools\\Invoke-WeeklyTeamsTranscript.ps1

## Confirm a successful run

Open the configured output folder and check:

1. A new `Teams-Chat-Transcript-*.md` file exists.  
2. `Teams-Chat-Transcript-LastRun.log` begins with `SUCCESS`.  
3. The console reports `zero active chat failures`.  
4. The message and chat counts are plausible.  
5. Expected recent conversations appear in the Markdown transcript.  
6. Private or excluded conversations do not appear.

The log contains retrieval counts and technical audit details. The transcript itself contains only the reader-facing conversations.

## Reply context and reactions

The exporter expands reply references using the original message when Microsoft Graph still makes it available.

If Graph returns the reply message but the original replied-to message was deleted or is otherwise unavailable, the transcript retains the reply and inserts a clear context-unavailable placeholder. This enrichment gap is recorded in the run log but does not discard an otherwise complete week.

Reaction names are shown when Graph supplies an identity that can be matched to chat members or message senders. Unresolvable reaction identities are counted in the run log.

## Failure behavior

The exporter does not create a completed transcript when:

- Authentication fails.  
- An active, non-excluded chat cannot be retrieved.  
- Graph returns an unrecoverable authorization or service error.  
- Chat coverage verification fails.  
- The saved transcript does not contain the expected number of message headers.

A partial `.incomplete` file is removed after failure. Review:

Teams-Chat-Transcript-LastRun.log

## Troubleshooting

### Script is not digitally signed

From the repository root:

Get-ChildItem \-LiteralPath . \-Recurse \-File | Unblock-File

Set-ExecutionPolicy \-Scope Process \-ExecutionPolicy Bypass \-Force

### Wrong Microsoft account

Disconnect-MgGraph \-ErrorAction SilentlyContinue

.\\tools\\Setup-WeeklyTeamsTranscript.ps1

Set `ExpectedAccount` in `config.psd1` when the computer regularly uses multiple Microsoft 365 accounts.

### Authentication expired

Run setup again:

.\\tools\\Setup-WeeklyTeamsTranscript.ps1

Password changes, revoked consent, multifactor authentication, token expiration, or Conditional Access may require a new sign-in.

### HTTP 403 for one meeting chat

Read the complete error. A stale meeting may remain in the chat inventory even though the signed-in user is no longer a roster member.

If the meeting is confirmed irrelevant, add its exact chat ID to `ExcludedChatIds`. Do not suppress every HTTP 403 response because that would hide legitimate coverage failures.

### HTTP 429 rate limit

Wait for the exporter. It automatically respects Microsoft Graph retry delays and retries the same request.

## Desktop shortcut

Keep the BAT file in the repository root. To place a launcher on the desktop:

1. Right-click `Run-WeeklyTeamsTranscript.bat`.  
2. Select **Show more options** if necessary.  
3. Select **Send to \> Desktop (create shortcut)**.

Use the shortcut for future manual weekly runs.

## Optional scheduling

`tools\Install-WeeklyTeamsTranscriptTask.ps1` is included for users who want Windows Task Scheduler automation. Scheduling is not required for installation or manual use.

## Privacy and GitHub safety

Teams transcripts may contain confidential information.

Before committing to GitHub, verify that these files are not staged:

- `config.psd1`  
- Generated transcripts  
- `Teams-Chat-Transcript-LastRun.log`  
- `.incomplete` files  
- Access tokens or authentication caches  
- Downloaded Teams attachments

The supplied `.gitignore` covers the normal configuration, transcript, log, and incomplete filenames, but always review `git status` before committing.

## Microsoft documentation

- [List chats](https://learn.microsoft.com/graph/api/chat-list)  
- [List chat messages](https://learn.microsoft.com/graph/api/chat-list-messages)  
- [Get a chat message](https://learn.microsoft.com/graph/api/chatmessage-get)  
- [List chat members](https://learn.microsoft.com/graph/api/chat-list-members)  
- [Microsoft Graph permissions reference](https://learn.microsoft.com/graph/permissions-reference)

## License

MIT  
