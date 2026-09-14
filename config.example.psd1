@{
    # Leave blank to accept whichever Microsoft 365 account is selected at sign-in.
    # Set an address to prevent accidentally exporting the wrong account.
    ExpectedAccount = ""


    # Complete local calendar days before today; valid range is 1 through 30.
    DaysBack = 7


    # Leave blank for: Documents\Teams Transcripts
    OutputDirectory = ""


    # Leave blank to use the Windows computer's local time zone.
    # Examples: "Central Standard Time" on Windows or "America/Chicago" cross-platform.
    TimeZoneId = ""


    # Exact Graph chat IDs that should never be queried.
    ExcludedChatIds = @()


    # Excludes a one-to-one chat when the only other participant uses one of these domains.
    ExcludedEmailDomains = @(
        # "example.com"
    )


    # Case-insensitive PowerShell regular expressions applied to chat topics.
    ExcludedTopicPatterns = @(
        # "\bpersonal\b"
    )


    # Retries for HTTP 429 and transient Microsoft Graph 5xx responses.
    MaxRetries = 8
}