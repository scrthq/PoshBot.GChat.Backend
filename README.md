# PoshBot.GChat.Backend <!-- omit in toc -->

> _Google Chat backend for PoshBot leveraging a Google Sheet as a message queue with Apps Script as the bot endpoint managing the Sheet contents_

* [Prerequisites](#prerequisites)
* [Setting up the GChat backend for PoshBot](#setting-up-the-gchat-backend-for-poshbot)
    * [Installing the PoshBot.GChat.Backend module](#installing-the-poshbotgchatbackend-module)
    * [Starting up PoshBot](#starting-up-poshbot)
    * [Running PoshBot as a service](#running-poshbot-as-a-service)
* [Pros and cons with using Google Chat with PoshBot versus Slack](#pros-and-cons-with-using-google-chat-with-poshbot-versus-slack)
    * [API Communication](#api-communication)
    * [Reactions](#reactions)
    * [Cards](#cards)

## Prerequisites

To start using PoshBot with Google Chat, you'll need to have a few things set up/installed first:

1. [PSGSuite](https://github.com/scrthq/PSGSuite)
    * [Documentation](https://github.com/scrthq/PSGSuite/wiki)
    * Miniumum required version: `2.13.0`
    * Developer Console Project must have the following API's enabled:
        * Hangouts Chat API (enabled and configured)
        * Sheets API
2. [Google Apps Script Sheet MQ](https://github.com/scrthq/GoogleAppsScriptSheetMQ)
    * [Documentation in README](https://github.com/scrthq/GoogleAppsScriptSheetMQ)
    * Make sure to validate the connection per the docs!
3. [PoshBot](https://github.com/poshbotio/PoshBot)
    * [Documentation](http://poshbot.readthedocs.io/en/latest/)
4. **PowerShell version 5 or greater**
    * PoshBot is built almost entirely with PowerShell classes which were first introduced in PowerShell 5.

## Setting up the GChat backend for PoshBot

### Installing the PoshBot.GChat.Backend module

You can install the `PoshBot.GChat.Backend` module directly from the PowerShell Gallery. If you do not have PSGSuite and/or PoshBot installed already, this will also install them:

```powershell
Install-Module PoshBot.GChat.Backend -Scope CurrentUser
```

### Starting up PoshBot

Here's a sample script that I use to get PoshBot started. Some important Google Chat/PSGSuite specific configuration items to note are:

1. `BotAdmins`: GChat bot admins need to be listed using their primary email address.
2. `ConfigName`: If you only use one config with PSGSuite, you can exclude this from the `BackendConfiguration` and it will retrieve the correct config name during backend instantiation
3. `SheetId`: **This is 100% necessary!** Without the SheetId, the backend will not know where to connect. The PSGSuite AdminEmail account will need **edit** access to this Sheet as well.
4. `PollingFrequency`: This defaults to 1500ms. You can exclude this from the `BackendConfiguration` if that is fine. Do not set the `PollingFrequency` below 1 second otherwise you risk being rate limited by the default Sheets Read quota of 100 Reads per 100 seconds.

```powershell
# Import necessary modules
Import-Module PoshBot
Import-Module PoshBot.GChat.Backend

# Store config path in variable
$configPath = 'E:\Scripts\PoshBot\GChat\GChatConfig.psd1'

# Create hashtable of parameters for New-PoshBotConfiguration
$botParams = @{
    # The friendly name of the bot instance
    Name                   = 'GChatBot'
    # The primary email address(es) of the admin(s) that can manage the bot
    BotAdmins              = @('admin@domain.com', 'coadmin@domain.com')
    # Universal command prefix for PoshBot.
    # If the message includes this at the start, PoshBot will try to parse the command and 
    # return an error if no matching command is found
    CommandPrefix          = '!'
    # PoshBot log level.
    LogLevel               = 'Verbose'
    # The path containing the configuration files for PoshBot
    ConfigurationDirectory = 'E:\Scripts\PoshBot\GChat'
    # The path where you would like the PoshBot logs to be created
    LogDirectory           = 'E:\Scripts\PoshBot\GChat'
    # The path containing your PoshBot plugins
    PluginDirectory        = 'E:\Scripts\PoshBot\Plugins'

    BackendConfiguration   = @{
        # This is the PSGSuite config name that you would like the GChat backend to run under.
        # This config needs to have access to the Sheet set up as the Message Queue
        ConfigName       = "mydomain"
        # This is the FileID of the Sheet set up as the Message Queue
        SheetId          = "1H7mJoKflklakoJKDSwo923lsdO5sK3mjg"
        # How frequently you'd like to poll the Sheet for new messages.
        # If this is greater than 1000, it's treated as milliseconds, otherwise it's treated as seconds
        PollingFrequency = 1500
        # The friendly name for the backend
        Name             = 'GChatBackend'
    }
}

# Create the bot backend
$backend = New-PoshBotGChatBackend -Configuration $botParams.BackendConfiguration

# Create the bot configuration
$myBotConfig = New-PoshBotConfiguration @botParams

# Save bot configuration
Save-PoshBotConfiguration -InputObject $myBotConfig -Path $configPath -Force

# Create the bot instance from the backend and configuration path
$bot = New-PoshBotInstance -Backend $backend -Path $configPath

# Start the bot
$bot | Start-PoshBot
```

### Running PoshBot as a service

Once you start PoshBot, it will hold the session open. The easiest way to have PoshBot running without tying up a visible PowerShell console is to run your Start script as a service.

Here's a quick guide to installing PoshBot as a service using `NSSM`: https://poshbot.readthedocs.io/en/latest/guides/run-poshbot-as-a-service/

**NOTE:** PSGSuite configurations are typically tied to the user who created them. Make sure you update the service to run as that account. If you are planning on using a service account to run the service, please create a PSGSuite configuration in the context of that service account so it is able to decrypt the configuration while running as a service.


## Pros and cons with using Google Chat with PoshBot versus Slack

Google Chat and Slack have a number of differences in regards to event types and message widgets. Use this to guide your own PoshBot plugin development to ensure that both the user experience and returned command results match expectations no matter what ChatOps client you're using!

### API Communication

Google Chat uses various endpoints, but does not have a WebSocket connection type equivalent to Slack's RealTimeMessaging API that PoshBot's Slack implementation uses. 

This means that Google only sends events to the Bot endpoint if...
* the message was sent via DM to the bot directly or...
* the message was sent in a room the bot is a member of **and** the bot is tagged, i.e. `@PoshBot !help`

### Reactions

Google Chat currently does not support adding reactions to messages, nor does it emit events when reactions are added. Due to this caveat, PoshBot is unable to signal to the sender that it is currently processing the message (indicated by a gear in Slack), the message was processed successfully (green check mark), or any others (i.e. warnings).

### Cards

Google Chat does not support certain card widgets that Slack does, i.e. Thumbnail images. There is logic in place in the GChat backend to provide a best-effort translation, but results may vary.

To assist with developing PoshBot plugins compatible with multiple backends, the module `PoshBot.GChat.Backend` comes with a helper function `New-PoshBotGChatCardResponse`. This function includes the same parameters as `New-PoshBotCardResponse`, while also supporting pipeline input of Google Chat widgets the same as you would use with `Send-GSChatMessage`.

Here's an example from my `Plex` plugin (unreleased) that shows the current process information for Plex Media Server:

```powershell
if ($procs = Get-Process "Plex Media Server" -ErrorAction SilentlyContinue) {
    foreach ($proc in $procs) {
        $Fields = @{
            ProcName  = $proc.Name
            PID       = $proc.Id
            StartTime = $proc.StartTime.ToString("yyyy-MM-dd HH:mm:ss")
        }
        $Fields.Keys | ForEach-Object {
            $title = $_
            Add-GSChatKeyValue -TopLabel $title -Content $Fields[$title] -Icon CONFIRMATION_NUMBER_ICON
        } | Add-GSChatCardSection -SectionHeader "Process Details" | Add-GSChatCard | New-PoshBotGChatCardResponse -Text "Plex is running!" -Fields $fields
    }
}
else {
    New-PoshBotTextResponse -Text "*Plex Media Server is not currently running!* Type ``plex start`` to start Plex"
}
```

When `plex status` is ran from Slack, it returns the following...

![Slack command example](https://github.com/scrthq/PoshBot.GChat.Backend/blob/master/.github/Slack%20Command%20Example.png?raw=true)

... and when ran from Google Chat...

![Google Chat command example](https://github.com/scrthq/PoshBot.GChat.Backend/blob/master/.github/GChat%20Command%20Example.png?raw=true)