Using module PoshBot

if (!(Get-Module PSGSuite)) {
    Import-Module PSGSuite -MinimumVersion "2.13.0" -Force
}

$Script:_gChatAcked = New-Object System.Collections.ArrayList

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingConvertToSecureStringWithPlainText', '', Scope='Class', Target='*')]
class GChatBackend : Backend {

    # The types of message that we care about from GChat
    # All othere will be ignored
    [string[]]$MessageTypes = @(
        'MESSAGE'
        'REMOVED_FROM_SPACE'
        'ADDED_TO_SPACE'
        'CARD_CLICKED'
    )

    [int]$MaxMessageLength = 4000

    GChatBackend ([string]$ConfigName,[string]$SheetId,[string]$SheetName,[int]$PollingFrequency) {
        if (!(Get-Module PSGSuite)) {
            Import-Module PSGSuite -MinimumVersion "2.13.0" -Verbose:$false -Force
        }
        if ((Show-PSGSuiteConfig).ConfigName -ne $ConfigName) {
            Switch-PSGSuiteConfig -ConfigName $ConfigName
        }
        $config = [ConnectionConfig]::new()
        $config.Credential = New-Object System.Management.Automation.PSCredential($ConfigName,(ConvertTo-SecureString -String $SheetName -AsPlainText -Force))
        $config.Endpoint = $SheetId
        $conn = [GChatConnection]::New($ConfigName,$SheetId,$SheetName,$PollingFrequency)
        $conn.Config = $config
        $this.Connection = $conn
    }

    # Connect to GChat
    [void]Connect() {
        $this.LogInfo('Connecting to backend')
        $this.LogInfo('Listening for the following message types. All others will be ignored', $this.MessageTypes)
        $this.Connection.Connect()
        $this.BotId = $this.GetBotIdentity()
        $this.LoadRooms()
        $this.LoadUsers()
    }

    # Receive a message from the websocket
    [Message[]]ReceiveMessage() {
        $messages = New-Object -TypeName System.Collections.ArrayList
        try {
            # Read the output stream from the receive job and get any messages since our last read
            $jsonResult = $this.Connection.ReadReceiveJob()

            if ($null -ne $jsonResult -and $jsonResult -ne [string]::Empty) {
                #Write-Debug -Message "[GChatBackend:ReceiveMessage] Received `n$jsonResult"
                $this.LogDebug('Received message', $jsonResult)

                $gChatMessages = @($jsonResult | ConvertFrom-Json)
                foreach ($gChatMessage in $gChatMessages) {
                    $gChatEvent = ConvertFrom-Json $gChatMessage.Event

                    # We only care about certain message types from GChat
                    if ($gChatEvent.type -in $this.MessageTypes) {
                        $msg = [Message]::new()

                        # Set the message type and optionally the subtype
                        #$msg.Type = $gChatEvent.type
                        $this.LogVerbose("New [$($gChatEvent.type)] Chat event received")
                        switch ($gChatEvent.type) {
                            'ADDED_TO_SPACE' {
                                $msg.Type = [MessageType]::Message
                                $msg.SubType = [MessageSubtype]::ChannelJoined
                                $msg.To = $gChatEvent.space.name
                                $msg.ToName = $gChatEvent.space.displayName
                                $msg.Text = "$($gChatEvent.type) OriginalMessage: $($gChatMessage.Event -join '')"
                            }
                            'REMOVED_FROM_SPACE' {
                                $msg.Type = [MessageType]::Message
                                $msg.SubType = [MessageSubtype]::ChannelLeft
                                $msg.Text = "$($gChatEvent.type) OriginalMessage: $($gChatMessage.Event -join '')"
                            }
                            'MESSAGE' {
                                $msg.Type = [MessageType]::Message
                                $msg.From = $gChatEvent.message.sender.name -replace "users\/",""
                                $msg.FromName = $gChatEvent.message.sender.displayName
                                $msg.To = $gChatEvent.message.thread.name
                                $msg.ToName = $gChatEvent.message.space.displayName
                                $msg.Text = $gChatEvent.message.argumentText.Trim().Replace('  ',' ').Replace('  ',' ')
                                $msg.Id = $gChatEvent.message.name
                                if ($gChatEvent.space.type -eq 'DM') {
                                    $this.LogDebug("MESSAGE is a DM!")
                                    $msg.IsDM = $true
                                    $msg.ToName = "@$($gChatEvent.user.displayName)"
                                }
                            }
                            'CARD_CLICKED' {
                                $msg.Type = [MessageType]::Message
                                $msg.From = $gChatEvent.user.name
                                $msg.FromName = $gChatEvent.user.displayName
                                $msg.To = $gChatEvent.message.name
                                $msg.ToName = $gChatEvent.message.space.displayName
                                $msg.Text = "CARD_CLICKED ActionMethodName: $($gChatEvent.action.actionMethodName) ActionMethodParams: $(if ($gChatEvent.action.parameters) {"$($gChatEvent.action.parameters | ConvertTo-Json -Depth 5 -Compress)"}else{'{}'}) OriginalMessage: $($gChatMessage.Event -join '')"
                                $msg.Id = $gChatEvent.message.name
                                if ($gChatEvent.space.type -eq 'DM') {
                                    $this.LogDebug("CARD_CLICKED event is a DM!")
                                    $msg.IsDM = $true
                                    $msg.ToName = "@$($gChatEvent.user.displayName)"
                                }
                            }
                        }

                        $this.LogDebug("Message type is [$($msg.Type)`:$($msg.Subtype)] :: From [$($msg.FromName)`:$($msg.From)] :: To [$($msg.ToName)`:$($msg.To)]")

                        $msg.RawMessage = $gChatMessage
                        $this.LogDebug('Raw message', $gChatMessage)
                        # Get time of message
                        $unixEpoch = [datetime]'1970-01-01T00:00:00.0000Z'
                        $msg.Time = if ($gChatEvent.eventTime.seconds) {
                            $unixEpoch.AddSeconds($gChatEvent.eventTime.seconds)
                        }
                        else {
                            (Get-Date).ToUniversalTime()
                        }
                        if ($gChatEvent.type -eq 'REMOVED_FROM_SPACE') {
                            $messages.Add($msg) | Out-Null
                            $this.LoadRooms()
                        }
                        elseif ($gChatEvent.type -eq 'ADDED_TO_SPACE') {
                            $messages.Add($msg) | Out-Null
                            $this.LoadRooms()
                        }
                        else {
                            $messages.Add($msg) | Out-Null
                        }
                    } 
                    else {
                        $this.LogDebug("Message type is [$($gChatEvent.type)]. Ignoring and marking as complete")
                        $fullSheet = Import-GSSheet -SpreadsheetId $this.SheetId -SheetName $this.SheetName -Range "A1:D" -ErrorAction Stop
                        $fullSheetCount = if (!$fullSheet.Count) {
                            1
                        }
                        else {
                            $fullSheet.Count
                        }
                        for ($i = 0; $i -lt $fullSheetCount; $i++) {
                            if ($fullSheet[$i].Id -eq $gChatMessage.Id) {
                                break
                            }
                        }
                        $rowId = $i + 2
                        Export-GSSheet -SpreadsheetId $this.SheetId -Value "Yes" -SheetName $this.SheetName -Range "C$($rowId)" -ErrorAction Stop | Out-Null
                    }
                }
            }
        }
        catch {
            Write-Error $_
        }
        return $messages
    }

    # Send a GChat ping - (not really needed for this implementation)
    [void]Ping() { }

    # Send a message back to GChat
    [void]SendMessage([Response]$Response) {
        if (!$Script:_gChatAcked.Contains($Response.OriginalMessage.RawMessage.Id)) {
            if ((Show-PSGSuiteConfig).ConfigName -ne $this.Connection.ConfigName) {
                Switch-PSGSuiteConfig $this.Connection.ConfigName -Verbose
            }
            # Process any custom responses
            $this.LogVerbose("[$($Response.Data.Count)] custom responses and [$($Response.Text.Count)] text responses")
            $this.LogVerbose("Message Details :: [ConfigName:$($this.Connection.ConfigName) | SheetId:$($this.Connection.SheetId) | SheetName:$($this.Connection.SheetName) | PollingFrequency:$($this.Connection.PollingFrequency)]")
            foreach ($customResponse in $Response.Data) {

                [string]$sendTo = $Response.To
                if ($customResponse.DM) {
                    $rawMessageType = (ConvertFrom-Json $Response.OriginalMessage.RawMessage.Event).space.type
                    if ($rawMessageType -ne 'DM') {
                        $this.LogVerbose("Response is [DM] and original message space type is [$($rawMessageType)] - parsing UserID to DM Name")
                        $sendToHash = $this.UserIdToDMName("users/$($Response.MessageFrom)")
                        if ($sendToHash.ContainsKey('name')) {
                            $sendTo = $sendToHash['name']
                            $this.LogVerbose("UserID [$($Response.MessageFrom)] successfully parsed to DM Name [$sendTo]")
                            $respText = "<users/$($Response.MessageFrom)> The information you requested has been sent to you via Direct Message. Thank you!"
                            Send-GSChatMessage -Text $respText -Thread $Response.To -Parent "$($Response.To.Split("/")[0..1] -join "/")"
                        }
                        else {
                            $respText = "<users/$($Response.MessageFrom)> Your request was received, but the information requested is only available to be sent via Direct Message. Please open a Direct Message with me first then submit your command again. Thank you!"
                            Send-GSChatMessage -Text $respText -Thread $Response.To -Parent "$($Response.To.Split("/")[0..1] -join "/")"
                            break
                        }
                    }
                    else {
                        $this.LogVerbose("Response is [DM] and original message space type is [$($rawMessageType)] - no need to parse DM Name")
                    }
                }
                
                switch -Regex ($customResponse.PSObject.TypeNames[0]) {
                    '(.*?)PoshBot\.Card\.Response' {
                        $this.LogVerbose("Custom response is [$($customResponse.PSObject.TypeNames[0])]")
                        $sendParams = @{}
                        $fbText = ''
                        if ($customResponse.CustomData) {
                            $deserializedItem = try {
                                [System.Management.Automation.PSSerializer]::Deserialize($customResponse.CustomData)
                                $this.LogVerbose("CardResponse::CustomData :: Type [$($customResponse.CustomData.PSObject.TypeNames[0])] :: Succesfully deserialized", $customResponse.CustomData)
                            }
                            catch {
                                try {
                                    if ($customResponse.CustomData -is [System.Collections.Hashtable] -or $customResponse.CustomData -is [System.Management.Automation.PSCustomObject]) {
                                        $this.LogVerbose("CardResponse::CustomData :: Type [$($customResponse.CustomData.PSObject.TypeNames[0])] :: Item is already correct type", $customResponse.CustomData)
                                        $customResponse.CustomData
                                    }
                                    elseif ($jsonConvert = ConvertFrom-Json $customResponse.CustomData) {
                                        $this.LogVerbose("CardResponse::CustomData :: Type [$($customResponse.CustomData.PSObject.TypeNames[0])] :: Item is a JSON string, returning converted object", $customResponse.CustomData)
                                        $jsonConvert
                                    }
                                    else {
                                        $null
                                    }
                                }
                                catch {
                                    $null
                                }
                            }
                            if ($deserializedItem.token -and $deserializedItem.body) {
                                $this.LogVerbose("Deserialized Body", $deserializedItem.body)
                                $this.LogVerbose("Deserialized Token Present", $(if($deserializedItem.token){$true}else{$false}))
                                $deserBody = ConvertTo-Json -InputObject $deserializedItem.body -Depth 20
                                $restParams = @{
                                    ContentType = 'application/json'
                                    Verbose = $false
                                    Headers = @{
                                        Authorization = "Bearer $($deserializedItem.token)"
                                    }
                                    Body = $deserBody
                                }
                                $gChatResponse = if ($sendTo -like "spaces/*/messages/*") {
                                    $this.LogVerbose("Updating parsed message [$sendTo]")
                                    $updateMask = @()
                                    if ($deserializedItem.body.text) {
                                        $updateMask += 'text'
                                    }
                                    if ($deserializedItem.body.cards) {
                                        $updateMask += 'cards'
                                    }
                                    $restParams['Uri'] = ([Uri]"https://chat.googleapis.com/v1/$($sendTo)?updateMask=$($updateMask -join ',')")
                                    $restParams['Method'] = 'Put'
                                    Invoke-RestMethod @restParams
                                }
                                elseif ($sendTo -like "spaces/*/threads/*") {
                                    $this.LogVerbose("Sending parsed response to thread [$sendTo]")
                                    $deserializedItem.body | Add-Member -MemberType NoteProperty -Name thread -Value $(@{
                                        name = $sendTo
                                    }) -Force
                                    $newDeserBody = ConvertTo-Json -InputObject $deserializedItem.body -Depth 20
                                    $restParams['Body'] = $newDeserBody
                                    $updatedUri = "https://chat.googleapis.com/v1/$($sendTo.Split("/")[0..1] -join "/")/messages"
                                    $restParams['Uri'] = ([Uri]$updatedUri)
                                    $restParams['Method'] = 'Post'
                                    Invoke-RestMethod @restParams
                                }
                                else {
                                    $this.LogVerbose("Sending parsed message to space [$sendTo]")
                                    $restParams['Uri'] = ([Uri]"https://chat.googleapis.com/v1/$($sendTo)/messages")
                                    $restParams['Method'] = 'Post'
                                    Invoke-RestMethod @restParams
                                }
                            }
                            else {
                                $this.LogInfo([LogSeverity]::Warning, "Unable to parse Card's CustomData as a GChat response and token! SKIPPING", $customResponse.CustomData)
                            }
                        }
                        else {
                            $widgets = @()
                            if (-not [string]::IsNullOrEmpty($customResponse.Text)) {
                                $this.LogDebug("Response size [$($customResponse.Text.Length)]")
                                $sendParams.Text = $customResponse.Text
                                $fbText = $customResponse.Text
                            }
                            $sendParams.FallbackText = $fbText
                            if ($customResponse.Fields) {
                                $widgets += foreach ($key in $customResponse.Fields.Keys) {
                                    Add-GSChatKeyValue -TopLabel $key -Content $customResponse.Fields[$key] 
                                }
                            }
                            if ($customResponse.ImageUrl) {
                                $widgets += Add-GSChatImage -ImageUrl $customResponse.ImageUrl -LinkImage
                            }
                            if ($widgets) {
                                $cardParams = @{}
                                if ($customResponse.Title) {
                                    $cardParams.HeaderTitle = $customResponse.Title
                                }
                                $sendParams.MessageSegment = $widgets
                            }
                            $gChatResponse = if ($sendTo -like "spaces/*/messages/*") {
                                $this.LogVerbose("Updating message [$sendTo]", $sendParams)
                                Update-GSChatMessage @sendParams -MessageId $sendTo -Verbose:$false
                            }
                            elseif ($sendTo -like "spaces/*/threads/*") {
                                $this.LogVerbose("Sending response to thread [$sendTo]", $sendParams)
                                Send-GSChatMessage @sendParams -Thread $sendTo -Parent $($sendTo.Split("/")[0..1] -join "/") -Verbose:$false
                            }
                            else {
                                $this.LogVerbose("Sending message to space [$sendTo]", $sendParams)
                                Send-GSChatMessage @sendParams -Parent $sendTo -Verbose:$false
                            }
                        }
                        break
                    }
                    '(.*?)PoshBot\.Text\.Response' {
                        $this.LogVerbose("Custom response is [$($customResponse.PSObject.TypeNames[0])]")
                        $chunks = $this._ChunkString($customResponse.Text)
                        $i = 0
                        foreach ($chunk in $chunks) {
                            $t = if ($customResponse.AsCode) {
                                '```' + $chunk + '```'
                            } else {
                                $chunk
                            }
                            $gChatResponse = if ($sendTo -like "spaces/*/messages/*") {
                                $this.LogDebug("Updating message [$sendTo]", $t)
                                Update-GSChatMessage -MessageId $sendTo -Text $t -UpdateMask text -Verbose:$false
                            }
                            elseif ($sendTo -like "spaces/*/threads/*") {
                                $this.LogDebug("Sending response to thread [$sendTo]", $t)
                                Send-GSChatMessage -Text $t -Thread $sendTo -Parent $($sendTo.Split("/")[0..1] -join "/") -Verbose:$false
                            }
                            else {
                                $this.LogDebug("Sending message to space [$sendTo]", $t)
                                Send-GSChatMessage -Text $t -Parent $sendTo -Verbose:$false
                            }
                            $i++
                        }
                        break
                    }
                    '(.*?)PoshBot\.File\.Upload' {
                        $this.LogInfo([LogSeverity]::Error, "Custom response is [$($customResponse.PSObject.TypeNames[0])]. Google Chat does not currently support File Upload via API/SDK call.")
                        # TODO: Must build out once Google Chat supports it.
                        break
                    }
                    default {
                        $this.LogVerbose("Custom response is [$($customResponse.PSObject.TypeNames[0])]")
                    }
                }
            }
            if ($Response.Text.Count -gt 0) {
                [string]$sendTo = $Response.To
                if ($customResponse.DM) {
                    $sendToHash = "$($this.UserIdToDMName($Response.MessageFrom))"
                    if ($sendToHash.ContainsKey('name')) {
                        $sendTo = $sendToHash['name']
                    }
                }
                [string]$sentFrom = $Response.From
                $i = 0
                $total = $Response.Text.Count
                foreach ($item in $Response.Text) {
                    $i++
                    $deserializedItem = try {
                        [System.Management.Automation.PSSerializer]::Deserialize($item)
                        $this.LogVerbose("Text Item [$i/$total] :: Type [$($item.PSObject.TypeNames[0])] :: Succesfully deserialized", $item)
                    }
                    catch {
                        try {
                            if ($item -is [System.Collections.Hashtable] -or $item -is [System.Management.Automation.PSCustomObject]) {
                                $this.LogVerbose("Text Item [$i/$total] :: Type [$($item.PSObject.TypeNames[0])] :: Item is already correct type", $item)
                                $item
                            }
                            elseif ($jsonConvert = ConvertFrom-Json $item) {
                                $this.LogVerbose("Text Item [$i/$total] :: Type [$($item.PSObject.TypeNames[0])] :: Item is a JSON string, returning converted object", $item)
                                $jsonConvert
                            }
                            else {
                                $null
                            }
                        }
                        catch {
                            $null
                        }
                    }
                    if ($deserializedItem.token -and $deserializedItem.body) {
                        $this.LogVerbose("Deserialized Body", $deserializedItem.body)
                        $this.LogVerbose("Deserialized Token Present", $(if($deserializedItem.token){$true}else{$false}))
                        $deserBody = ConvertTo-Json -InputObject $deserializedItem.body -Depth 20
                        $restParams = @{
                            ContentType = 'application/json'
                            Verbose = $false
                            Headers = @{
                                Authorization = "Bearer $($deserializedItem.token)"
                            }
                            Body = $deserBody
                        }
                        $gChatResponse = if ($sendTo -like "spaces/*/messages/*") {
                            $this.LogVerbose("Updating parsed message [$sendTo]")
                            $updateMask = @()
                            if ($deserializedItem.body.text) {
                                $updateMask += 'text'
                            }
                            if ($deserializedItem.body.cards) {
                                $updateMask += 'cards'
                            }
                            $restParams['Uri'] = ([Uri]"https://chat.googleapis.com/v1/$($sendTo)?updateMask=$($updateMask -join ',')")
                            $restParams['Method'] = 'Put'
                            Invoke-RestMethod @restParams
                        }
                        elseif ($sendTo -like "spaces/*/threads/*") {
                            $this.LogVerbose("Sending parsed response to thread [$sendTo]")
                            $deserializedItem.body | Add-Member -MemberType NoteProperty -Name thread -Value $(@{
                                name = $sendTo
                            }) -Force
                            $newDeserBody = ConvertTo-Json -InputObject $deserializedItem.body -Depth 20
                            $restParams['Body'] = $newDeserBody
                            $updatedUri = "https://chat.googleapis.com/v1/$($sendTo.Split("/")[0..1] -join "/")/messages"
                            $restParams['Uri'] = ([Uri]$updatedUri)
                            $restParams['Method'] = 'Post'
                            Invoke-RestMethod @restParams
                        }
                        else {
                            $this.LogVerbose("Sending parsed message to space [$sendTo]")
                            $restParams['Uri'] = ([Uri]"https://chat.googleapis.com/v1/$($sendTo)/messages")
                            $restParams['Method'] = 'Post'
                            Invoke-RestMethod @restParams
                        }
                    }
                    else {
                        $chunks = $this._ChunkString($item)
                        foreach ($t in $chunks) {
                            $this.LogDebug("Sending response back to GChat channel [$($Response.To)]", $t)
                            $gChatResponse = if ($Response.To -like "spaces/*/messages/*") {
                                $this.LogDebug("Updating message [$($Response.To)]", $t)
                                Update-GSChatMessage -MessageId $Response.To -Text $t -UpdateMask text -Verbose:$false
                            }
                            elseif ($Response.To -like "spaces/*/threads/*") {
                                $this.LogDebug("Sending response to thread [$($Response.To)]", $t)
                                Send-GSChatMessage -Text $t -Thread $Response.To -Parent $($Response.To.Split("/")[0..1] -join "/") -Verbose:$false
                            }
                            else {
                                $this.LogDebug("Sending message to space [$($Response.To)]", $t)
                                Send-GSChatMessage -Text $t -Parent $Response.To -Verbose:$false
                            }
                        }
                    }
                }
            }
            $this.LogInfo([LogSeverity]::Warning,"Marking message Id $($Response.OriginalMessage.RawMessage.Id) as acknowledged")
            $Script:_gChatAcked.Add($Response.OriginalMessage.RawMessage.Id) | Out-Null
        }
        else {
            $this.LogInfo([LogSeverity]::Warning,"Skipping message Id $($Response.OriginalMessage.RawMessage.Id) | Message already tracked as complete.")
        }
    }

    # Add a reaction to an existing chat message
    [void]AddReaction([Message]$Message, [ReactionType]$Type, [string]$Reaction) {
        $this.LogDebug("Reactions are not yet supported in Google Chat - Ignoring")
        # TODO: Must build out once Google Chat supports it.
    }

    # Remove a reaction from an existing chat message
    [void]RemoveReaction([Message]$Message, [ReactionType]$Type, [string]$Reaction) {
        $this.LogDebug("Reactions are not yet supported in Google Chat - Ignoring")
        # TODO: Must build out once Google Chat supports it.
    }

    # Resolve a channel name to an Id
    [string]ResolveChannelId([string]$ChannelName) {
        if ($ChannelName -match '^#') {
            $ChannelName = $ChannelName.TrimStart('#')
        }
        $channelId = ($this.Connection.LoginData.channels | Where-Object name -eq $ChannelName).id
        if (-not $ChannelId) {
            $channelId = ($this.Connection.LoginData.channels | Where-Object id -eq $ChannelName).id
        }
        $this.LogDebug("Resolved channel [$ChannelName] to [$channelId]")
        return $channelId
    }

    # Populate the list of users the GChat team
    [void]LoadUsers() {
        $this.LogVerbose('Getting Google Chat users')
        $allUsers = Get-GSUser -Filter "isSuspended -eq '$false' changePasswordAtNextLogin -eq '$false'" -Verbose:$false
        $this.LogVerbose("[$($allUsers.Count)] users returned")
        $allUsers | ForEach-Object {
            $user = [GChatPerson]::new()
            $user.Id = "users/$($_.Id)"
            $user.NickName = $_.Name.FullName
            $user.FullName = $_.Name.FullName
            $user.FirstName = $_.Name.GivenName
            $user.LastName = $_.Name.FamilyName
            $user.Email = $_.PrimaryEmail
            $user.Phones = $_.Phones
            $user.IsAdmin = $_.IsAdmin
            $user.IsDelegatedAdmin = $_.IsDelegatedAdmin
            $user.IsEnforcedIn2Sv = $_.IsEnforcedIn2Sv
            $user.IsEnrolledIn2Sv = $_.IsEnrolledIn2Sv
            $user.OrgUnitPath = $_.OrgUnitPath
            $user.CreationTimeRaw = $_.CreationTimeRaw
            $user.CreationTime = $_.CreationTime
            $user.LastLoginTimeRaw = $_.LastLoginTimeRaw
            $user.LastLoginTime = $_.LastLoginTime
            $user.ThumbnailPhotoUrl = $_.ThumbnailPhotoUrl
            if (-not $this.Users.ContainsKey("users/$($_.Id)")) {
                $this.LogDebug("Adding user [users/$($_.Id):$($_.Name.FullName)]")
                $this.Users["users/$($_.Id)"] =  $user
            }
        }

        foreach ($key in $this.Users.Keys | Where-Object {($_ -replace 'users\/','') -notin $allUsers.Id}) {
            $this.LogDebug("Removing outdated user [$key]")
            $this.Users.Remove($key)
        }
    }

    # Populate the list of channels in the GChat team
    [void]LoadRooms() {
        $this.LogVerbose('Getting Google Chat spaces')
        $allChannels = Get-GSChatSpace -Verbose:$false
        $this.LogVerbose("[$($allChannels.Count)] spaces returned")

        $allChannels | ForEach-Object {
            $channel = [GChatChannel]::new()
            $channel.Id = $_.Name
            if ($_.DisplayName) {
                $channel.Name = $_.DisplayName
            }
            else {
                $channel.Name = "DM"
            }
            $channel.Type = $_.Type
            $channelMembers = Get-GSChatMember -Space $_.Name -Verbose:$false
            $channel.MemberCount = $channelMembers.Count
            foreach ($member in $channelMembers) {
                $channel.Members.Add($member, $null)
            }
            $this.LogDebug("Adding space: $($_.DisplayName):$($_.Name)")
            $this.Rooms[$_.Name] = $channel
        }

        foreach ($key in $this.Rooms.Keys | Where-Object {$_ -notin $allChannels.Name}) {
            $this.LogDebug("Removing outdated channel [$key]")
            $this.Rooms.Remove($key)
        }
    }

    # Get the bot identity Id
    [string]GetBotIdentity() {
        $id = $this.Connection.LoginData.self.id
        $this.LogVerbose("Bot identity is [$id]")
        return $id
    }

    # Determine if incoming message was from the bot
    [bool]MsgFromBot([string]$From) {
        $frombot = ($this.BotId -eq $From)
        if ($fromBot) {
            $this.LogDebug("Message is from bot [From: $From == Bot: $($this.BotId)]. Ignoring")
        } else {
            $this.LogDebug("Message is not from bot [From: $From <> Bot: $($this.BotId)]")
        }
        return $fromBot
    }

    # Get a user by their Id
    [GChatPerson]GetUser([string]$UserId) {
        $user = $this.Users[$UserId]
        if (-not $user) {
            $this.LogDebug([LogSeverity]::Warning, "User [$UserId] not found. Refreshing users")
            $this.LoadUsers()
            $user = $this.Users[$UserId]
        }

        if ($user) {
            $this.LogDebug("Resolved user [$UserId]", $user)
        } else {
            $this.LogDebug([LogSeverity]::Warning, "Could not resolve user [$UserId]")
        }
        return $user
    }

    [hashtable]GetUserInfo([string]$UserId) {
        if ($UserId -notlike "users/*") {
            $UserId = "users/$UserId"
        }
        $user = $this.Users[$UserId]
        if (-not $user) {
            $this.LogDebug([LogSeverity]::Warning, "User [$UserId] not found. Refreshing users")
            $this.LoadUsers()
            $user = $this.Users[$UserId]
        }

        if ($user) {
            $this.LogDebug("Resolved user [$UserId]", $user)
            return $user.ToHash()
        } else {
            $this.LogDebug([LogSeverity]::Warning, "Could not resolve user [$UserId]")
            return $null
        }
    }

    # Get a user Id by their name
    [string]UsernameToUserId([string]$Username) {
        $Username = $Username.TrimStart('@')
        $user = $this.Users.Values | Where-Object {$_.Nickname -eq $Username -or $_.Email -eq $Username -or $_.FullName -eq $Username}
        $id = $null
        if ($user) {
            $id = $user.Id
        } else {
            # User each doesn't exist or is not in the local cache
            # Refresh it and try again
            $this.LogDebug([LogSeverity]::Warning, "User [$Username] not found. Refreshing users")
            $this.LoadUsers()
            $user = $this.Users.Values | Where-Object {$_.Nickname -eq $Username -or $_.Email -eq $Username -or $_.FullName -eq $Username}
            if (-not $user) {
                $id = $null
            } else {
                $id = $user.Id
            }
        }
        if ($id) {
            $this.LogDebug("Resolved [$Username] to [$id]")
        } else {
            $this.LogDebug([LogSeverity]::Warning, "Could not resolve user [$Username]")
        }
        return $id
    }

    # Get a user name by their Id
    [string]UserIdToUsername([string]$UserId) {
        $name = $null
        if ((Get-GSChatConfig).Spaces.ContainsKey("$UserId")) {
            $name = $this.Users[$UserId].Nickname
        } else {
            $this.LogDebug([LogSeverity]::Warning, "User [$UserId] not found. Refreshing users")
            $this.LoadUsers()
            $name = $this.Users[$UserId].Nickname
        }
        if ($name) {
            $this.LogDebug("Resolved [$UserId] to [$name]")
        } else {
            $this.LogDebug([LogSeverity]::Warning, "Could not resolve user [$UserId]")
        }
        return $name
    }

    # Get a user name by their Id
    [hashtable]UserIdToDMName([string]$UserId) {
        $hash = @{}
        if ((Get-GSChatConfig).Spaces.ContainsKey($UserId)) {
            $hash['name'] = (Get-GSChatConfig).Spaces[$UserId]
        }
        if ($hash.ContainsKey('name')) {
            $this.LogDebug("Resolved [$UserId] to DM name [$($hash['name'])]")
        } 
        else {
            $this.LogDebug([LogSeverity]::Warning, "Could not resolve user [$UserId] to a DM. Advising user to DM the bot to initialize the space first.")
        }
        return $hash
    }

    # Get the channel name by Id
    [string]ChannelIdToName([string]$ChannelId) {
        $name = $null
        if ($this.Rooms.ContainsKey($ChannelId)) {
            $name = $this.Rooms[$ChannelId].Name
        } else {
            $this.LogDebug([LogSeverity]::Warning, "Channel [$ChannelId] not found. Refreshing channels")
            $this.LoadRooms()
            $name = $this.Rooms[$ChannelId].Name
        }
        if ($name) {
            $this.LogDebug("Resolved [$ChannelId] to [$name]")
        } else {
            $this.LogDebug([LogSeverity]::Warning, "Could not resolve channel [$ChannelId]")
        }
        return $name
    }

    # Break apart a string by number of characters
    hidden [System.Collections.ArrayList] _ChunkString([string]$Text) {
        $chunks = [regex]::Split($Text, "(?<=\G.{$($this.MaxMessageLength)})", [System.Text.RegularExpressions.RegexOptions]::Singleline)
        $this.LogDebug("Split response into [$($chunks.Count)] chunks")
        return $chunks
    }

    # Resolve a reaction type to an emoji
    hidden [string]_ResolveEmoji([ReactionType]$Type) {
        $emoji = [string]::Empty
        Switch ($Type) {
            'Success'        { return 'white_check_mark' }
            'Failure'        { return 'exclamation' }
            'Processing'     { return 'gear' }
            'Warning'        { return 'warning' }
            'ApprovalNeeded' { return 'closed_lock_with_key'}
            'Cancelled'      { return 'no_entry_sign'}
            'Denied'         { return 'x'}
        }
        return $emoji
    }

    # Translate formatted @mentions like @bod@domain.com into @devblackops
    hidden [string]_ProcessMentions([string]$Text) {
        $processed = $Text

        $mentions = $processed | Select-String -Pattern '(@\S*|@\S*)' -AllMatches | ForEach-Object {
            $_.Matches | ForEach-Object {
                [pscustomobject]@{
                    FormattedId = $_.Value
                    UnformattedId = $_.Value.TrimStart('<@').TrimEnd('>')
                }
            }
        }
        $mentions | ForEach-Object {
            if ($name = $this.UsernameToUserId($_.UnformattedId)) {
                $processed = $processed -replace $_.FormattedId, "<users/$($name)>"
                $this.LogDebug($processed)
            } else {
                $this.LogDebug([LogSeverity]::Warning, "Unable to translate @mention [$($_.FormattedId)] into a username")
            }
        }

        return $processed
    }
}

class GChatChannel : Room {
    [string]$Id
    [string]$Name
    [string]$Type
    [int]$MemberCount
}

class GChatConnection : Connection {

    [string]$ConfigName
    [String]$SheetId
    [string]$SheetName
    [int]$PollingFrequency
    [bool]$Connected
    [object]$ReceiveJob = $null

    GChatConnection([string]$ConfigName,[string]$SheetId,[string]$SheetName,[int]$PollingFrequency) {
        if ((Show-PSGSuiteConfig).ConfigName -ne $ConfigName) {
            Switch-PSGSuiteConfig -ConfigName $ConfigName
        }
        $this.ConfigName = $ConfigName
        $this.SheetId = $SheetId
        $this.SheetName = $SheetName
        $this.PollingFrequency = $PollingFrequency
    }

    # Connect to GChat and start receiving messages
    [void]Connect() {
        if ($null -eq $this.ReceiveJob -or $this.ReceiveJob.State -ne 'Running') {
            $this.LogDebug('Connecting to Google Sheet MQ')
            #$this.TestConnect()
            $this.StartReceiveJob()
        } else {
            $this.LogDebug([LogSeverity]::Warning, 'Receive job is already running')
        }
    }

    # Log in to GChat with the bot token and get a URL to connect to via websockets
    [void]TestConnect() { 
        try {
            if (Import-GSSheet -SpreadsheetId $this.SheetId -SheetName $this.SheetName -Range "A1" -ErrorAction Stop) {
                $this.LogVerbose("Connection to Sheet validated!")
            }
            else {
                $this.LogInfo([LogSeverity]::Error, 'Failed to connect to Sheet!')
            }
        } 
        catch {
            $this.LogInfo([LogSeverity]::Error, 'Failed to connect to Sheet!')
            throw $_
        }
    }

    # Setup the websocket receive job
    [void]StartReceiveJob() {
        $recv = {
            [CmdLetBinding()]
            Param
            (
                [parameter(Mandatory = $true,Position = 0)]
                $ConfigName,
                [parameter(Mandatory = $true,Position = 1)]
                $SheetId,
                [parameter(Mandatory = $true,Position = 2)]
                $SheetName,
                [parameter(Mandatory = $true,Position = 3)]
                $PollingFrequency
            )
            if (!(Get-Module PSGSuite)) {
                Import-Module PSGSuite -MinimumVersion "2.13.0" -Verbose:$false -Force
            }
            if ((Show-PSGSuiteConfig).ConfigName -ne $ConfigName) {
                Switch-PSGSuiteConfig -ConfigName $ConfigName
            }
            # Connect to Google Sheet MQ
            Write-Warning "[GChatBackend:ReceiveJob] Connecting to Google Sheet MQ at [$($SheetId)::$($SheetName)]"

            # Receive messages and put on output stream so the backend can read them
            while ($true) {
                $completeCount = 0
                Write-Verbose "[GChatBackend:ReceiveJob] Polling Sheet MQ for new messages"
                try {
                    if ($fullSheet = Import-GSSheet -SpreadsheetId $SheetId -SheetName $SheetName -ErrorAction Stop) {
                        $fullSheetCount = if ($fullSheet.Id[0] -eq 'Event') {
                            0
                        }
                        elseif (!$fullSheet.Count) {
                            1
                        }
                        else {
                            $fullSheet.Count
                        }
                        Write-Verbose "[GChatBackend:ReceiveJob] Received [$fullSheetCount] new messages"
                        for ($i = 0; $i -lt $fullSheetCount; $i++) {
                            if ($fullSheet[$i].Acked -eq "No") {
                                $message = $fullSheet[$i]
                                    $msg = ConvertFrom-Json -InputObject $message.Event
                                    Write-Verbose "Message ID [$($message.Id)] received from [$($msg.user.email)] with text [$($msg.message.argumentText)]. Sending serialized event JSON to Output stream"
                                    ConvertTo-Json -InputObject $message -Depth 20
                                    $rowId = $i + 2 - $completeCount
                                    Export-GSSheet -SpreadsheetId $SheetId -Value "Yes" -SheetName $SheetName -Range "C$($rowId)" -ErrorAction Stop | Out-Null
                                    $completeCount++
                            }
                        }
                    }
                }
                catch {
                    Write-Warning $_
                }
                finally {
                    Start-Sleep -Seconds $PollingFrequency
                }
            }
        }
        try {
            $this.LogVerbose("Starting Google Sheet MQ receive job [ConfigName:$($this.ConfigName) | SheetId:$($this.SheetId) | SheetName:$($this.SheetName) | PollingFrequency:$($this.PollingFrequency)]")
            $this.ReceiveJob = Start-Job -Name ReceiveSheetMessages -ScriptBlock $recv -ArgumentList $this.ConfigName,$this.SheetId,$this.SheetName,$this.PollingFrequency -ErrorAction Stop -Verbose
            $this.Connected = $true
            $this.Status = [ConnectionStatus]::Connected
            $this.LogInfo("Started Google Sheet MQ receive job [$($this.ReceiveJob.Id)]")
        } catch {
            $this.LogInfo([LogSeverity]::Error, "$($_.Exception.Message)", [ExceptionFormatter]::Summarize($_))
        }
    }

    # Read all available data from the job
    [string]ReadReceiveJob() {
        # Read stream info from the job so we can log them
        $infoStream = $this.ReceiveJob.ChildJobs[0].Information.ReadAll()
        $warningStream = $this.ReceiveJob.ChildJobs[0].Warning.ReadAll()
        $errStream = $this.ReceiveJob.ChildJobs[0].Error.ReadAll()
        $verboseStream = $this.ReceiveJob.ChildJobs[0].Verbose.ReadAll()
        $debugStream = $this.ReceiveJob.ChildJobs[0].Debug.ReadAll()
        foreach ($item in $infoStream) {
            $this.LogInfo($item.ToString())
        }
        foreach ($item in $warningStream) {
            $this.LogInfo([LogSeverity]::Warning, $item.ToString())
        }
        foreach ($item in $errStream) {
            $this.LogInfo([LogSeverity]::Error, $item.ToString())
        }
        foreach ($item in $verboseStream) {
            $this.LogVerbose($item.ToString())
        }
        foreach ($item in $debugStream) {
            $this.LogVerbose($item.ToString())
        }

        # The receive job stopped for some reason. Reestablish the connection if the job isn't running
        if ($this.ReceiveJob.State -ne 'Running') {
            $this.LogInfo([LogSeverity]::Warning, "Receive job state is [$($this.ReceiveJob.State)]. Attempting to reconnect...")
            Start-Sleep -Seconds 5
            $this.Connect()
        }

        if ($this.ReceiveJob.HasMoreData) {
            return $this.ReceiveJob.ChildJobs[0].Output.ReadAll()
        } else {
            return $null
        }
    }

    # Stop the receive job
    [void]Disconnect() {
        $this.LogInfo('Closing connection')
        if ($this.ReceiveJob) {
            $this.LogInfo("Stopping receive job [$($this.ReceiveJob.Id)]")
            $this.ReceiveJob | Stop-Job -Confirm:$false -PassThru | Remove-Job -Force -ErrorAction SilentlyContinue
        }
        $this.Connected = $false
        $this.Status = [ConnectionStatus]::Disconnected
    }
}

enum GChatMessageType {
    MESSAGE
    ADDED_TO_SPACE
    REMOVED_FROM_SPACE
    CARD_CLICKED
}

class GChatMessage : Message {

    [GChatMessageType]$MessageType = [GChatMessageType]::MESSAGE

    GChatMessage(
        [string]$To,
        [string]$From,
        [string]$Body = [string]::Empty
    ) {
        $this.To = $To
        $this.From = $From
        $this.Body = $Body
    }
}

class GChatPerson : Person {
    [string]$Id
    [string]$FullName
    [string]$FirstName
    [string]$LastName
    [string]$Email
    [string]$Phones
    [bool]$IsAdmin
    [bool]$IsDelegatedAdmin
    [bool]$IsEnforcedIn2Sv
    [bool]$IsEnrolledIn2Sv
    [string]$OrgUnitPath
    [string]$CreationTimeRaw
    [datetime]$CreationTime
    [string]$LastLoginTimeRaw
    [datetime]$LastLoginTime
    [string]$ThumbnailPhotoUrl

    [hashtable]ToHash() {
        $hash = @{}
        $this | Get-Member -MemberType Property | Foreach-Object {
            $hash.Add($_.Name, $this.($_.name))
        }
        return $hash
    }
}

# Functions included

function New-PoshBotGChatBackend {
    <#
    .SYNOPSIS
    Create a new instance of a Google Chat backend

    .DESCRIPTION
    Create a new instance of a Google Chat backend

    .PARAMETER Configuration
    The hashtable containing backend-specific properties on how to create the Google Chat backend instance.

    .EXAMPLE
    PS C:\> $backendConfig = @{Name = 'PSGSuiteBot'; ConfigName = 'domain1'; SheetId = '1H7mJoKfE8BGRnOSEF893JK032olstpyjOjNcO5sK3mjg'; SheetName = 'Queue'; PollingFrequency = 5}
    PS C:\> $backend = New-PoshBotGChatBackend -Configuration $backendConfig

    Create a Google Chat backend using the specified values
    
    .INPUTS
    Hashtable
    
    .OUTPUTS
    GChatBackend
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Scope='Function', Target='*')]
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('BackendConfiguration')]
        [hashtable[]]$Configuration
    )
    Process {
        foreach ($item in $Configuration) {
            if (-not $item.SheetId) {
                throw 'Configuration is missing [SheetId] parameter'
            } else {
                if (-not $item.ConfigName) {
                    $item['ConfigName'] = (Show-PSGSuiteConfig).ConfigName
                }
                if (-not $item.SheetName) {
                    $item['SheetName'] = 'Sheet1'
                }
                if (-not $item.PollingFrequency) {
                    $item['PollingFrequency'] = 5
                }
                Write-Verbose "Creating new GChat backend instance:`n$(([PSCustomObject]$item | Format-List * | Out-String).Trim())"
                $backend = [GChatBackend]::new($item.ConfigName,$item.SheetId,$item.SheetName,$item.PollingFrequency)
                if ($item.Name) {
                    $backend.Name = $item.Name
                }
                $backend
            }
        }
    }
}

Export-ModuleMember -Function 'New-PoshBotGChatBackend'

function New-PoshBotGChatCardResponse {
    <#
    .SYNOPSIS
    Create a new instance of a Google Chat backend

    .DESCRIPTION
    Create a new instance of a Google Chat backend

    .PARAMETER Configuration
    The hashtable containing backend-specific properties on how to create the Google Chat backend instance.

    .EXAMPLE
    Add-GSChatTextParagraph -Text "Guys...","We <b>NEED</b> to <i>stop</i> spending money on <b>chocolate</b>!" |
    Add-GSChatKeyValue -TopLabel "Chocolate Budget" -Content '$5.00' -Icon DOLLAR |
    Add-GSChatKeyValue -TopLabel "Actual Spending" -Content '$5,000,000!' -BottomLabel "WTF" -Icon AIRPLANE |
    Add-GSChatImage -ImageUrl "https://media.tenor.com/images/f78545a9b520ecf953578b4be220f26d/tenor.gif" -LinkImage |
    Add-GSChatCardSection -SectionHeader "Dollar bills, y'all" | 
    Add-GSChatButton -Text "Launch nuke" -OnClick (Add-GSChatOnClick -ActionMethodName launchNuke -ActionParameters @{decryptCodes = $true;callUN = $true}) | 
    Add-GSChatButton -Text "Unleash hounds" -OnClick (Add-GSChatOnClick -ActionMethodName unleashHounds) | 
    Add-GSChatCardSection -SectionHeader "What should we do?" | 
    Add-GSChatCardAction -ActionLabel "CardAction" -OnClick (Add-GSChatOnClick -Url "https://vaporshell.io/") |
    Add-GSChatCard -HeaderTitle "Makin' moves with" -HeaderSubtitle "DEM GOODIES" -OutVariable card |
    Add-GSChatTextParagraph -Text "This message sent by <b>PSGSuite</b>!" | 
    Add-GSChatCardSection -SectionHeader "Additional Info" | 
    New-PoshBotCardResponse -Text "Budget Report"

    Create a Google Chat backend using the specified values
    
    .OUTPUTS
    PoshBotCardResponse
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Scope='Function', Target='*')]
    [cmdletbinding()]
    param(
        [ValidateSet('Normal', 'Warning', 'Error')]
        [string]
        $Type = 'Normal',
        [switch]
        $DM,
        [string]
        $Text = [string]::empty,
        [string]
        $Title,
        [ValidateScript({
            $uri = $null
            if ([system.uri]::TryCreate($_, [System.UriKind]::Absolute, [ref]$uri)) {
                return $true
            } else {
                $msg = 'ThumbnailUrl must be a valid URL'
                throw [System.Management.Automation.ValidationMetadataException]$msg
            }
        })]
        [string]
        $ThumbnailUrl,
        [ValidateScript({
            $uri = $null
            if ([system.uri]::TryCreate($_, [System.UriKind]::Absolute, [ref]$uri)) {
                return $true
            } else {
                $msg = 'ImageUrl must be a valid URL'
                throw [System.Management.Automation.ValidationMetadataException]$msg
            }
        })]
        [string]
        $ImageUrl,
        [ValidateScript({
            $uri = $null
            if ([system.uri]::TryCreate($_, [System.UriKind]::Absolute, [ref]$uri)) {
                return $true
            } else {
                $msg = 'LinkUrl must be a valid URL'
                throw [System.Management.Automation.ValidationMetadataException]$msg
            }
        })]
        [string]
        $LinkUrl,
        [hashtable]
        $Fields,
        [ValidateScript({
            if ($_ -match '^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$') {
                return $true
            } else {
                $msg = 'Color but be a valid hexidecimal color code e.g. #008000'
                throw [System.Management.Automation.ValidationMetadataException]$msg
            }
        })]
        [string]
        $Color = '#D3D3D3',
        [object]
        $CustomData,
        [parameter(Mandatory = $false, ValueFromPipeline = $true)]
        [Object[]]
        $MessageSegment
    )
    Begin {
        $sendParams = @{}
        $finalSegment = @()
        $response = [ordered]@{
            PSTypeName = 'PoshBot.Card.Response'
            Type = $Type
            Text = $Text.Trim()
            DM = $PSBoundParameters['DM']
        }
        foreach ($key in $PSBoundParameters.Keys) {
            switch -Regex ($key) {
                'Text' {
                    $sendParams.Text = $Text.Trim()
                }
                '(Title|ThumbnailUrl|ImageUrl|LinkUrl|Fields|CustomData)' {
                    $response.$key = $PSBoundParameters[$key]
                }
            }
        }
        if (!$PSBoundParameters['Color']) {
            $response.Color = $Color
        }
        else {
            switch ($Type) {
                'Normal' {
                    $response.Color = '#008000'
                }
                'Warning' {
                    $response.Color = '#FFA500'
                }
                'Error' {
                    $response.Color = '#FF0000'
                }
            }
        }
    }
    Process {
        foreach ($segment in $MessageSegment) {
            $finalSegment += $segment
        }
    }
    End {
        if ($finalSegment) {
            $json = $finalSegment | Send-GSChatMessage -BodyPassThru @sendParams
            $response.CustomData = $json
        }
        [pscustomobject]$response
    }
}

Export-ModuleMember -Function 'New-PoshBotGChatCardResponse'