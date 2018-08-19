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
                            $this.LogVerbose("The response includes CustomData! Parsing...")
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
                                $this.LogVerbose("Deserialized Token Present", $($null -ne $deserializedItem.token))
                                $deserBody = ConvertTo-Json -InputObject $deserializedItem.body -Depth 20
                                $restParams = @{
                                    ContentType = 'application/json'
                                    Verbose = $false
                                    Headers = @{
                                        Authorization = "Bearer $($deserializedItem.token)"
                                    }
                                    Body = $deserBody
                                }
                                if ($sendTo -like "spaces/*/messages/*") {
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
                                }
                                else {
                                    $this.LogVerbose("Sending parsed message to space [$sendTo]")
                                    $restParams['Uri'] = ([Uri]"https://chat.googleapis.com/v1/$($sendTo)/messages")
                                    $restParams['Method'] = 'Post'
                                }
                                Invoke-RestMethod @restParams
                            }
                            else {
                                $this.LogInfo([LogSeverity]::Warning, "Unable to parse Card's CustomData as a GChat response and token! SKIPPING", $customResponse.CustomData)
                            }
                        }
                        else {
                            $this.LogVerbose("The response DOES NOT include CustomData! Parsing PoshBot CardResponse to Google Chat Card object...")
                            $widgets = @()
                            if (-not [string]::IsNullOrEmpty($customResponse.Text)) {
                                $this.LogDebug("Response size [$($customResponse.Text.Length)]")
                                $formattedText = if ($customResponse.LinkUrl) {
                                    "<$($customResponse.LinkUrl)|$($customResponse.Text)>"
                                }
                                else {
                                    $customResponse.Text
                                }
                                $sendParams.Text = $formattedText
                                $fbText = $customResponse.Text
                            }
                            elseif ($customResponse.LinkUrl) {
                                $sendParams.Text = "<$($customResponse.LinkUrl)|View Details>"
                                $fbText = $customResponse.LinkUrl
                            }
                            $sendParams.FallbackText = $fbText
                            if ($customResponse.ThumbnailUrl) {
                                $widgets += Add-GSChatImage -ImageUrl $customResponse.ThumbnailUrl -LinkImage
                            }
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
                                $card = $widgets | Add-GSChatCard @cardParams
                                $sendParams.MessageSegment = $card
                            }
                            if ($sendTo -like "spaces/*/messages/*") {
                                $this.LogVerbose("Updating message [$sendTo]", $sendParams)
                                try {
                                    Update-GSChatMessage @sendParams -MessageId $sendTo -Verbose:$false -ErrorAction Stop
                                }
                                catch {
                                    $this.LogInfo([LogSeverity]::Error, $_.Exception.Message, $_)
                                }
                            }
                            elseif ($sendTo -like "spaces/*/threads/*") {
                                $this.LogVerbose("Sending response to thread [$sendTo]", $sendParams)
                                try {
                                    Send-GSChatMessage @sendParams -Thread $sendTo -Parent $($sendTo.Split("/")[0..1] -join "/") -Verbose:$false -ErrorAction Stop
                                }
                                catch {
                                    $this.LogInfo([LogSeverity]::Error, $_.Exception.Message, $_)
                                }
                            }
                            else {
                                $this.LogVerbose("Sending message to space [$sendTo]", $sendParams)
                                try {
                                    Send-GSChatMessage @sendParams -Parent $sendTo -Verbose:$false -ErrorAction Stop
                                }
                                catch {
                                    $this.LogInfo([LogSeverity]::Error, $_.Exception.Message, $_)
                                }
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
                            if ($sendTo -like "spaces/*/messages/*") {
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
                        if ($sendTo -like "spaces/*/messages/*") {
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
                            if ($Response.To -like "spaces/*/messages/*") {
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
