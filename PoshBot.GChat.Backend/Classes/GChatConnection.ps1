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
                    $sleepParams = @{}
                    if ($PollingFrequency -ge 1000) {
                        $sleepParams['Milliseconds'] = $PollingFrequency
                    }
                    else {
                        $sleepParams['Seconds'] = $PollingFrequency
                    }
                    Start-Sleep @sleepParams
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
