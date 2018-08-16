function New-PoshBotGChatCardResponse {
    <#
    .SYNOPSIS
        Tells PoshBot to send a specially formatted response. Also includes pipeline support for Google Chat message segment and can be used as a direct swap-in for Send-GSChatMessage
    .DESCRIPTION
        Responses from PoshBot commands can either be plain text or formatted. Returning a response with New-PoshBotRepsonse will tell PoshBot
        to craft a specially formatted message when sending back to the chat network.
    .PARAMETER Type
        Specifies a preset color for the card response. If the [Color] parameter is specified as well, it will override this parameter.
        | Type    | Color  | Hex code |
        |---------|--------|----------|
        | Normal  | Greed  | #008000  |
        | Warning | Yellow | #FFA500  |
        | Error   | Red    | #FF0000  |
    .PARAMETER Text
        The text response from the command.
    .PARAMETER DM
        Tell PoshBot to redirect the response to a DM channel.
    .PARAMETER Title
        The title of the response. This will be the card title in chat networks like Slack.
    .PARAMETER ThumbnailUrl
        A URL to a thumbnail image to display in the card response.
    .PARAMETER ImageUrl
        A URL to an image to display in the card response.
    .PARAMETER LinkUrl
        Will turn the title into a hyperlink
    .PARAMETER Fields
        A hashtable to display as a table in the card response.
    .PARAMETER COLOR
        The hex color code to use for the card response. In Slack, this will be the color of the left border in the message attachment.
    .PARAMETER CustomData
        Any additional custom data you'd like to pass on. Useful for custom backends, in case you want to pass a specifically formatted response
        in the Data stream of the responses received by the backend. Any data sent here will be skipped by the built-in backends provided with PoshBot itself.
    .PARAMETER MessageSegment
        Google Chat message segments sent through the pipeline.
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
        New-PoshBotGChatCardResponse -Text "Budget Report" -DM

        Create a Google Chat card response
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
