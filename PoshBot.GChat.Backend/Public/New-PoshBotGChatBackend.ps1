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
                    $item['SheetName'] = 'Queue'
                }
                if (-not $item.PollingFrequency) {
                    $item['PollingFrequency'] = 1500
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
