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