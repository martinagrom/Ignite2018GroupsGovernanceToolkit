# guests.ps1
# checks if a member of the Office 365 Group is a guest user.
# If yes, create one queue message per owner and send an email message (only if we see at least one guest).
# September 2018, atwork.at. Script by Christoph Wilfing, Martina Grom, Toni Pohl

#-------------------------------------------
# Get all groups to check
#-------------------------------------------
$guests = Get-Content $externalguests -Raw | ConvertFrom-Json

#-------------------------------------------
#region Initialize Authorization
#-------------------------------------------
function Initialize-Authorization {
    param
    (
        [string]
        $ResourceURL = 'https://graph.microsoft.com',
  
        [string]
        [parameter(Mandatory)]
        $TenantID,
      
        [string]
        [Parameter(Mandatory)]
        $ClientKey,
  
        [string]
        [Parameter(Mandatory)]
        $AppID
    )

    $Authority = "https://login.windows.net/$TenantID/oauth2/token"

    [Reflection.Assembly]::LoadWithPartialName("System.Web") | Out-Null
    $EncodedKey = [System.Web.HttpUtility]::UrlEncode($ClientKey)

    $body = "grant_type=client_credentials&client_id=$AppID&client_secret=$EncodedKey&resource=$ResourceUrl"

    # Request a Token from the graph api
    $ownerresult = Invoke-RestMethod -Method Post `
        -Uri $Authority `
        -ContentType 'application/x-www-form-urlencoded' `
        -Body $body

    $script:APIHeader = @{'Authorization' = "Bearer $($ownerresult.access_token)" }
}

#-------------------------------------------
# Initialize Authorization
#-------------------------------------------
Initialize-Authorization -TenantID $env:TenantID -ClientKey $env:AppSecret -AppID $env:AppID
#endregion

#region get owners of the group
Write-Output "guests: query owners"
try {
    $uri = "https://graph.microsoft.com/beta/groups/$($guests.GroupID)/owners"
    Write-output  "query uri $uri"
    $owners = Invoke-RestMethod `
        -Method Get `
        -Uri $uri `
        -ContentType 'application/json' `
        -Headers $script:APIHeader `

}
catch [System.Net.WebException] {
    Write-Output "guests: WebException: $($_.exception)"
    Write-Output "guests: ErrorCode: [$($_.Exception.Response.StatusCode.value__)]"
}
catch {
    Write-Output "guests: Another Exception caught: [$($_.Exception)]"
}

if ($owners.value.count -eq 0) {
    Write-Output "guests: this group does not have an owner: $($guests.GroupID)"
}

#endregion

#-------------------------------------------
#region get members of a group
#-------------------------------------------
Write-Output "guests: query members"
try {
    # We (still) need to use the beta endpoint to get the userType...
    $uri = "https://graph.microsoft.com/beta/groups/$($guests.GroupId)/members"
    Write-output "query uri $uri"
    $Members = Invoke-RestMethod `
        -Method Get `
        -Uri $uri `
        -ContentType 'application/json' `
        -Headers $script:APIHeader `

}
catch [System.Net.WebException] {
    Write-Output "guests: WebException: $($_.exception)"
    Write-Output "guests: ErrorCode: [$($_.Exception.Response.StatusCode.value__)]"
}
catch {
    Write-Output "guests: Another Exception caught: [$($_.Exception)]"
}
#endregion

#-------------------------------------------
#region Cycle through all members and create queue message
#-------------------------------------------
$GuestList = New-Object -TypeName System.Collections.ArrayList
foreach ($Member in $Members.value) {

    #-------------------------------------------
    # if we want to protocol all members into a table...
    #-------------------------------------------
    <#
    $TableEntry = [PSObject]@{
        PartitionKey     = 'member'
        RowKey           = $(new-guid).Guid
        GroupDisplayName = $guests.GroupDisplayName
        GroupId          = $guests.GroupId
        GroupMail        = $guests.GroupMail
        Member           = $Member.Displayname
        MemberUPN        = $Member.Userprincipalname
        MemberType       = $Member.userType
        MemberMail       = $Member.Mail
    }
    Write-Output "guests: $TableEntry"
    $TableEntry | ConvertTo-Json |  Out-File -FilePath $groupsmembers -Encoding utf8
    #>
    
    #-------------------------------------------
    # add guest to the array
    #-------------------------------------------
    if ($Member.userType -eq 'Guest' ) {
        $GuestEntry = [PSObject]@{
            Member     = $Member.Displayname
            MemberUPN  = $Member.Userprincipalname
            MemberType = $Member.UserType
            MemberMail = $Member.Mail
        }
        $GuestList += $GuestEntry
        Write-Output  "guests: $($guests.GroupDisplayName) => Guest Member: $($Member.Displayname) | $($Member.userprincipalname)"

        $TableEntry = [PSObject]@{
            PartitionKey     = 'guest'
            RowKey           = $(new-guid).Guid
            GroupDisplayName = $guests.GroupDisplayName
            GroupId          = $guests.GroupId
            GroupMail        = $guests.GroupMail
            Member           = $Member.Displayname
            MemberUPN        = $Member.Userprincipalname
            MemberType       = $Member.userType
            MemberMail       = $Member.Mail
        }
        Write-Output "guests: $TableEntry"
        $TableEntry | ConvertTo-Json |  Out-File -FilePath $groupsguests -Encoding utf8
    }
}

#-------------------------------------------
# Are there guests in the Array?
#-------------------------------------------
if ($GuestList.count -gt 0) {

    foreach ($Owner in $owners.value) {
        # create a table entry just to have it stored somewhere (e.g. for later use in PowerBI...)
        $OwnerEntry = [PSObject]@{
            PartitionKey     = 'guests'
            RowKey           = $(new-guid).Guid
            OwnerMail        = $Owner.mail
            OwnerDisplayName = $Owner.Displayname
            GroupDisplayName = $guests.GroupDisplayName
            GroupID          = $guests.GroupID
            GroupMail        = $guests.GroupMail
            Guests           = $GuestList
        }
        $OwnerEntryJson = $OwnerEntry | ConvertTo-Json
        Write-Output "guests: $OwnerEntryJson"
        $OwnerEntryJson | Out-File -FilePath $groupsguests -Encoding utf8

        # create one queue message per Owner and send the message (only if we see at least one guest)
        $OwnerQueueEntry = [PSObject]@{
            OwnerMail        = $Owner.mail
            OwnerDisplayName = $Owner.Displayname
            GroupDisplayName = $guests.GroupDisplayName
            GroupID          = $guests.GroupID
            GroupMail        = $guests.GroupMail
            Guests           = $GuestList            
        }
        $OwnerQueueEntryJson = $OwnerQueueEntry | ConvertTo-Json
        Write-Output "guests: $OwnerQueueEntryJson"
        $OwnerQueueEntryJson | Out-File -FilePath $sendownermessage -Encoding utf8
    }
}
else {
    Write-Output "guests: group $($guests.GroupDisplayName) does not have any guests"
}