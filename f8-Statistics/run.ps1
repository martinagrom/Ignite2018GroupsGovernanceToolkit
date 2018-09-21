#-------------------------------------------
# onegroup.ps1
#-------------------------------------------
# checks properties, owners and members of an Office 365 Group and outputs the result to an Azure Table storage for use with Power BI.
# September 2018, atwork.at. Script by Christoph Wilfing, Martina Grom, Toni Pohl

#-------------------------------------------
# Get OneGroup
#-------------------------------------------
$onegroup = Get-Content $onegroup -Raw | ConvertFrom-Json

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

#-------------------------------------------
#region get group properties
#-------------------------------------------
try {
    $uri = "https://graph.microsoft.com/v1.0/groups/$($onegroup.GroupId)"
    Write-Host "query uri $uri"
    $grpresult = Invoke-RestMethod `
        -Method Get `
        -Uri $uri `
        -ContentType 'application/json' `
        -Headers $script:APIHeader
}
catch [System.Net.WebException] {
    Write-Output "onegroup group: WebException: $($_.exception)"
    Write-Output "onegroup group: ErrorCode: [$($_.Exception.Response.StatusCode.value__)]"
}
catch {
    Write-Output "onegroup: group Another Exception caught: [$($_.Exception)]"
}

#-------------------------------------------
#region check if it's a team
#-------------------------------------------
$isTeam = "No"
try {
    $uri = "https://graph.microsoft.com/beta/groups/$($onegroup.GroupId)/endpoints"
    Write-Host "query uri $uri"
    $teamresult = Invoke-RestMethod `
        -Method Get `
        -Uri $uri `
        -ContentType 'application/json' `
        -Headers $script:APIHeader

        # if we get an id, then the group is a team.
    #Write-Output "onegroup: isTeam: [$($teamresult.value.id)]"
    if ($teamresult.value.id) { $isTeam = "Yes" }
}
catch [System.Net.WebException] {
    Write-Output "onegroup team: WebException: $($_.exception)"
    Write-Output "onegroup team: ErrorCode: [$($_.Exception.Response.StatusCode.value__)]"
}
catch {
    Write-Output "onegroup team: Another Exception caught: [$($_.Exception)]"
}


#-------------------------------------------
#region check if the group has owners
#-------------------------------------------
try {
    $uri = "https://graph.microsoft.com/v1.0/groups/$($onegroup.GroupId)/owners"
    Write-Host "query uri $uri"
    $ownerresult = Invoke-RestMethod `
        -Method Get `
        -Uri $uri `
        -ContentType 'application/json' `
        -Headers $script:APIHeader
}
catch [System.Net.WebException] {
    Write-Output "onegroup owners: WebException: $($_.exception)"
    Write-Output "onegroup owners: ErrorCode: [$($_.Exception.Response.StatusCode.value__)]"
}
catch {
    Write-Output "onegroup owners: Another Exception caught: [$($_.Exception)]"
}

#-------------------------------------------
#region count the group members
#-------------------------------------------
try {
    # We (still) need to use the beta endpoint to get the userType...
    $uri = "https://graph.microsoft.com/beta/groups/$($onegroup.GroupId)/members"
    Write-Host "query uri $uri"
    $membersresult = Invoke-RestMethod `
        -Method Get `
        -Uri $uri `
        -ContentType 'application/json' `
        -Headers $script:APIHeader
}
catch [System.Net.WebException] {
    Write-Output "onegroup members: WebException: $($_.exception)"
    Write-Output "onegroup members: ErrorCode: [$($_.Exception.Response.StatusCode.value__)]"
}
catch {
    Write-Output "onegroup members: Another Exception caught: [$($_.Exception)]"
}

#-------------------------------------------
#region count the group members
#-------------------------------------------
$guests = 0
try {
    #$GuestList = New-Object -TypeName System.Collections.ArrayList
    Write-Output "onegroup: membersresult.value: $($membersresult.value)"
    foreach ($Member in $membersresult.value) {
        if ($Member.userType -eq 'Guest' ) { $guests++ }
    }
    Write-Output "onegroup: guests: $($guests)"
}
catch [System.Net.WebException] {
    Write-Output "onegroup: WebException: $($_.exception)"
    Write-Output "onegroup: ErrorCode: [$($_.Exception.Response.StatusCode.value__)]"
}
catch {
    Write-Output "onegroup: Another Exception caught: [$($_.Exception)]"
}

#-------------------------------------------
#region Write the result
#-------------------------------------------
# prettify
$classification = "None"
if ($grpresult.classification) { $classification = $grpresult.classification }
$visibility = "None"
if ($grpresult.visibility) { $visibility = $grpresult.visibility }
$renewedDateTime = "None"
if ($grpresult.renewedDateTime) { $renewedDateTime = $grpresult.renewedDateTime.SubString(0,7) }


$TableEntry = [PSObject]@{
    PartitionKey     = 'groupstatistics'
    RowKey           = $(new-guid).Guid
    GroupDisplayName = $onegroup.GroupDisplayName
    GroupId          = $onegroup.GroupId
    GroupMail        = $onegroup.GroupMail
    GroupOwnerCount  = $ownerresult.value.count
    MembersCount     = $membersresult.value.count
    GuestsCount      = $guests
    Classification   = $classification
    Visibility       = $visibility
    RenewedDateTime  = $renewedDateTime
    isTeam           = $isTeam
}

$TableEntry | ConvertTo-Json | Out-File -FilePath $groupsstatistics -Encoding utf8
#endregion
# Done with that group.
