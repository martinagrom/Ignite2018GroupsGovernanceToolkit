#-------------------------------------------
# CreateTeam.ps1
#-------------------------------------------
# provisions a Team out of an Office 365 Group
# September 2018, atwork.at. Script by Christoph Wilfing, Martina Grom, Toni Pohl

$TeamInfo = Get-Content $createteam -Raw | ConvertFrom-Json

# 09/20/2018 - Teams team updated the API in the production environment, now this works (again).
# There must be a user set as owner of the group that it works.
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
    $result = Invoke-RestMethod -Method Post `
                        -Uri $Authority `
                        -ContentType 'application/x-www-form-urlencoded' `
                        -Body $body

    $script:APIHeader = @{'Authorization' = "Bearer $($result.access_token)" }
}

# Initialize Authorization
Initialize-Authorization -TenantID $env:TenantID -ClientKey $env:AppSecret -AppID $env:AppID

try {
    Write-Output "creating team on $($TeamInfo.GroupId)"
    $result = Invoke-RestMethod `
                            -Method Put `
                            -Uri "https://graph.microsoft.com/beta/groups/$($TeamInfo.GroupId)/team" `
                            -ContentType 'application/json' `
                            -Headers $script:APIHeader `
                            -Body "{}" `
                            -ErrorAction Stop

    $TeamResult = $result.id
    Write-Output "Team created: ID:[$TeamResult]"
} catch [System.Net.WebException] {
    Write-Output "EnableTeamOnExistingGroup: $($TeamResult.error.code) => $($TeamResult.error.message)"
    throw "EnableTeamOnExistingGroup: $($TeamResult.error.code) => $($TeamResult.error.message)"
}
catch {
    Write-Output "EnableTeamOnExistingGroup: Another Exception caught: [$($_.Exception)]"
    throw "EnableTeamOnExistingGroup: $($TeamResult.error.code) => $($TeamResult.error.message)"
}
