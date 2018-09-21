#-------------------------------------------
# Ownerless.ps1
#-------------------------------------------
# checks if a given Office 365 Group has at least one owner.
# If there is no owner, the group is added to the Azure Table storage ownerless.
# September 2018, atwork.at. Script by Christoph Wilfing, Martina Grom, Toni Pohl

#-------------------------------------------
# dequeue
#-------------------------------------------
$ownerless = Get-Content $ownerless -Raw | ConvertFrom-Json

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
    $result = Invoke-RestMethod -Method Post `
        -Uri $Authority `
        -ContentType 'application/x-www-form-urlencoded' `
        -Body $body

    $script:APIHeader = @{'Authorization' = "Bearer $($result.access_token)" }
}

#-------------------------------------------
# Initialize Authorization
#-------------------------------------------
Initialize-Authorization -TenantID $env:TenantID -ClientKey $env:AppSecret -AppID $env:AppID
#endregion

#-------------------------------------------
#region check if the group has owners
#-------------------------------------------
try {
    $uri = "https://graph.microsoft.com/v1.0/groups/$($ownerless.GroupId)/owners"
    Write-Host "query uri $uri"
    $result = Invoke-RestMethod `
        -Method Get `
        -Uri $uri `
        -ContentType 'application/json' `
        -Headers $script:APIHeader
}
catch [System.Net.WebException] {
    Write-Output "ownerless: WebException: $($_.exception)"
    Write-Output "ownerless: ErrorCode: [$($_.Exception.Response.StatusCode.value__)]"
}
catch {
    Write-Output "ownerless: Another Exception caught: [$($_.Exception)]"
}

#-------------------------------------------
# our organization's policy:
#-------------------------------------------
if ($result.value.count -le 2) {
    # if the group has not 2 owners, so write it to our table
    Write-Output "ownerless: This group has less than two owners: $($ownerless.GroupDisplayName), $($ownerless.GroupId)"
    $TableEntry = [PSObject]@{
        PartitionKey     = 'ownerless'
        RowKey           = $(new-guid).Guid
        GroupDisplayName = $ownerless.GroupDisplayName
        GroupId          = $ownerless.GroupId
        GroupMail        = $ownerless.GroupMail
    }

    $TableEntry | ConvertTo-Json |  Out-File -FilePath $groupsownerless -Encoding utf8
}
#endregion
