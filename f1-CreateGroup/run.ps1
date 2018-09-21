#-------------------------------------------
# CreateGroup.ps1
#-------------------------------------------
# provisions an Office 365 group through an application in the tenant
# September 2018, atwork.at. Script by Christoph Wilfing, Martina Grom, Toni Pohl

#-------------------------------------------
# Get the required data from the body
#-------------------------------------------
$GroupInfo = Get-Content $creategroup -Raw | ConvertFrom-Json

#-------------------------------------------
#region Authorization
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

    $Authority = "https://login.microsoftonline.com/$TenantID/oauth2/token"

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
#region Create the Office 365 group
#-------------------------------------------
# Make email address safe...
$EMail = $GroupInfo.GroupName.Trim().Replace(" ","-")

$Body = [PSCustomObject]@{
    "displayname"     = "$($GroupInfo.GroupName)"
    "mailenabled"     = $true
    "mailnickname"    = "$($EMail)"
    "securityenabled" = $false
    "description"     = "$($GroupInfo.GroupName)"
    "groupTypes"      = @('Unified')
}

#-------------------------------------------
# add additional properties as neccessary
#-------------------------------------------
if ($GroupInfo.visibility) {
    Add-Member -InputObject $Body -MemberType NoteProperty -Name 'visibility' -Value $GroupInfo.visibility
}

$BodyJson = $Body | ConvertTo-Json
Write-Output "body: $BodyJson"

try {
    $result = Invoke-RestMethod `
        -Method Post `
        -Uri "https://graph.microsoft.com/v1.0/groups" `
        -ContentType 'application/json' `
        -Headers $script:APIHeader `
        -Body $BodyJson `
        -ErrorAction Stop
    
    # and save the generated Group ID
    $GroupID = $result.id
    Write-Output "Group created: ID:[$GroupID]"
    Out-File -InputObject $GroupID -FilePath $output
}
catch [System.Net.WebException] {
    $result = $_.Exception.Response.GetResponseStream()
    $reader = New-Object System.IO.StreamReader($result)
    $responseBody = $reader.ReadToEnd() | ConvertFrom-Json
    
    throw "CreateGroup: ErrorCode: [$($responseBody.error.message)] | [$($responseBody.error.details.ToString())]"
}
catch {
    Write-Output "CreateGroup: Another Exception caught: [$($_.Exception)]"
    throw "CreateGroup: Another Exception caught: [$($_.Exception)]"
}

#endregion

#-------------------------------------------
#region Add Classification
#-------------------------------------------
if ($GroupInfo.classification) {
    $body = [PSCustomObject]@{
        "classification" = "$($GroupInfo.classification)"
    }
    $BodyJson = $Body | ConvertTo-Json
    try {
        $result = Invoke-RestMethod `
            -Method Patch `
            -Uri "https://graph.microsoft.com/beta/groups/$GroupID" `
            -ContentType 'application/json' `
            -Headers $script:APIHeader `
            -Body $BodyJson `
            -ErrorAction Stop
        
        Write-Output "Group classification set"
    }
    catch [System.Net.WebException] {
        $result = $_.Exception.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($result)
        $responseBody = $reader.ReadToEnd() | ConvertFrom-Json
        
        throw "CreateGroup: ErrorCode: [$($responseBody.error.message)] | [$($responseBody.error.details.ToString())]"
    }
    catch {
        Write-Output "CreateGroup: Another Exception caught: [$($_.Exception)]"
        throw "CreateGroup: Another Exception caught: [$($_.Exception)]"
    }
}
#endregion

#-------------------------------------------
#region get the user by UPN
#-------------------------------------------
try {
    $result = Invoke-RestMethod `
        -Method get `
        -Uri "https://graph.microsoft.com/v1.0/users?`$filter=startswith(userPrincipalName,'$($GroupInfo.owner)')" `
        -ContentType 'application/json' `
        -Headers $script:APIHeader `
        -ErrorAction Stop

    $UserID = $result.value.id
    Write-Output "user found: ID:[$($result.value)]"
}
catch [System.Net.WebException] {
    Write-Output "AddGroupOwner: WebException: $($_.exception)"
    Write-Output "AddGroupOwner: ErrorCode: [$($_.Exception.Response.StatusCode.value__)]"
    Write-Output "AddGroupOwner: $($result.error.code) => $($result.error.message)"
    throw "AddGroupOwner: $($result.error.code) => $($result.error.message)"
}
catch {
    Write-Output "AddGroupOwner: Another Exception caught: [$($_.Exception)]"
    throw "AddGroupOwner: $($result.error.code) => $($result.error.message)"
}

#-------------------------------------------
#region add owner and member
#-------------------------------------------
$user = [PSCustomObject]@{
    '@odata.id' = "https://graph.microsoft.com/beta/users/$UserID"
}
$userJson = $user | ConvertTo-Json
Write-output "UserJson: $userJson"

try {
    $result = Invoke-RestMethod `
        -Method Post `
        -Uri "https://graph.microsoft.com/beta/groups/$GroupID/owners/`$ref" `
        -ContentType 'application/json' `
        -Headers $script:APIHeader `
        -Body $userJson `
        -ErrorAction Stop
                           
    $result = Invoke-RestMethod `
        -Method Post `
        -Uri "https://graph.microsoft.com/beta/groups/$GroupID/members/`$ref" `
        -ContentType 'application/json' `
        -Headers $script:APIHeader `
        -Body $userJson `
        -ErrorAction Stop

    $groupresult = $result.id
    Write-Output "group owner added"
}
catch [System.Net.WebException] {
    Write-Output "AddGroupOwner: WebException: $($_.exception)"
    Write-Output "AddGroupOwner: ErrorCode: [$($_.Exception.Response.StatusCode.value__)]"
    Write-Output "AddGroupOwner: $($groupresult.error.code) => $($groupresult.error.message)"
    throw "AddGroupOwner: $($groupresult.error.code) => $($groupresult.error.message)"
}
catch {
    Write-Output "AddGroupOwner: Another Exception caught: [$($_.Exception)]"
    throw "AddGroupOwner: $($groupresult.error.code) => $($groupresult.error.message)"
}
#endregion

#-------------------------------------------
#region Create Team on top of the group
#-------------------------------------------
Write-Output "Enable Team: $($GroupInfo.enableteam.trim().tolower())"
if ($GroupInfo.enableteam.trim().tolower() -eq 'yes') {
    $TeamInfo = [PSCustomObject]@{
        GroupID = $GroupID
    }
    $TeamInfoJson = $TeamInfo | ConvertTo-Json
    Write-Output "Forwarded TeamInfo: $TeamInfoJson"
    Out-File -FilePath $createteam -InputObject $TeamInfoJson
}
# Done, return. 
