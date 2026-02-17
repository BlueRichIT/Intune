<#
.SYNOPSIS
  Sync a (static) Entra device group based on users in two (static) user groups.

.REQUIREMENTS
  Microsoft Graph app permissions (Application):
    - Group.Read.All
    - GroupMember.Read.All
    - Directory.Read.All
    - GroupMember.ReadWrite.All  (to add/remove members in target group)
  Use client secret or certificate auth.

.NOTES
  Dynamic device group can't do this natively; this script maintains a static device group instead.
#>

param(
    [Parameter(Mandatory)] [string] $TenantId,
    [Parameter(Mandatory)] [string] $ClientId,
    [Parameter(Mandatory)] [string] $ClientSecret,

    [Parameter(Mandatory)] [string] $UserGroupId1,
    [Parameter(Mandatory)] [string] $UserGroupId2,

    [Parameter(Mandatory)] [string] $TargetDeviceGroupId
)

$ErrorActionPreference = "Stop"
$GraphBase = "https://graph.microsoft.com/v1.0"

function Get-GraphToken {
    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "https://graph.microsoft.com/.default"
        grant_type    = "client_credentials"
    }
    (Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body $body).access_token
}

function Invoke-Graph {
    param(
        [Parameter(Mandatory)] [string] $Method,
        [Parameter(Mandatory)] [string] $Uri,
        $Body
    )

    $headers = @{ Authorization = "Bearer $script:Token" }

    try {
        if ($Body) {
            return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -ContentType "application/json" -Body ($Body | ConvertTo-Json -Depth 10)
        } else {
            return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers
        }
    }
    catch {
        # Basic throttle handling
        if ($_.Exception.Response -and $_.Exception.Response.StatusCode.value__ -eq 429) {
            Start-Sleep -Seconds 5
            if ($Body) {
                return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -ContentType "application/json" -Body ($Body | ConvertTo-Json -Depth 10)
            } else {
                return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers
            }
        }
        throw
    }
}

function Get-AllPages($firstUrl) {
    $items = @()
    $url = $firstUrl
    while ($url) {
        $res = Invoke-Graph -Method GET -Uri $url
        if ($res.value) { $items += $res.value }
        $url = $res.'@odata.nextLink'
    }
    $items
}

$script:Token = Get-GraphToken

Write-Host "1) Getting users from both groups..."
$users1 = Get-AllPages "$GraphBase/groups/$UserGroupId1/transitiveMembers/microsoft.graph.user?`$select=id,userPrincipalName"
$users2 = Get-AllPages "$GraphBase/groups/$UserGroupId2/transitiveMembers/microsoft.graph.user?`$select=id,userPrincipalName"

$userMap = @{}
foreach ($u in ($users1 + $users2)) { $userMap[$u.id] = $true }
$userIds = $userMap.Keys
Write-Host "   Total unique users: $($userIds.Count)"

Write-Host "2) Getting devices for each user (registeredDevices)..."
$desiredDeviceIds = New-Object System.Collections.Generic.HashSet[string]

foreach ($uid in $userIds) {
    $devs = Get-AllPages "$GraphBase/users/$uid/registeredDevices?`$select=id"
    foreach ($d in $devs) {
        if ($d.id) { [void]$desiredDeviceIds.Add($d.id) }
    }
}

Write-Host "   Total desired devices: $($desiredDeviceIds.Count)"

Write-Host "3) Reading current members of target device group..."
$currentMembers = Get-AllPages "$GraphBase/groups/$TargetDeviceGroupId/members/microsoft.graph.device?`$select=id"
$currentDeviceIds = New-Object System.Collections.Generic.HashSet[string]
foreach ($m in $currentMembers) { if ($m.id) { [void]$currentDeviceIds.Add($m.id) } }

# Compare sets
$toAdd    = $desiredDeviceIds.Where({ -not $currentDeviceIds.Contains($_) })
$toRemove = $currentDeviceIds.Where({ -not $desiredDeviceIds.Contains($_) })

Write-Host "4) Adding missing devices: $($toAdd.Count)"
foreach ($did in $toAdd) {
    $body = @{
        "@odata.id" = "$GraphBase/devices/$did"
    }
    Invoke-Graph -Method POST -Uri "$GraphBase/groups/$TargetDeviceGroupId/members/`$ref" -Body $body | Out-Null
}

Write-Host "5) Removing extra devices: $($toRemove.Count)"
foreach ($did in $toRemove) {
    # Need the directoryObject reference in group membership; Graph supports delete by member id:
    Invoke-Graph -Method DELETE -Uri "$GraphBase/groups/$TargetDeviceGroupId/members/$did/`$ref" | Out-Null
}

Write-Host "DONE. Added=$($toAdd.Count) Removed=$($toRemove.Count)"
