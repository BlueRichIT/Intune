<#
.SYNOPSIS
Author- BRIT IT Solutions
Script Description — Sync Devices to a Static Entra ID Device Group from Two User Groups
Name: Sync-DeviceGroupFromUserGroups.ps1
Author: BRIT IT Solutions
Purpose:
This script synchronizes membership of a static Azure Entra ID device group so that it contains all devices registered to users who belong to either of two specified (static) user groups. It’s a practical workaround for scenarios where a “dynamic device group” cannot express “devices owned by users in group A or group B.” The script queries Microsoft Graph to compute the desired device set and then adds/removes devices in the target device group to match.
⚠️ Note: Dynamic device membership rules in Entra ID cannot natively select devices based on user group membership. This script maintains a static device group to achieve that outcome.
What the Script Does (High Level)
Authentication
Obtains an application (client credentials) token for Microsoft Graph using the provided Tenant ID, Client ID, and Client Secret.
Collect Source Users
Retrieves transitive members (users) of the two provided user groups and builds a unique set of user IDs (union of both groups).
Resolve Devices per User
For each user in the union set, pulls their registered devices and builds the desired device ID set.
Read Current Target Group Membership
Gets the current device members of the target device group.
Delta Calculation & Apply Changes
Add devices that are missing from the target group.
Remove devices that are no longer desired in the target group.
Prints a summary with Added/Removed counts.

Parameters
-TenantId (String, Required)
Your Azure AD/Entra tenant ID (GUID).
-ClientId (String, Required)
The App Registration (Enterprise App) Client ID that has Graph application permissions.
-ClientSecret (String, Required)
The client secret issued to the App Registration.
-UserGroupId1 (String, Required)
Object ID of the first user group (static) to consider.
-UserGroupId2 (String, Required)
Object ID of the second user group (static) to consider.
-TargetDeviceGroupId (String, Required)
Object ID of the device group (static) the script will keep in sync.
Required Microsoft Graph Permissions (Application)
Grant these application permissions to the App Registration and admin consent them:
Group.Read.All
GroupMember.Read.All
Directory.Read.All
GroupMember.ReadWrite.All (required to add/remove target group members)
Authentication uses client credentials (client secret). Certificate-based auth can be substituted with minimal changes.
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

