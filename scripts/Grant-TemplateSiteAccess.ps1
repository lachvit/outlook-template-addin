#Requires -Version 5.1
<#
.SYNOPSIS
    Grants (or revokes) the add-in's app registration access to the SharePoint site that holds the templates.

.DESCRIPTION
    The add-in uses the delegated Microsoft Graph permission Sites.Selected. That permission gives the app
    access to NO site until an administrator grants it access to specific sites. This script does that grant
    for one site with Microsoft Graph (POST /sites/{site-id}/permissions).

    Users still only see what they themselves can read on the site: delegated Sites.Selected access is the
    intersection of the app's grant and the user's own SharePoint permissions.

    Run it as a Global Administrator, Privileged Role Administrator or SharePoint Administrator (the
    signed-in account needs Sites.FullControl.All, which requires admin consent). It is safe to run again.

.PARAMETER SiteUrl
    The templates site, for example https://contoso.sharepoint.com/sites/EmailTemplates

.PARAMETER ClientId
    Application (client) ID of the add-in's app registration.

.PARAMETER DisplayName
    Name stored with the grant, shown when listing the site's app permissions.

.PARAMETER Role
    read (default; all the add-in needs) or write.

.PARAMETER TenantId
    Optional tenant ID or domain to sign in to.

.PARAMETER Revoke
    Removes the app's grant from the site instead of adding it.

.EXAMPLE
    ./scripts/Grant-TemplateSiteAccess.ps1 -SiteUrl https://contoso.sharepoint.com/sites/EmailTemplates -ClientId 00000000-0000-0000-0000-000000000000

.EXAMPLE
    ./scripts/Grant-TemplateSiteAccess.ps1 -SiteUrl https://contoso.sharepoint.com/sites/EmailTemplates -ClientId 00000000-0000-0000-0000-000000000000 -Revoke
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^https://')]
    [string]$SiteUrl,

    [Parameter(Mandatory)]
    [ValidatePattern('^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')]
    [string]$ClientId,

    [string]$DisplayName = 'Outlook Email Templates add-in',

    [ValidateSet('read', 'write')]
    [string]$Role = 'read',

    [string]$TenantId,

    [switch]$Revoke
)

$ErrorActionPreference = 'Stop'
. "$PSScriptRoot/TemplateLibrary.Common.ps1"

# The application ids a site permission was granted to (v1.0 returns both property shapes).
function Get-PermissionAppId {
    param($Permission)

    $ids = @()
    foreach ($property in 'grantedToIdentitiesV2', 'grantedToIdentities') {
        if ($Permission.ContainsKey($property)) {
            foreach ($identity in @($Permission[$property])) {
                if ($identity -and $identity.ContainsKey('application') -and $identity['application']) {
                    $ids += [string]$identity['application']['id']
                }
            }
        }
    }
    return $ids
}

Connect-TemplateGraph -Scopes 'Sites.FullControl.All' -TenantId $TenantId

$site = Get-TemplateSite -SiteUrl $SiteUrl
$siteId = $site['id']
$permissionsUri = "$GraphBaseUrl/sites/$siteId/permissions"
$existing = @(Get-GraphCollection -Uri $permissionsUri | Where-Object { (Get-PermissionAppId -Permission $_) -contains $ClientId })

if ($Revoke) {
    if ($existing.Count -eq 0) {
        Write-Information "The app $ClientId has no access to $SiteUrl. Nothing to revoke." -InformationAction Continue
        return
    }
    foreach ($permission in $existing) {
        if ($PSCmdlet.ShouldProcess($SiteUrl, "Revoke access for app $ClientId")) {
            Invoke-MgGraphRequest -Method DELETE -Uri "$permissionsUri/$($permission['id'])" | Out-Null
        }
    }
    Write-Information "Revoked access for app $ClientId on $SiteUrl." -InformationAction Continue
    return
}

$current = $existing | Where-Object { @($_['roles']) -contains $Role } | Select-Object -First 1
if ($current) {
    Write-Information "The app $ClientId already has '$Role' access to $SiteUrl. Nothing to do." -InformationAction Continue
    return
}

if ($existing.Count -gt 0) {
    $permission = $existing[0]
    if ($PSCmdlet.ShouldProcess($SiteUrl, "Change access for app $ClientId to '$Role'")) {
        Invoke-MgGraphRequest -Method PATCH -Uri "$permissionsUri/$($permission['id'])" -Body (@{ roles = @($Role) } | ConvertTo-Json) -ContentType 'application/json' | Out-Null
        Write-Information "Changed the access of app $ClientId on $SiteUrl to '$Role'." -InformationAction Continue
    }
    return
}

$body = @{
    roles               = @($Role)
    grantedToIdentities = @(
        @{ application = @{ id = $ClientId; displayName = $DisplayName } }
    )
} | ConvertTo-Json -Depth 5

if ($PSCmdlet.ShouldProcess($SiteUrl, "Grant '$Role' access to app $ClientId")) {
    Invoke-MgGraphRequest -Method POST -Uri $permissionsUri -Body $body -ContentType 'application/json' | Out-Null
    Write-Information "Granted '$Role' access on $SiteUrl to app $ClientId ($DisplayName)." -InformationAction Continue
}
