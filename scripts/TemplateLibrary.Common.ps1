# Shared helpers for the SharePoint setup scripts. Dot-source it: . "$PSScriptRoot/TemplateLibrary.Common.ps1"
# Needs the Microsoft.Graph.Authentication module (Install-Module Microsoft.Graph.Authentication -Scope CurrentUser).

Set-StrictMode -Version Latest

$GraphBaseUrl = 'https://graph.microsoft.com/v1.0'

function Assert-GraphModule {
    if (-not (Get-Command -Name Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
        throw 'The Microsoft.Graph.Authentication module is required. Install it with: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser'
    }
}

# Signs in (interactive) unless the current Graph session already has every scope needed.
function Connect-TemplateGraph {
    param(
        [Parameter(Mandatory)][string[]]$Scopes,
        [string]$TenantId
    )

    Assert-GraphModule
    $context = Get-MgContext
    $missing = @($Scopes)
    if ($context) {
        $missing = @($Scopes | Where-Object { @($context.Scopes) -notcontains $_ })
    }
    if ($missing.Count -eq 0 -and (-not $TenantId -or $context.TenantId -eq $TenantId)) {
        return
    }
    $connectParams = @{ Scopes = $Scopes; NoWelcome = $true }
    if ($TenantId) {
        $connectParams.TenantId = $TenantId
    }
    Connect-MgGraph @connectParams
}

# https://contoso.sharepoint.com/sites/Email Templates -> contoso.sharepoint.com:/sites/Email%20Templates (URL-ready)
function Get-GraphSiteKey {
    param([Parameter(Mandatory)][string]$SiteUrl)

    $uri = $null
    if (-not [uri]::TryCreate($SiteUrl, [System.UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -ne 'https') {
        throw "SiteUrl must be an https:// SharePoint site URL (got '$SiteUrl')."
    }
    $segments = @($uri.AbsolutePath.Split('/') | Where-Object { $_ } | ForEach-Object { [uri]::EscapeDataString([uri]::UnescapeDataString($_)) })
    if ($segments.Count -gt 0) {
        return "$($uri.Host):/$($segments -join '/')"
    }
    return $uri.Host
}

function Get-TemplateSite {
    param([Parameter(Mandatory)][string]$SiteUrl)

    $key = Get-GraphSiteKey -SiteUrl $SiteUrl
    $site = Invoke-MgGraphRequest -Method GET -Uri "$GraphBaseUrl/sites/${key}?`$select=id,displayName,webUrl"
    if (-not $site -or -not $site['id']) {
        throw "Site '$SiteUrl' was not found."
    }
    return $site
}

# Follows @odata.nextLink and outputs every item of a Graph collection (wrap calls in @()).
function Get-GraphCollection {
    param([Parameter(Mandatory)][string]$Uri)

    $items = New-Object System.Collections.Generic.List[object]
    $next = $Uri
    while ($next) {
        $page = Invoke-MgGraphRequest -Method GET -Uri $next
        foreach ($entry in @($page['value'])) {
            if ($null -ne $entry) {
                $items.Add($entry)
            }
        }
        $next = $null
        if ($page.ContainsKey('@odata.nextLink')) {
            $next = $page['@odata.nextLink']
        }
    }
    return $items.ToArray()
}
