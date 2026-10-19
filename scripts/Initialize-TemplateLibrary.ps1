#Requires -Version 5.1
<#
.SYNOPSIS
    Creates the SharePoint document library and columns the Outlook add-in reads templates from,
    and optionally uploads sample templates.

.DESCRIPTION
    Uses Microsoft Graph to:
      1. find the site;
      2. create the document library (-LibraryName) if it doesn't exist;
      3. add the columns the add-in reads, if they don't exist:
           TemplateCategory    "Template category"    choice (optional; empty shows as "Uncategorized")
           TemplateDescription "Template description" plain text, multiple lines
           TemplateKeywords    "Template keywords"    single line of text, comma-separated
         The built-in Title column is the name users see (the file name is used when Title is empty);
      4. with -SampleTemplatesPath, upload every *.html file in that folder (into -FolderPath, if given) and fill
         the columns from the file's <meta name="template-..."> tags (only this script reads those tags; the
         add-in reads the columns).

    Run it as an owner of the site (or a SharePoint admin). It is safe to run again: existing libraries and
    columns are left alone, and sample files with the same name are overwritten.

    This script does NOT grant the add-in access to the site; run Grant-TemplateSiteAccess.ps1 for that.

.PARAMETER SiteUrl
    The templates site, for example https://contoso.sharepoint.com/sites/EmailTemplates

.PARAMETER LibraryName
    Display name (or URL name) of the document library. Must match the add-in's SHAREPOINT_LIBRARY setting.
    Use "Documents" for a site's default library.

.PARAMETER FolderPath
    Optional folder inside the library for uploaded samples, for example "Email". Must match the add-in's
    SHAREPOINT_FOLDER setting. Missing folders are created by the upload.

.PARAMETER Categories
    Choices for the Template category column. Categories found in sample templates are added automatically.
    To change the choices later, edit the column in SharePoint (Library settings > Template category).

.PARAMETER SampleTemplatesPath
    Optional folder with sample *.html templates to upload (for example samples/templates).

.PARAMETER TenantId
    Optional tenant ID or domain to sign in to.

.EXAMPLE
    ./scripts/Initialize-TemplateLibrary.ps1 -SiteUrl https://contoso.sharepoint.com/sites/EmailTemplates

.EXAMPLE
    ./scripts/Initialize-TemplateLibrary.ps1 -SiteUrl https://contoso.sharepoint.com/sites/EmailTemplates -SampleTemplatesPath samples/templates

.EXAMPLE
    ./scripts/Initialize-TemplateLibrary.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Office -LibraryName Documents -FolderPath Email
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^https://')]
    [string]$SiteUrl,

    [ValidateNotNullOrEmpty()]
    [string]$LibraryName = 'Email Templates',

    [string]$FolderPath,

    [string[]]$Categories = @('General', 'Customer Service', 'HR', 'Sales'),

    [string]$SampleTemplatesPath,

    [string]$TenantId
)

$ErrorActionPreference = 'Stop'
. "$PSScriptRoot/TemplateLibrary.Common.ps1"

function Get-HtmlAttribute {
    param([string]$Tag, [string]$Name)

    $pattern = '(?is)\s' + [regex]::Escape($Name) + '\s*=\s*(?:"([^"]*)"|''([^'']*)''|([^\s>]+))'
    $match = [regex]::Match($Tag, $pattern)
    if (-not $match.Success) {
        return $null
    }
    foreach ($group in 1..3) {
        if ($match.Groups[$group].Success) {
            return [System.Net.WebUtility]::HtmlDecode($match.Groups[$group].Value)
        }
    }
}

# Reads <meta name="template-..." content="..."> tags from a sample template's <head>.
function Get-TemplateMeta {
    param([string]$Html)

    $meta = @{}
    $head = [regex]::Match($Html, '(?is)<head\b[^>]*>(.*?)</head>')
    if (-not $head.Success) {
        return $meta
    }
    foreach ($tag in [regex]::Matches($head.Groups[1].Value, '(?is)<meta\b[^>]*>')) {
        $name = Get-HtmlAttribute -Tag $tag.Value -Name 'name'
        if ($name -and $name -like 'template-*') {
            $meta[$name.ToLowerInvariant()] = ([string](Get-HtmlAttribute -Tag $tag.Value -Name 'content')).Trim()
        }
    }
    return $meta
}

$sampleFiles = @()
$samples = @()
if ($SampleTemplatesPath) {
    $sampleFiles = @(Get-ChildItem -LiteralPath $SampleTemplatesPath -Filter '*.html' -File | Sort-Object -Property Name)
    if ($sampleFiles.Count -eq 0) {
        throw "No *.html files found in '$SampleTemplatesPath'."
    }
    foreach ($file in $sampleFiles) {
        $html = [System.IO.File]::ReadAllText($file.FullName)
        $samples += [pscustomobject]@{ File = $file; Meta = (Get-TemplateMeta -Html $html) }
    }
    $Categories = @($Categories) + @($samples | ForEach-Object { $_.Meta['template-category'] } | Where-Object { $_ })
}
$Categories = @($Categories | Where-Object { $_ } | ForEach-Object { $_.Trim() } | Sort-Object -Unique)

Connect-TemplateGraph -Scopes 'Sites.Manage.All' -TenantId $TenantId

$site = Get-TemplateSite -SiteUrl $SiteUrl
$siteId = $site['id']
Write-Information "Site: $($site['displayName']) ($($site['webUrl']))" -InformationAction Continue

# 1. Library
$lists = @(Get-GraphCollection -Uri "$GraphBaseUrl/sites/$siteId/lists?`$select=id,displayName,name,webUrl")
$library = $lists | Where-Object { $_['displayName'] -eq $LibraryName -or $_['name'] -eq $LibraryName } | Select-Object -First 1
if ($library) {
    Write-Information "Library '$LibraryName' already exists." -InformationAction Continue
}
elseif ($PSCmdlet.ShouldProcess($SiteUrl, "Create document library '$LibraryName'")) {
    $body = @{ displayName = $LibraryName; list = @{ template = 'documentLibrary' } } | ConvertTo-Json -Depth 5
    $library = Invoke-MgGraphRequest -Method POST -Uri "$GraphBaseUrl/sites/$siteId/lists" -Body $body -ContentType 'application/json'
    Write-Information "Created library '$LibraryName'." -InformationAction Continue
}
else {
    return
}
$listId = $library['id']

# 2. Columns
$columnDefinitions = @(
    @{
        name        = 'TemplateCategory'
        displayName = 'Template category'
        description = 'Groups the template in the add-in''s category filter.'
        required    = $false
        choice      = @{ allowTextEntry = $false; choices = @($Categories); displayAs = 'dropDownMenu' }
    },
    @{
        name        = 'TemplateDescription'
        displayName = 'Template description'
        description = 'Shown under the template name in the add-in and searchable.'
        text        = @{ allowMultipleLines = $true; linesForEditing = 3; textType = 'plain' }
    },
    @{
        name        = 'TemplateKeywords'
        displayName = 'Template keywords'
        description = 'Extra search words, separated by commas.'
        text        = @{ allowMultipleLines = $false; maxLength = 255 }
    }
)
$existingColumns = @(Get-GraphCollection -Uri "$GraphBaseUrl/sites/$siteId/lists/$listId/columns?`$select=name" | ForEach-Object { $_['name'] })
foreach ($column in $columnDefinitions) {
    if ($existingColumns -contains $column.name) {
        Write-Information "Column '$($column.name)' already exists." -InformationAction Continue
        continue
    }
    if ($PSCmdlet.ShouldProcess($LibraryName, "Add column '$($column.name)'")) {
        Invoke-MgGraphRequest -Method POST -Uri "$GraphBaseUrl/sites/$siteId/lists/$listId/columns" -Body ($column | ConvertTo-Json -Depth 5) -ContentType 'application/json' | Out-Null
        Write-Information "Added column '$($column.name)'." -InformationAction Continue
    }
}

# 3. Sample templates
foreach ($sample in $samples) {
    $file = $sample.File
    $meta = $sample.Meta
    if (-not $PSCmdlet.ShouldProcess($LibraryName, "Upload sample '$($file.Name)'")) {
        continue
    }
    $segments = @(@($FolderPath -split '/') + $file.Name | Where-Object { $_ } | ForEach-Object { [uri]::EscapeDataString($_.Trim()) })
    $name = $segments -join '/'
    $driveItem = Invoke-MgGraphRequest -Method PUT -Uri "$GraphBaseUrl/sites/$siteId/lists/$listId/drive/root:/${name}:/content" -InputFilePath $file.FullName -ContentType 'text/html'

    $fields = @{ Title = $meta['template-name'] }
    if ($meta['template-category']) { $fields.TemplateCategory = $meta['template-category'] }
    if ($meta['template-description']) { $fields.TemplateDescription = $meta['template-description'] }
    if ($meta['template-keywords']) { $fields.TemplateKeywords = $meta['template-keywords'] }
    $driveId = $driveItem['parentReference']['driveId']
    Invoke-MgGraphRequest -Method PATCH -Uri "$GraphBaseUrl/drives/$driveId/items/$($driveItem['id'])/listItem/fields" -Body ($fields | ConvertTo-Json) -ContentType 'application/json' | Out-Null
    Write-Information "Uploaded '$($file.Name)' ($($fields.Title))." -InformationAction Continue
}

Write-Information "Done. Library: $($library['webUrl'])" -InformationAction Continue
