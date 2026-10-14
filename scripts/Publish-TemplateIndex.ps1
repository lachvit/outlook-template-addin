#Requires -Version 5.1
<#
.SYNOPSIS
    Validates the HTML email templates and writes index.json for the Outlook add-in.

.DESCRIPTION
    Reads every *.html file directly inside -TemplateRoot (subfolders are ignored) and checks that:
      - the file name uses only lowercase letters, digits and hyphens;
      - <meta name="template-name"> and <meta name="template-category"> exist in <head>;
      - there is a <body>;
      - there are no <script> tags, on* event attributes or javascript: URLs;
      - every relative <img src> points to an existing file inside -TemplateRoot
        (the check is case-sensitive on Linux, like GitHub Pages).

    If every template passes, writes index.json to -OutputPath. If any template fails, nothing is
    written and the script exits with code 1. With -CheckOnly, only the checks run.

.PARAMETER TemplateRoot
    Folder containing the template *.html files.

.PARAMETER OutputPath
    Where to write index.json.

.PARAMETER BaseUrl
    Public https:// URL of the templates folder, for example https://contoso.github.io/repo/templates/

.PARAMETER CheckOnly
    Validate the templates without writing index.json.

.EXAMPLE
    ./scripts/Publish-TemplateIndex.ps1 -TemplateRoot templates -CheckOnly

.EXAMPLE
    ./scripts/Publish-TemplateIndex.ps1 -TemplateRoot templates -OutputPath dist/templates/index.json -BaseUrl https://contoso.github.io/repo/templates/
#>
[CmdletBinding(DefaultParameterSetName = 'Publish')]
param(
    [Parameter(Mandatory)]
    [string]$TemplateRoot,

    [Parameter(Mandatory, ParameterSetName = 'Publish')]
    [string]$OutputPath,

    [Parameter(Mandatory, ParameterSetName = 'Publish')]
    [ValidatePattern('^https://')]
    [string]$BaseUrl,

    [Parameter(Mandatory, ParameterSetName = 'Check')]
    [switch]$CheckOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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
            $meta[$name.ToLowerInvariant()] = [string](Get-HtmlAttribute -Tag $tag.Value -Name 'content')
        }
    }
    return $meta
}

# Returns a plain object[] of problem descriptions (callers wrap in @()).
function Test-TemplateFile {
    param([System.IO.FileInfo]$File, [string]$Root, [string]$Html, [hashtable]$Meta)

    $problems = New-Object System.Collections.Generic.List[string]

    if ($File.BaseName -cnotmatch '^[a-z0-9]+(-[a-z0-9]+)*$') {
        $problems.Add('File name must use only lowercase letters, digits and single hyphens (for example "sales-quote.html").')
    }
    foreach ($required in 'template-name', 'template-category') {
        if (-not $Meta.ContainsKey($required) -or [string]::IsNullOrWhiteSpace($Meta[$required])) {
            $problems.Add("Missing required <meta name=""$required"" content=""..."">.")
        }
    }
    if ($Html -notmatch '(?i)<body\b') {
        $problems.Add('Missing <body> element.')
    }
    if ($Html -match '(?i)<script\b') {
        $problems.Add('Contains a <script> tag. Scripts are not allowed in templates.')
    }
    if ($Html -match '(?is)<[a-z][^>]*\son[a-z]+\s*=') {
        $problems.Add('Contains an event handler attribute (onclick, onload, ...). These are not allowed.')
    }
    if ($Html -match '(?is)<[^>]+=\s*["'']?\s*javascript:') {
        $problems.Add('Contains a javascript: URL. These are not allowed.')
    }

    $rootPrefix = $Root.TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar
    foreach ($img in [regex]::Matches($Html, '(?is)<img\b[^>]*>')) {
        $src = Get-HtmlAttribute -Tag $img.Value -Name 'src'
        if ([string]::IsNullOrWhiteSpace($src)) {
            $problems.Add('An <img> tag has no src.')
            continue
        }
        if ($src -match '^(?i)(https|data|cid):') {
            continue
        }
        if ($src -match '^[a-zA-Z][a-zA-Z0-9+.-]*:' -or $src.StartsWith('/')) {
            $problems.Add("Image '$src' must be a path relative to the template (for example images/logo.png) or an https:// URL.")
            continue
        }
        $relative = [uri]::UnescapeDataString(($src -split '[?#]')[0])
        $full = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($File.DirectoryName, $relative))
        if (-not $full.StartsWith($rootPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            $problems.Add("Image '$src' is outside the templates folder and would not be published.")
        }
        elseif (-not (Test-Path -LiteralPath $full -PathType Leaf)) {
            $problems.Add("Image '$src' was not found. File names are case-sensitive on GitHub Pages.")
        }
    }

    return $problems.ToArray()
}

function Get-LastModified {
    param([System.IO.FileInfo]$File)

    # git writes to stderr outside a repository; do not let that become a terminating error.
    $ErrorActionPreference = 'Continue'
    if (Get-Command -Name git -CommandType Application -ErrorAction SilentlyContinue) {
        $output = & git -C $File.DirectoryName log -1 --format=%cI -- $File.Name 2>$null
        $date = [string]($output | Select-Object -First 1)
        if ($LASTEXITCODE -eq 0 -and $date) {
            $parsed = [DateTimeOffset]::Parse($date, [System.Globalization.CultureInfo]::InvariantCulture)
            return $parsed.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ssZ', [System.Globalization.CultureInfo]::InvariantCulture)
        }
    }
    return $File.LastWriteTimeUtc.ToString('yyyy-MM-ddTHH:mm:ssZ', [System.Globalization.CultureInfo]::InvariantCulture)
}

function Write-TemplateProblem {
    param([System.IO.FileInfo]$File, [string]$Message)

    Write-Warning "$($File.Name): $Message"
    if ($env:GITHUB_ACTIONS -eq 'true') {
        $path = $File.FullName
        if ($env:GITHUB_WORKSPACE -and $path.StartsWith($env:GITHUB_WORKSPACE)) {
            $path = $path.Substring($env:GITHUB_WORKSPACE.Length).TrimStart('/', '\')
        }
        # Shows the problem as an annotation on the file in the GitHub Actions summary.
        Write-Information "::error file=${path}::$Message" -InformationAction Continue
    }
}

$rootItem = Get-Item -LiteralPath $TemplateRoot
if (-not $rootItem.PSIsContainer) {
    throw "TemplateRoot '$TemplateRoot' is not a folder."
}
$root = $rootItem.FullName

$files = @(Get-ChildItem -LiteralPath $root -Filter '*.html' -File | Sort-Object -Property Name)
if ($files.Count -eq 0) {
    Write-Warning "No *.html templates found in '$root'."
    exit 1
}

$entries = New-Object System.Collections.Generic.List[object]
$failedCount = 0
foreach ($file in $files) {
    $html = [System.IO.File]::ReadAllText($file.FullName)
    $meta = Get-TemplateMeta -Html $html
    $problems = @(Test-TemplateFile -File $file -Root $root -Html $html -Meta $meta)
    if ($problems.Count -gt 0) {
        $failedCount++
        foreach ($problem in $problems) {
            Write-TemplateProblem -File $file -Message $problem
        }
        continue
    }
    if ($CheckOnly) {
        continue
    }

    $description = ''
    if ($meta.ContainsKey('template-description')) {
        $description = $meta['template-description'].Trim()
    }
    $keywords = @()
    if ($meta.ContainsKey('template-keywords')) {
        $keywords = @($meta['template-keywords'] -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }

    $entries.Add([pscustomobject]@{
            id           = $file.BaseName
            name         = $meta['template-name'].Trim()
            category     = $meta['template-category'].Trim()
            description  = $description
            keywords     = $keywords
            url          = $BaseUrl.TrimEnd('/') + '/' + [uri]::EscapeDataString($file.Name)
            lastModified = Get-LastModified -File $file
        })
}

if ($failedCount -gt 0) {
    Write-Warning "$failedCount of $($files.Count) template(s) failed validation. Fix the problems above; nothing was published."
    exit 1
}

if ($CheckOnly) {
    Write-Information "All $($files.Count) template(s) passed validation." -InformationAction Continue
    exit 0
}

$index = [ordered]@{
    generatedAt = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ', [System.Globalization.CultureInfo]::InvariantCulture)
    templates   = $entries.ToArray()
}
$json = ConvertTo-Json -InputObject $index -Depth 5

$fullOutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
$outputFolder = Split-Path -Parent $fullOutputPath
if ($outputFolder -and -not (Test-Path -LiteralPath $outputFolder)) {
    $null = New-Item -ItemType Directory -Path $outputFolder -Force
}
[System.IO.File]::WriteAllText($fullOutputPath, $json, (New-Object System.Text.UTF8Encoding $false))

Write-Information "Wrote $($entries.Count) template(s) to $fullOutputPath" -InformationAction Continue
exit 0
