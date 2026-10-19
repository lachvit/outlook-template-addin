#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
# PSScriptAnalyzer can't see that BeforeAll/BeforeEach variables are used in It blocks.
[System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]
param()

BeforeAll {
    $ScriptsRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../scripts')).ProviderPath
    $Graph = 'https://graph.microsoft.com/v1.0'
    $SiteUrl = 'https://contoso.sharepoint.com/sites/Email Templates'
    $SiteId = 'contoso.sharepoint.com,1111,2222'
    $ClientId = '6f1d2c84-3b7a-4e59-9a0d-8c2e5b41f7a3'
    # Mutable state the mocks read; each test sets what Graph should return.
    $State = @{}

    # Stand-ins so the tests also run where Microsoft.Graph.Authentication isn't installed; every call is mocked.
    foreach ($name in 'Invoke-MgGraphRequest', 'Connect-MgGraph', 'Get-MgContext') {
        if (-not (Get-Command -Name $name -ErrorAction SilentlyContinue)) {
            Set-Item -Path "function:global:$name" -Value { param([Parameter(ValueFromRemainingArguments)]$Rest) throw 'not mocked' }
        }
    }
}

Describe 'Get-GraphSiteKey' {
    BeforeAll {
        . (Join-Path $ScriptsRoot 'TemplateLibrary.Common.ps1')
    }

    It 'builds a URL-ready path-based site key' {
        Get-GraphSiteKey -SiteUrl 'https://contoso.sharepoint.com/sites/Email Templates/' | Should -Be 'contoso.sharepoint.com:/sites/Email%20Templates'
        Get-GraphSiteKey -SiteUrl 'https://contoso.sharepoint.com/sites/Email%20Templates' | Should -Be 'contoso.sharepoint.com:/sites/Email%20Templates'
        Get-GraphSiteKey -SiteUrl 'https://contoso.sharepoint.com' | Should -Be 'contoso.sharepoint.com'
    }

    It 'rejects non-https URLs' {
        { Get-GraphSiteKey -SiteUrl 'http://contoso.sharepoint.com/sites/x' } | Should -Throw '*https://*'
    }
}

Describe 'Grant-TemplateSiteAccess.ps1' {
    BeforeAll {
        $Script = Join-Path $ScriptsRoot 'Grant-TemplateSiteAccess.ps1'
    }

    BeforeEach {
        $State.Permissions = @()
        Mock Get-MgContext { $null }
        Mock Connect-MgGraph {}
        Mock Invoke-MgGraphRequest { @{ id = $SiteId; displayName = 'Email Templates'; webUrl = $SiteUrl } } -ParameterFilter { $Uri -like "$Graph/sites/contoso.sharepoint.com:*" }
        Mock Invoke-MgGraphRequest { @{ value = $State.Permissions } } -ParameterFilter { $Method -eq 'GET' -and $Uri -eq "$Graph/sites/$SiteId/permissions" }
        Mock Invoke-MgGraphRequest { @{} } -ParameterFilter { $Method -ne 'GET' }
    }

    It 'signs in with Sites.FullControl.All and grants read access to the app' {
        & $Script -SiteUrl $SiteUrl -ClientId $ClientId 6>$null

        Should -Invoke Connect-MgGraph -Times 1 -ParameterFilter { $Scopes -contains 'Sites.FullControl.All' }
        Should -Invoke Invoke-MgGraphRequest -Times 1 -ParameterFilter {
            if (-not $Body) { return $false }
            $body = $Body | ConvertFrom-Json
            $Method -eq 'POST' -and $Uri -eq "$Graph/sites/$SiteId/permissions" -and
            $body.roles -contains 'read' -and $body.grantedToIdentities[0].application.id -eq $ClientId
        }
    }

    It 'does nothing when the app already has the role' {
        $State.Permissions = @(@{ id = 'p1'; roles = @('read'); grantedToIdentitiesV2 = @(@{ application = @{ id = $ClientId } }) })

        & $Script -SiteUrl $SiteUrl -ClientId $ClientId 6>$null

        Should -Invoke Invoke-MgGraphRequest -Times 0 -ParameterFilter { $Method -ne 'GET' }
    }

    It 'changes the role of an existing grant' {
        $State.Permissions = @(@{ id = 'p1'; roles = @('write'); grantedToIdentities = @(@{ application = @{ id = $ClientId } }) })

        & $Script -SiteUrl $SiteUrl -ClientId $ClientId 6>$null

        Should -Invoke Invoke-MgGraphRequest -Times 1 -ParameterFilter { $Method -eq 'PATCH' -and $Uri -eq "$Graph/sites/$SiteId/permissions/p1" }
        Should -Invoke Invoke-MgGraphRequest -Times 0 -ParameterFilter { $Method -eq 'POST' }
    }

    It 'revokes only the grants of this app' {
        $State.Permissions = @(
            @{ id = 'other'; roles = @('read'); grantedToIdentitiesV2 = @(@{ application = @{ id = 'aaaaaaaa-0000-0000-0000-000000000000' } }) },
            @{ id = 'mine'; roles = @('read'); grantedToIdentitiesV2 = @(@{ application = @{ id = $ClientId } }) }
        )

        & $Script -SiteUrl $SiteUrl -ClientId $ClientId -Revoke 6>$null

        Should -Invoke Invoke-MgGraphRequest -Times 1 -ParameterFilter { $Method -eq 'DELETE' -and $Uri -eq "$Graph/sites/$SiteId/permissions/mine" }
        Should -Invoke Invoke-MgGraphRequest -Times 0 -ParameterFilter { $Method -eq 'DELETE' -and $Uri -like '*/other' }
    }
}

Describe 'Initialize-TemplateLibrary.ps1' {
    BeforeAll {
        $Script = Join-Path $ScriptsRoot 'Initialize-TemplateLibrary.ps1'
    }

    BeforeEach {
        $State.Lists = @()
        $State.Columns = @()
        Mock Get-MgContext { @{ Scopes = @('Sites.Manage.All'); TenantId = 't' } }
        Mock Connect-MgGraph {}
        Mock Invoke-MgGraphRequest { @{ id = $SiteId; displayName = 'Email Templates'; webUrl = $SiteUrl } } -ParameterFilter { $Uri -like "$Graph/sites/contoso.sharepoint.com:*" }
        Mock Invoke-MgGraphRequest { @{ value = $State.Lists } } -ParameterFilter { $Method -eq 'GET' -and $Uri -like "$Graph/sites/$SiteId/lists?*" }
        Mock Invoke-MgGraphRequest { @{ value = $State.Columns } } -ParameterFilter { $Method -eq 'GET' -and $Uri -like '*/columns?*' }
        Mock Invoke-MgGraphRequest { @{ id = 'lib1'; webUrl = "$SiteUrl/Email Templates" } } -ParameterFilter { $Method -eq 'POST' -and $Uri -eq "$Graph/sites/$SiteId/lists" }
        Mock Invoke-MgGraphRequest { @{} } -ParameterFilter { $Method -eq 'POST' -and $Uri -like '*/columns' }
        Mock Invoke-MgGraphRequest { @{ id = 'file1'; parentReference = @{ driveId = 'drive1' } } } -ParameterFilter { $Method -eq 'PUT' }
        Mock Invoke-MgGraphRequest { @{} } -ParameterFilter { $Method -eq 'PATCH' }
    }

    It 'creates the library and the three columns when they are missing' {
        & $Script -SiteUrl $SiteUrl -Categories 'Sales', 'HR' 6>$null

        Should -Invoke Connect-MgGraph -Times 0
        Should -Invoke Invoke-MgGraphRequest -Times 1 -ParameterFilter {
            $Method -eq 'POST' -and $Uri -eq "$Graph/sites/$SiteId/lists" -and ($Body | ConvertFrom-Json).list.template -eq 'documentLibrary'
        }
        foreach ($name in 'TemplateCategory', 'TemplateDescription', 'TemplateKeywords') {
            Should -Invoke Invoke-MgGraphRequest -Times 1 -ParameterFilter {
                $Method -eq 'POST' -and $Uri -eq "$Graph/sites/$SiteId/lists/lib1/columns" -and ($Body | ConvertFrom-Json).name -eq $name
            }
        }
        Should -Invoke Invoke-MgGraphRequest -Times 1 -ParameterFilter {
            if (-not $Body) { return $false }
            $column = $Body | ConvertFrom-Json
            $column.name -eq 'TemplateCategory' -and -not $column.required -and (@($column.choice.choices) -join ',') -eq 'HR,Sales'
        }
    }

    It 'leaves an existing library and columns alone' {
        $State.Lists = @(@{ id = 'lib1'; displayName = 'Email Templates'; webUrl = "$SiteUrl/Email Templates" })
        $State.Columns = @(@{ name = 'TemplateCategory' }, @{ name = 'TemplateDescription' }, @{ name = 'TemplateKeywords' })

        & $Script -SiteUrl $SiteUrl 6>$null

        Should -Invoke Invoke-MgGraphRequest -Times 0 -ParameterFilter { $Method -ne 'GET' }
    }

    It 'uploads samples and fills the columns from their meta tags' {
        $State.Lists = @(@{ id = 'lib1'; displayName = 'Email Templates'; webUrl = "$SiteUrl/Email Templates" })
        $State.Columns = @(@{ name = 'TemplateCategory' }, @{ name = 'TemplateDescription' }, @{ name = 'TemplateKeywords' })
        $folder = Join-Path $TestDrive 'samples'
        $null = New-Item -ItemType Directory -Path $folder
        $html = '<html><head><meta name="template-name" content="Welcome &amp; hello"><meta name="template-category" content="HR">' +
            '<meta name="template-keywords" content="a, b"></head><body><p>Hi</p></body></html>'
        Set-Content -LiteralPath (Join-Path $folder 'hr welcome.html') -Value $html

        & $Script -SiteUrl $SiteUrl -SampleTemplatesPath $folder 6>$null

        Should -Invoke Invoke-MgGraphRequest -Times 1 -ParameterFilter {
            $Method -eq 'PUT' -and $Uri -eq "$Graph/sites/$SiteId/lists/lib1/drive/root:/hr%20welcome.html:/content" -and $InputFilePath -like '*hr welcome.html'
        }
        Should -Invoke Invoke-MgGraphRequest -Times 1 -ParameterFilter {
            if (-not $Body) { return $false }
            $fields = $Body | ConvertFrom-Json
            $Method -eq 'PATCH' -and $Uri -eq "$Graph/drives/drive1/items/file1/listItem/fields" -and
            $fields.Title -eq 'Welcome & hello' -and $fields.TemplateCategory -eq 'HR' -and $fields.TemplateKeywords -eq 'a, b' -and
            -not $fields.PSObject.Properties['TemplateDescription']
        }
    }

    It 'finds the default library by URL name and uploads into a folder' {
        $State.Lists = @(@{ id = 'docs'; displayName = 'Documents'; name = 'Shared Documents'; webUrl = "$SiteUrl/Shared Documents" })
        $State.Columns = @(@{ name = 'TemplateCategory' }, @{ name = 'TemplateDescription' }, @{ name = 'TemplateKeywords' })
        $folder = Join-Path $TestDrive 'samples2'
        $null = New-Item -ItemType Directory -Path $folder
        Set-Content -LiteralPath (Join-Path $folder 'a.html') -Value '<html><head><meta name="template-name" content="A"></head><body></body></html>'

        & $Script -SiteUrl $SiteUrl -LibraryName 'Shared Documents' -FolderPath '/Email/Team A/' -SampleTemplatesPath $folder 6>$null

        Should -Invoke Invoke-MgGraphRequest -Times 0 -ParameterFilter { $Method -eq 'POST' -and $Uri -eq "$Graph/sites/$SiteId/lists" }
        Should -Invoke Invoke-MgGraphRequest -Times 1 -ParameterFilter {
            $Method -eq 'PUT' -and $Uri -eq "$Graph/sites/$SiteId/lists/docs/drive/root:/Email/Team%20A/a.html:/content"
        }
    }
}
