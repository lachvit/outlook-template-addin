#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
# PSScriptAnalyzer can't see that BeforeAll/BeforeEach variables are used in It blocks.
[System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]
param()

BeforeAll {
    $ScriptPath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../scripts/Publish-TemplateIndex.ps1')).ProviderPath
    $BaseUrl = 'https://contoso.example/repo/templates/'
    $ValidHead = '<meta charset="utf-8">' +
        '<meta name="template-name" content="Welcome &amp; hello">' +
        '<meta name="template-category" content="HR">' +
        '<meta name="template-description" content="First day info">' +
        '<meta name="template-keywords" content="onboarding, new hire">'

    function New-TemplateFolder {
        $folder = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path (Join-Path $folder 'images') -Force
        Set-Content -LiteralPath (Join-Path $folder 'images/logo.png') -Value 'not really a png'
        return $folder
    }

    function Add-Template {
        param([string]$Folder, [string]$Name, [string]$Head, [string]$Body)
        $html = "<!DOCTYPE html><html><head>$Head</head><body>$Body</body></html>"
        Set-Content -LiteralPath (Join-Path $Folder $Name) -Value $html -Encoding UTF8
    }

    # Runs the script and returns its exit code; warnings and info messages are discarded.
    function Invoke-PublishScript {
        param([string]$Folder, [string]$OutputPath, [switch]$CheckOnly)
        if ($CheckOnly) {
            & $ScriptPath -TemplateRoot $Folder -CheckOnly 3>$null 6>$null
        }
        else {
            & $ScriptPath -TemplateRoot $Folder -OutputPath $OutputPath -BaseUrl $BaseUrl 3>$null 6>$null
        }
        return $LASTEXITCODE
    }
}

Describe 'Publish-TemplateIndex.ps1' {
    BeforeEach {
        $folder = New-TemplateFolder
        $output = Join-Path $TestDrive "out-$([guid]::NewGuid().ToString('N'))/index.json"
    }

    Context 'with valid templates' {
        It 'writes index.json with all fields and exits 0' {
            Add-Template -Folder $folder -Name 'hr-welcome.html' -Head $ValidHead -Body '<p>Hi</p><img src="images/logo.png" alt="Logo">'

            Invoke-PublishScript -Folder $folder -OutputPath $output | Should -Be 0
            Test-Path -LiteralPath $output | Should -BeTrue

            $json = Get-Content -LiteralPath $output -Raw
            $index = $json | ConvertFrom-Json
            @($index.templates).Count | Should -Be 1
            $entry = @($index.templates)[0]
            $entry.id | Should -Be 'hr-welcome'
            $entry.name | Should -Be 'Welcome & hello'
            $entry.category | Should -Be 'HR'
            $entry.description | Should -Be 'First day info'
            @($entry.keywords) -join '|' | Should -Be 'onboarding|new hire'
            $entry.url | Should -Be 'https://contoso.example/repo/templates/hr-welcome.html'
            # Checked on the raw JSON because PowerShell 7's ConvertFrom-Json turns ISO dates into DateTime.
            $json | Should -Match '"lastModified":\s*"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z"'
        }

        It 'accepts templates without optional metadata and https images' {
            $head = '<meta name="template-name" content="Short"><meta name="template-category" content="General">'
            Add-Template -Folder $folder -Name 'short.html' -Head $head -Body '<img src="https://cdn.example/a.png" alt="">'

            Invoke-PublishScript -Folder $folder -OutputPath $output | Should -Be 0
            $entry = @((Get-Content -LiteralPath $output -Raw | ConvertFrom-Json).templates)[0]
            $entry.description | Should -Be ''
            @($entry.keywords).Count | Should -Be 0
        }

        It 'writes nothing in -CheckOnly mode' {
            Add-Template -Folder $folder -Name 'hr-welcome.html' -Head $ValidHead -Body '<p>Hi</p>'

            Invoke-PublishScript -Folder $folder -CheckOnly | Should -Be 0
            Test-Path -LiteralPath (Join-Path $folder 'index.json') | Should -BeFalse
        }
    }

    Context 'with invalid templates' {
        It 'fails when template-category is missing' {
            Add-Template -Folder $folder -Name 'no-category.html' -Head '<meta name="template-name" content="X">' -Body '<p>Hi</p>'

            Invoke-PublishScript -Folder $folder -OutputPath $output | Should -Be 1
            Test-Path -LiteralPath $output | Should -BeFalse
        }

        It 'fails when template-name is empty' {
            $head = '<meta name="template-name" content=" "><meta name="template-category" content="HR">'
            Add-Template -Folder $folder -Name 'empty-name.html' -Head $head -Body '<p>Hi</p>'

            Invoke-PublishScript -Folder $folder -OutputPath $output | Should -Be 1
        }

        It 'fails when the template contains a script tag' {
            Add-Template -Folder $folder -Name 'script.html' -Head $ValidHead -Body '<p>Hi</p><script>alert(1)</script>'

            Invoke-PublishScript -Folder $folder -OutputPath $output | Should -Be 1
        }

        It 'fails when the template contains an event handler attribute' {
            Add-Template -Folder $folder -Name 'handler.html' -Head $ValidHead -Body '<p onclick="alert(1)">Hi</p>'

            Invoke-PublishScript -Folder $folder -OutputPath $output | Should -Be 1
        }

        It 'fails when the template contains a javascript URL' {
            Add-Template -Folder $folder -Name 'js-link.html' -Head $ValidHead -Body '<a href="javascript:alert(1)">x</a>'

            Invoke-PublishScript -Folder $folder -OutputPath $output | Should -Be 1
        }

        It 'fails when a referenced image does not exist' {
            Add-Template -Folder $folder -Name 'missing-image.html' -Head $ValidHead -Body '<img src="images/missing.png" alt="">'

            Invoke-PublishScript -Folder $folder -OutputPath $output | Should -Be 1
            Test-Path -LiteralPath $output | Should -BeFalse
        }

        It 'fails when an image points outside the templates folder' {
            Add-Template -Folder $folder -Name 'outside.html' -Head $ValidHead -Body '<img src="../secret.png" alt="">'

            Invoke-PublishScript -Folder $folder -OutputPath $output | Should -Be 1
        }

        It 'fails when an image uses plain http' {
            Add-Template -Folder $folder -Name 'http-image.html' -Head $ValidHead -Body '<img src="http://cdn.example/a.png" alt="">'

            Invoke-PublishScript -Folder $folder -OutputPath $output | Should -Be 1
        }

        It 'fails when the file name is not lowercase-with-hyphens' {
            Add-Template -Folder $folder -Name 'My Template.html' -Head $ValidHead -Body '<p>Hi</p>'

            Invoke-PublishScript -Folder $folder -OutputPath $output | Should -Be 1
        }

        It 'fails the whole run and writes nothing when one of several templates is invalid' {
            Add-Template -Folder $folder -Name 'good.html' -Head $ValidHead -Body '<p>Hi</p>'
            Add-Template -Folder $folder -Name 'bad.html' -Head $ValidHead -Body '<script></script>'

            Invoke-PublishScript -Folder $folder -OutputPath $output | Should -Be 1
            Test-Path -LiteralPath $output | Should -BeFalse
        }

        It 'fails when the folder has no templates' {
            Invoke-PublishScript -Folder $folder -OutputPath $output | Should -Be 1
        }
    }
}
