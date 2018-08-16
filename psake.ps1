properties {
    $projectRoot = $ENV:BHProjectPath
    if(-not $projectRoot) {
        $projectRoot = $PSScriptRoot
    }

    $sut = $env:BHModulePath
    $tests = "$projectRoot\Tests"
    $outputDir = Join-Path -Path $projectRoot -ChildPath 'out'
    $outputModDir = Join-Path -Path $outputDir -ChildPath $env:BHProjectName
    $manifest = Import-PowerShellDataFile -Path $env:BHPSModuleManifest
    $outputModVerDir = Join-Path -Path $outputModDir -ChildPath $manifest.ModuleVersion
    $psVersion = $PSVersionTable.PSVersion.Major
}

task default -depends Publish

task Init {
    "`nSTATUS: Testing with PowerShell $psVersion"
    "Build System Details:"
    Get-Item ENV:BH*
    "`n"

    'Configuration', 'Pester', 'PSScriptAnalyzer', 'PSSlack', 'PoshBot', 'PSGSuite' | Foreach-Object {
        if (-not (Get-Module -Name $_ -ListAvailable -ErrorAction SilentlyContinue)) {
            Install-Module -Name $_ -Repository PSGallery -Scope CurrentUser -AllowClobber -Confirm:$false -ErrorAction Stop
        }
        Import-Module -Name $_ -Verbose:$false -Force -ErrorAction Stop
    }
} -description 'Initialize build environment'

task Test -Depends Init, Analyze, Pester -description 'Run test suite'

task Analyze -Depends Build {
    $analysis = Invoke-ScriptAnalyzer -Path $outputModVerDir -Verbose:$false
    $errors = $analysis | Where-Object {$_.Severity -eq 'Error'}
    $warnings = $analysis | Where-Object {$_.Severity -eq 'Warning'}

    if (($errors.Count -eq 0) -and ($warnings.Count -eq 0)) {
        '    PSScriptAnalyzer passed without errors or warnings'
    }

    if (@($errors).Count -gt 0) {
        Write-Error -Message 'One or more Script Analyzer errors were found. Build cannot continue!'
        $errors | Format-Table
    }

    if (@($warnings).Count -gt 0) {
        Write-Warning -Message 'One or more Script Analyzer warnings were found. These should be corrected.'
        $warnings | Format-Table
    }
} -description 'Run PSScriptAnalyzer'

task Pester -Depends Build {
    Push-Location
    Set-Location -PassThru $outputModDir
    if(-not $ENV:BHProjectPath) {
        Set-BuildEnvironment -Path $PSScriptRoot\..
    }
    Remove-Module $ENV:BHProjectName -ErrorAction SilentlyContinue -Verbose:$false
    Import-Module -Name $outputModDir -Force -Verbose:$false
    $testResultsXml = Join-Path -Path $outputDir -ChildPath 'testResults.xml'
    $testResults = Invoke-Pester -Path $tests -PassThru -OutputFile $testResultsXml -OutputFormat NUnitXml
    if ($env:APPVEYOR) {
        (New-Object 'System.Net.WebClient').UploadFile(([Uri]"https://ci.appveyor.com/api/testresults/nunit/$($env:APPVEYOR_JOB_ID)"),$testResultsXml)
    }
    if ($testResults.FailedCount -gt 0) {
        $testResults | Format-List
        Write-Error -Message 'One or more Pester tests failed. Build cannot continue!'
    }
    Pop-Location
} -description 'Run Pester tests'

Task Publish -Depends Test {
    if (
        $env:BHProjectName -and $env:BHProjectName.Count -eq 1 -and
        $env:BHBuildSystem -ne 'Unknown' -and
        $env:BHBranchName -eq "master" -and
        $env:BHCommitMessage -match '!deploy' -and
        $env:APPVEYOR_BUILD_WORKER_IMAGE -like '*2017*' -and
        $env:APPVEYOR_PULL_REQUEST_NUMBER -eq $null 
    ) {
        "    Deploying version [$($manifest.ModuleVersion)] to PSGallery..."
        Publish-Module -Path $outputModVerDir -NuGetApiKey $env:NuGetApiKey -Repository PSGallery
    }
    else {
        "    Skipping deployment: To deploy, ensure that...`n" +
        "        * You are in a known build system (Current: $ENV:BHBuildSystem)`n" +
        "        * You are committing to the master branch (Current: $ENV:BHBranchName) `n" +
        "        * You are not building a Pull Request (Current: $ENV:APPVEYOR_PULL_REQUEST_NUMBER) `n" +
        "        * Your commit message includes !deploy (Current: $ENV:BHCommitMessage) `n" +
        "        * Your build image is Visual Studio 2017 (Current: $ENV:APPVEYOR_BUILD_WORKER_IMAGE)" |
            Write-Host
    }
}

task Clean -depends Init {
    Remove-Module -Name $env:BHProjectName -Force -ErrorAction SilentlyContinue

    if (Test-Path -Path $outputDir) {
        Get-ChildItem -Path $outputDir -Recurse | Remove-Item -Force -Recurse
    } else {
        New-Item -Path $outputDir -ItemType Directory > $null
    }
    "    Cleaned previous output directory [$outputDir]"
} -description 'Cleans module output directory'

task Compile -depends Clean {
    # Create module output directory
    $modDir = New-Item -Path $outputModDir -ItemType Directory
    New-Item -Path $outputModVerDir -ItemType Directory > $null

    # Append items to psm1
    Write-Verbose -Message 'Creating psm1...'
    $psm1 = New-Item -Path (Join-Path -Path $outputModVerDir -ChildPath "$($ENV:BHProjectName).psm1") -ItemType File

    'using module PoshBot' | Out-File -FilePath $psm1 -Encoding utf8 -Append

    Get-ChildItem -Path (Join-Path -Path $sut -ChildPath 'Classes') -File -Recurse |
        Get-Content -Raw | Add-Content -Path $psm1 -Encoding UTF8

    Get-ChildItem -Path (Join-Path -Path $sut -ChildPath 'Public') -Recurse |
        Get-Content -Raw | Add-Content -Path $psm1 -Encoding UTF8

    # Copy over other items
    Copy-Item -Path $env:BHPSModuleManifest -Destination $outputModVerDir

    # Set functions to export
    $functions = (Get-ChildItem -Path (Join-Path -Path $sut -ChildPath 'Public')).BaseName
    Update-ModuleManifest -Path (Join-Path -Path $outputModVerDir -ChildPath (Split-Path -Path $env:BHPSModuleManifest -Leaf)) -FunctionsToExport $functions

    "    Created compiled module at [$modDir]"
} -description 'Compiles module from source'

task Build -depends Compile {
}