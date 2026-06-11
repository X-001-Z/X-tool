param(
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$ReleaseRoot = Join-Path $PSScriptRoot "release"

if (-not $SkipBuild) {
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\build.ps1"
}

$Version = & ".\.venv-build\Scripts\python.exe" -c "import pdf2ppt; print(pdf2ppt.VERSION)"
$PortableName = "PDF2PPT-$Version-win64-portable"
$PortableDir = Join-Path $ReleaseRoot $PortableName
$SourceName = "PDF2PPT-$Version-source"
$SourceDir = Join-Path $ReleaseRoot $SourceName

Remove-Item -LiteralPath $PortableDir -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $SourceDir -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Path $PortableDir | Out-Null
New-Item -ItemType Directory -Path $SourceDir | Out-Null

$PortableFiles = @(
    ".\outputs\PDF2PPT.exe",
    ".\README.md",
    ".\LICENSE",
    ".\PRIVACY.md",
    ".\THIRD-PARTY-NOTICES.md",
    ".\CHANGELOG.md"
)
Copy-Item -LiteralPath $PortableFiles -Destination $PortableDir

$SourceFiles = @(
    ".\pdf2ppt.py",
    ".\test_pdf2ppt.py",
    ".\make_icon.py",
    ".\pdf2ppt.ico",
    ".\PDF2PPT.spec",
    ".\version_info.txt",
    ".\build.ps1",
    ".\release.ps1",
    ".\requirements-build.txt",
    ".\requirements-test.txt",
    ".\README.md",
    ".\LICENSE",
    ".\PRIVACY.md",
    ".\SECURITY.md",
    ".\CONTRIBUTING.md",
    ".\RELEASE_CHECKLIST.md",
    ".\THIRD-PARTY-NOTICES.md",
    ".\CHANGELOG.md",
    ".\.gitignore"
)
Copy-Item -LiteralPath $SourceFiles -Destination $SourceDir
Copy-Item -LiteralPath ".\.github" -Destination $SourceDir -Recurse

$SitePackages = & ".\.venv-build\Scripts\python.exe" -c "import sysconfig; print(sysconfig.get_paths()['purelib'])"
$LicenseDir = Join-Path $PortableDir "licenses"
New-Item -ItemType Directory -Path $LicenseDir | Out-Null

$LicensePackages = @(
    @{ Name = "pypdfium2"; Pattern = "pypdfium2-*.dist-info" },
    @{ Name = "python-pptx"; Pattern = "python_pptx-*.dist-info" },
    @{ Name = "Pillow"; Pattern = "pillow-*.dist-info" },
    @{ Name = "lxml"; Pattern = "lxml-*.dist-info" }
)
foreach ($Package in $LicensePackages) {
    $MetadataDir = Get-ChildItem -LiteralPath $SitePackages -Directory -Filter $Package.Pattern |
        Select-Object -First 1
    if ($MetadataDir) {
        $PackageDir = Join-Path $LicenseDir $Package.Name
        New-Item -ItemType Directory -Path $PackageDir | Out-Null
        Get-ChildItem -LiteralPath $MetadataDir.FullName -File |
            Where-Object { $_.Name -match "^(LICENSE|COPYING|NOTICE)" } |
            Copy-Item -Destination $PackageDir
        $NestedLicenses = Join-Path $MetadataDir.FullName "licenses"
        if (Test-Path $NestedLicenses) {
            Copy-Item -Path "$NestedLicenses\*" -Destination $PackageDir -Recurse
        }
    }
}

$PythonRoot = & ".\.venv-build\Scripts\python.exe" -c "import sys; print(sys.base_prefix)"
$PythonLicense = Join-Path $PythonRoot "LICENSE.txt"
if (Test-Path $PythonLicense) {
    Copy-Item -LiteralPath $PythonLicense -Destination (Join-Path $LicenseDir "CPython-LICENSE.txt")
}

$TclLicense = Join-Path $PythonRoot "tcl\tcl8.6\license.terms"
$TkLicense = Join-Path $PythonRoot "tcl\tk8.6\license.terms"
if (Test-Path $TclLicense) {
    Copy-Item -LiteralPath $TclLicense -Destination (Join-Path $LicenseDir "Tcl-LICENSE.txt")
}
if (Test-Path $TkLicense) {
    Copy-Item -LiteralPath $TkLicense -Destination (Join-Path $LicenseDir "Tk-LICENSE.txt")
}

$PortableZip = Join-Path $ReleaseRoot "$PortableName.zip"
$SourceZip = Join-Path $ReleaseRoot "$SourceName.zip"
Remove-Item -LiteralPath $PortableZip, $SourceZip -Force -ErrorAction SilentlyContinue
Compress-Archive -Path "$PortableDir\*" -DestinationPath $PortableZip -CompressionLevel Optimal
Compress-Archive -Path "$SourceDir\*" -DestinationPath $SourceZip -CompressionLevel Optimal

$Hashes = @(
    Get-FileHash -Algorithm SHA256 -LiteralPath $PortableZip
    Get-FileHash -Algorithm SHA256 -LiteralPath $SourceZip
    Get-FileHash -Algorithm SHA256 -LiteralPath ".\outputs\PDF2PPT.exe"
)
$HashLines = $Hashes | ForEach-Object { "$($_.Hash)  $([IO.Path]::GetFileName($_.Path))" }
$HashLines | Set-Content -LiteralPath (Join-Path $ReleaseRoot "SHA256SUMS.txt") -Encoding ascii

$OutputRoot = Join-Path $PSScriptRoot "outputs"
New-Item -ItemType Directory -Path $OutputRoot -Force | Out-Null
Copy-Item -LiteralPath $PortableZip, $SourceZip, (Join-Path $ReleaseRoot "SHA256SUMS.txt") `
    -Destination $OutputRoot -Force

Get-Item $PortableZip, $SourceZip, (Join-Path $ReleaseRoot "SHA256SUMS.txt") |
    Select-Object FullName, Length, LastWriteTime
