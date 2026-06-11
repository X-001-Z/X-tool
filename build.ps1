$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$Venv = Join-Path $PSScriptRoot ".venv-build"
$VenvPython = Join-Path $Venv "Scripts\python.exe"
$PyInstaller = Join-Path $Venv "Scripts\pyinstaller.exe"

if (-not (Test-Path $VenvPython)) {
    $PyLauncher = Get-Command py -ErrorAction SilentlyContinue
    if ($PyLauncher) {
        & $PyLauncher.Source -3.12 -m venv $Venv
    } else {
        $SystemPython = Get-Command python -ErrorAction Stop
        & $SystemPython.Source -m venv $Venv
    }
}

& $VenvPython -m pip install --upgrade pip
& $VenvPython -m pip install -r ".\requirements-test.txt"
& $VenvPython ".\make_icon.py"
& $VenvPython -m unittest -v ".\test_pdf2ppt.py"
& $PyInstaller --noconfirm --clean --distpath ".\outputs" --workpath ".\work\build" ".\PDF2PPT.spec"

Get-Item ".\outputs\PDF2PPT.exe" | Select-Object FullName, Length, LastWriteTime
