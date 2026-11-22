$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
python -m PyInstaller --noconfirm --onefile --windowed --name 'Orbit Office' --add-data "$PSScriptRoot/index.html;." --collect-all webview app.py
if ($LASTEXITCODE -ne 0) { throw 'Сборка завершилась с ошибкой' }
Copy-Item -LiteralPath README.md -Destination dist/README.md
Copy-Item -LiteralPath LICENSE -Destination dist/LICENSE
$launcherFiles = @{'Orbit Write.cmd'='write';'Orbit Calc.cmd'='sheet';'Orbit Slides.cmd'='slides'}
foreach ($entry in $launcherFiles.GetEnumerator()) {
    $body = '@echo off' + "`r`n" + 'start "" "%~dp0Orbit Office.exe" --kind ' + $entry.Value + "`r`n"
    [IO.File]::WriteAllText((Join-Path $PSScriptRoot "dist/$($entry.Key)"), $body, [Text.Encoding]::ASCII)
}
