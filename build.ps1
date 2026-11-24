$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
python -m PyInstaller --noconfirm --onefile --windowed --name 'KalibriOffice' --add-data "$PSScriptRoot/index.html;." --add-data "$PSScriptRoot/features.js;." --collect-all webview app.py
if ($LASTEXITCODE -ne 0) { throw 'Сборка завершилась с ошибкой' }
Copy-Item -LiteralPath README.md -Destination dist/README.md
Copy-Item -LiteralPath LICENSE -Destination dist/LICENSE
$launcherFiles = @{'Kalibri Write.cmd'='write';'Kalibri Calc.cmd'='sheet';'Kalibri Slides.cmd'='slides'}
foreach ($entry in $launcherFiles.GetEnumerator()) {
    $body = '@echo off' + "`r`n" + 'start "" "%~dp0KalibriOffice.exe" --kind ' + $entry.Value + "`r`n"
    [IO.File]::WriteAllText((Join-Path $PSScriptRoot "dist/$($entry.Key)"), $body, [Text.Encoding]::ASCII)
}
