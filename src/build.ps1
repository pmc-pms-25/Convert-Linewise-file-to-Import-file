# Build Convert MTO: page HTML (online + local), self-test pages and the exe launcher.
# Usage: powershell -ExecutionPolicy Bypass -File src\build.ps1
# Outputs: ..\Convert MTO.html, ..\Convert MTO.exe ; intermediates in src\build\ (git-ignored)
$ErrorActionPreference = 'Stop'
$S = $PSScriptRoot
$Root = Split-Path $S -Parent
$B = Join-Path $S 'build'
New-Item -ItemType Directory $B -Force | Out-Null
$utf8 = New-Object Text.UTF8Encoding($false)

$tpl = [IO.File]::ReadAllText("$S\convert_ui_template.html")
$tpl = $tpl.Replace('/*MASTER*/', [IO.File]::ReadAllText("$S\master.js"))
$data = [IO.File]::ReadAllText("$S\data\typical_data.js") + "`n" + [IO.File]::ReadAllText("$S\data\lookup_data.js")
$page = $tpl.Replace('/*DATA*/', $data)

# 1) artifact / online version (no html skeleton; SheetJS from cdnjs)
[IO.File]::WriteAllText("$B\convert-import.html", $page, $utf8)

# 2) local version: own head so file:// opens with correct encoding
$head = "<!doctype html><html lang=""vi""><head><meta charset=""utf-8""><meta name=""viewport"" content=""width=device-width,initial-scale=1,viewport-fit=cover""><style>:root{color-scheme:light}body{margin:0}[hidden]{display:none!important}</style></head><body>`n"
$local = $head + $page + "`n</body></html>"
[IO.File]::WriteAllText("$Root\Convert MTO.html", $local, $utf8)

# 3) self-test page (used by tools\test_page.ps1): ?f=<file url>[&mode=str|eit][&b64=1]
$test = [IO.File]::ReadAllText("$S\tools\selftest.js")
[IO.File]::WriteAllText("$B\convert-import-test.html", $head + $tpl.Replace('/*DATA*/', $data + "`n" + $test) + "`n</body></html>", $utf8)

# 4) exe: same page, SheetJS as a separate file (inlining breaks: the lib contains "<!--")
$cdn = '<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>'
if (-not $local.Contains($cdn)) { throw 'cdn script tag not found' }
$app = $local.Replace($cdn, '<script src="xlsx.full.min.js"></script>').Replace('<title>Convert MTO Import</title>', '<title>Convert MTO</title>')
[IO.File]::WriteAllText("$B\app.html", $app, $utf8)
[IO.File]::WriteAllText("$B\app-test.html", $app.Replace('</body></html>', "<script>`n$test`nwindow.__selftest();</script></body></html>"), $utf8)
Copy-Item "$S\exe\xlsx.full.min.js" "$B\xlsx.full.min.js" -Force
$csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
& $csc /nologo /target:winexe /optimize+ "/out:$Root\Convert MTO.exe" "/win32icon:$S\exe\app.ico" "/resource:$B\app.html,app.html" "/resource:$S\exe\xlsx.full.min.js,xlsx.full.min.js" /r:System.Windows.Forms.dll "$S\exe\Launcher.cs"
if ($LASTEXITCODE) { throw "csc failed ($LASTEXITCODE)" }
"Convert MTO.html: $((Get-Item "$Root\Convert MTO.html").Length) bytes"
"Convert MTO.exe : $((Get-Item "$Root\Convert MTO.exe").Length) bytes"