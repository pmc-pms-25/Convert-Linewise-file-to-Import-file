# Headless Edge self-test: opens a build\*-test.html page with ?f=<file>, waits for the result, optional -B64 saves the output xlsx to build\out\
param([string]$Page = 'convert-import-test.html', [string]$File, [string]$Name, [string]$Query = '', [switch]$B64, [int]$MaxSec = 600)
$ErrorActionPreference = 'Stop'
$S = Join-Path (Split-Path $PSScriptRoot -Parent) 'build'
$edge = 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
$url = ([Uri]"$S\$Page").AbsoluteUri + '?f=' + [Uri]::EscapeDataString(([Uri]$File).AbsoluteUri) + $Query + $(if ($B64) { '&b64=1' } else { '' })
$port = 9333
$p = Start-Process $edge -ArgumentList @('--headless=new','--disable-gpu','--no-first-run','--allow-file-access-from-files',"--user-data-dir=`"$env:TEMP\convertmto-edgetest`"","--remote-debugging-port=$port",'--remote-allow-origins=*',"`"$url`"") -PassThru
try {
  $ws = $null
  for ($i = 0; $i -lt 30 -and -not $ws; $i++) { Start-Sleep -Milliseconds 500; try { $ws = (@(Invoke-RestMethod "http://127.0.0.1:$port/json") | ForEach-Object { $_ } | Where-Object { $_.type -eq 'page' } | Select-Object -First 1).webSocketDebuggerUrl } catch {} }
  if (-not $ws) { throw 'no devtools' }
  $c = New-Object Net.WebSockets.ClientWebSocket
  $c.ConnectAsync([Uri]$ws, [Threading.CancellationToken]::None).Wait()
  $id = 0
  function Eval($expr) {
    $script:id++
    $msg = @{ id = $script:id; method = 'Runtime.evaluate'; params = @{ expression = $expr; returnByValue = $true } } | ConvertTo-Json -Compress
    $b = [Text.Encoding]::UTF8.GetBytes($msg)
    $c.SendAsync([ArraySegment[byte]]$b, 'Text', $true, [Threading.CancellationToken]::None).Wait()
    $ms = New-Object IO.MemoryStream; $buf = New-Object byte[] 1048576
    do { $r = $c.ReceiveAsync([ArraySegment[byte]]$buf, [Threading.CancellationToken]::None).Result; $ms.Write($buf, 0, $r.Count) } while (-not $r.EndOfMessage)
    $j = [Text.Encoding]::UTF8.GetString($ms.ToArray()) | ConvertFrom-Json
    $j.result.result.value
  }
  $t0 = Get-Date; $txt = ''
  while (((Get-Date) - $t0).TotalSeconds -lt $MaxSec) {
    Start-Sleep -Seconds 3
    $state = Eval "(document.getElementById('selftest')||{}).textContent ? 'DONE' : (document.getElementById('fmeta').textContent+' | '+document.getElementById('runMsg').textContent)"
    if ($state -eq 'DONE') { break }
  }
  "elapsed: $([int]((Get-Date) - $t0).TotalSeconds)s"
  $res = Eval "document.getElementById('selftest').textContent.split('\nB64 ')[0]"
  $res
  $mem = Eval "performance.memory ? Math.round(performance.memory.usedJSHeapSize/1048576)+' MB heap' : ''"; $mem
  if ($B64) {
    $len = Eval "document.getElementById('selftest').textContent.split('\nB64 ')[1].length"
    $sb = New-Object Text.StringBuilder
    for ($o = 0; $o -lt $len; $o += 500000) { [void]$sb.Append((Eval "document.getElementById('selftest').textContent.split('\nB64 ')[1].substr($o,500000)")) }
    $bytes = [Convert]::FromBase64String($sb.ToString()); New-Item -ItemType Directory "$S\out" -Force | Out-Null; [IO.File]::WriteAllBytes("$S\out\$Name.xlsx", $bytes); "saved build\out\$Name.xlsx ($($bytes.Length) bytes)"
  }
} finally { try { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue } catch {}; Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" | Where-Object { $_.CommandLine -like '*convertmto-edgetest*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue } }
