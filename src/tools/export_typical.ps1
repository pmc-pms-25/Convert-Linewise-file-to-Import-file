# Rebuild src\data\typical_data.js from sheet "NEW Typical" of Working_Nguyen.xlsx (a temp COPY is opened, read-only).
param([string]$Working = 'C:\Users\tructtt\Desktop\Convert Linewise file to Import file\Input\Working_Nguyen.xlsx')
$ErrorActionPreference = 'Stop'
$tmpCopy = Join-Path $env:TEMP ('convertmto-typical-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.xlsx')
Copy-Item -LiteralPath $Working $tmpCopy
$xl = New-Object -ComObject Excel.Application; $xl.DisplayAlerts = $false
try {
  $w = $xl.Workbooks.Open($tmpCopy, 0, $true); $t = $w.Worksheets.Item('NEW Typical')
  $n = $t.Cells.Item($t.Rows.Count, 1).End(-4162).Row
  $v = $t.Range("A1:P$n").Value2; $w.Close($false)
} finally { $xl.Quit(); Remove-Item -LiteralPath $tmpCopy -Force -ErrorAction SilentlyContinue }
function J($x) { if ($null -eq $x -or "$x" -eq '') { 'null' } else { $s = "$x"; $dd = 0.0; if ([double]::TryParse($s, [ref]$dd)) { $dd.ToString([Globalization.CultureInfo]::InvariantCulture) } else { '"' + ($s -replace '\\', '\\' -replace '"', '\"') + '"' } } }
$rows = New-Object System.Collections.Generic.List[string]; $conv = New-Object System.Collections.Generic.List[string]; $rules = [ordered]@{}
for ($r = 2; $r -le $n; $r++) {
  $cat = "$($v[$r,1])".Trim()
  if ($cat -ne '') {
    $cells = @((@($cat) + @(2..8 | ForEach-Object { $v[$r,$_] })) | ForEach-Object { J $_ }); $rows.Add('[' + ($cells -join ',') + ']')
    if (-not $rules.Contains($cat)) { $rules[$cat] = "$($v[$r,9])" }
  }
  if ("$($v[$r,15])" -ne '' -and $r -gt 1) { $conv.Add('[' + (J $v[$r,15]) + ',' + (J $v[$r,16]) + ']') }
}
"hdr: " + (($(for ($c = 1; $c -le 16; $c++) { "$c=$($v[1,$c])" })) -join ' | ')
$rules.GetEnumerator() | ForEach-Object { "$($_.Key): $($_.Value)" }
"rows=$($rows.Count) conv=$($conv.Count)"
$ruleJson = '{' + (($rules.GetEnumerator() | ForEach-Object { '"' + $_.Key + '":' + (J $_.Value) }) -join ',') + '}'
$js = "const TYPICAL=[`n" + ($rows -join ",`n") + "];`nconst CONV=[" + ($conv -join ',') + "];`nconst RULES=$ruleJson;"
[IO.File]::WriteAllText((Join-Path (Split-Path $PSScriptRoot -Parent) 'data\typical_data.js'), $js, (New-Object Text.UTF8Encoding($false)))
"conv: " + ($conv -join ' ')
$rows | Where-Object { $_ -match '"(OLET|COUP|UNIO|FTUB|CAP)"' } | Select-Object -First 10
