# Rebuild src\data\lookup_data.js (config WorkGroup, project/facility lists, input templates) from Excel.
# Source files are COPIED to a temp folder first and only the copies are opened (read-only).
param(
  [string]$ProjectFacilityDir = 'C:\Users\tructtt\Desktop\Convert Linewise file to Import file\Input\ProjectFacility',
  [string]$Config = 'C:\Users\tructtt\Desktop\MM\Bridges\WorkGroup Bridges\plan_config_work_group (file config).xlsx',
  [string]$StrInput = '\\192.168.49.4\mm1-bda\04.Public\1.05.Planning\PMS Operation\1.Design\0.Design Master MTO\Structural MTO Master.xlsx',
  [string]$EitInputDir = 'C:\Users\tructtt\Desktop\Convert Linewise file to Import file\Input\EIT'
)
$ErrorActionPreference = 'Stop'
$OutJs = Join-Path (Split-Path $PSScriptRoot -Parent) 'data\lookup_data.js'
$Tmp = Join-Path $env:TEMP ('convertmto-src-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory "$Tmp\EIT" -Force | Out-Null
Copy-Item "$ProjectFacilityDir\sys_project_list.xlsx", "$ProjectFacilityDir\sys_facility_list.xlsx", "$ProjectFacilityDir\sys_facility_type.xlsx" $Tmp
Copy-Item $Config "$Tmp\config.xlsx"; Copy-Item $StrInput "$Tmp\str_input.xlsx"; Copy-Item "$EitInputDir\*.xlsx" "$Tmp\EIT"
$PF = $Tmp; $Cfg = "$Tmp\config.xlsx"; $StrIn = "$Tmp\str_input.xlsx"; $EitDir = "$Tmp\EIT"
$PfDate = (Get-Item "$ProjectFacilityDir\sys_facility_list.xlsx").LastWriteTime.ToString('yyyy-MM-dd')
function J($v) {
  if ($null -eq $v) { return 'null' }
  if ($v -is [double] -or $v -is [int]) { if ($v -is [int] -and $v -lt -2000000000) { return 'null' }; return ([double]$v).ToString('R', [Globalization.CultureInfo]::InvariantCulture) }
  $s = "$v".Trim(); if ($s -eq '' -or $s -eq 'NULL') { return 'null' }
  '"' + ($s -replace '\\', '\\' -replace '"', '\"' -replace "`r", '' -replace "`n", '\n' -replace "`t", ' ') + '"'
}
$xl = New-Object -ComObject Excel.Application; $xl.DisplayAlerts = $false
try {
  function Read-SheetValues($p) { $w = $xl.Workbooks.Open($p, 0, $true); $s = $w.Worksheets.Item(1); $ur = $s.UsedRange; $n = $ur.Row + $ur.Rows.Count - 1; $c = $ur.Column + $ur.Columns.Count - 1; $v = $s.Range($s.Cells.Item(1, 1), $s.Cells.Item($n, $c)).Value2; $w.Close($false); @{ V = $v; N = $n; C = $c } }
  function Hd($t, $row) { $h = @{}; for ($i = 1; $i -le $t.C; $i++) { $k = "$($t.V[$row,$i])".Trim().ToUpper(); if ($k -and -not $h.ContainsKey($k)) { $h[$k] = $i } }; $h }
  function Tbl($t, $cols) { $h = Hd $t 1; $rows = @(); for ($r = 2; $r -le $t.N; $r++) { $a = @(); $any = $false; foreach ($c in $cols) { $x = $t.V[$r,$h[$c]]; $j = J $x; if ($j -ne 'null') { $any = $true }; $a += $j }; if ($any) { $rows += '[' + ($a -join ',') + ']' } }; '[' + ($rows -join ",`n") + ']' }
  $proj = Tbl (Read-SheetValues "$PF\sys_project_list.xlsx") @('CODE', 'NAME', 'PROJECTTYPE')
  $fac = Tbl (Read-SheetValues "$PF\sys_facility_list.xlsx") @('PROJECTCODE', 'CODE', 'GROUPFACILITYCODE', 'FACILITYTYPE')
  $ftype = Tbl (Read-SheetValues "$PF\sys_facility_type.xlsx") @('PROJECTTYPE', 'CODE', 'NAME', 'DESCRIPTION')
  $wg = Tbl (Read-SheetValues $Cfg) @('DISCIPLINECODE', 'DISCIPLINENAME', 'PROJECTTYPE', 'FACILITYTYPE', 'WORKGROUPCODE', 'WORKGROUPNAME', 'ISACTIVE')
  # input templates: flag row + header row (stop at last named header)
  function Tpl($p) { $t = Read-SheetValues $p; $hr = 0; for ($r = 1; $r -le 10; $r++) { for ($c = 1; $c -le $t.C; $c++) { if ("$($t.V[$r,$c])".Trim() -eq 'ProjectCode') { $hr = $r; break } }; if ($hr) { break } }
    $last = 0; for ($c = 1; $c -le $t.C; $c++) { if ("$($t.V[$hr,$c])".Trim()) { $last = $c } }
    $f = @(); $h = @(); $e = @(); for ($c = 1; $c -le $last; $c++) { $f += J $(if ($hr -gt 1) { $t.V[($hr - 1),$c] } else { $null }); $h += J $t.V[$hr,$c]; $e += J $t.V[($hr + 1),$c] }
    "{flags:[$($f -join ',')],heads:[$($h -join ',')],example:[$($e -join ',')]}" }
  $tp = @{
    STR = Tpl $StrIn
    CABLE = Tpl "$EitDir\EIT MTO Master ( cable schedule ).xlsx"
    TRAY = Tpl "$EitDir\EIT MTO Master ( Ladder).xlsx"
    SUPPORT = Tpl "$EitDir\EIT MTO Master ( Support).xlsx"
    EQUIPMENT = Tpl "$EitDir\EIT MTO Master (Equipment).xlsx"
    TUBING = Tpl "$EitDir\EIT MTO Master (Tubing).xlsx"
  }
  $js = "const LOOKUP_EMBED={date:'$((Get-Item $Config).LastWriteTime.ToString('yyyy-MM-dd'))',pfdate:'$PfDate',`nproj:$proj,`nfac:$fac,`nftype:$ftype,`nwg:$wg};`n"
  $js += "const INPUT_TPL={" + (($tp.GetEnumerator() | ForEach-Object { "$($_.Key):$($_.Value)" }) -join ",`n") + "};`n"
  [IO.File]::WriteAllText($OutJs, $js, (New-Object Text.UTF8Encoding($false)))
  "lookup_data.js: $((Get-Item $OutJs).Length) bytes"
} finally { $xl.Quit(); Remove-Item -LiteralPath $Tmp -Recurse -Force -ErrorAction SilentlyContinue }
