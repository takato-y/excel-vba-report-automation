# excel_job.ps1 - drive real Excel over COM for build and test (ASCII-only source).
# All Japanese text (sheet names, module code, arguments) arrives through a UTF-8 JSON
# spec, because Windows PowerShell 5.1 misreads non-ASCII in .ps1 files without a BOM.
# Cell values are written by VBA, not from PowerShell (per-cell COM writes are unreliable).
#
# spec keys:
#   book       : workbook path (opened if "open" is true, otherwise a new workbook)
#   open       : true to open an existing book
#   modules    : [ { path, type } ]  type 1 = standard module, 2 = class module
#   init       : macro to run after modules are added (optional)
#   remove     : module names to remove before saving (optional)
#   testCode   : VBA code injected as a temporary module (removed before saving)
#   call       : { macro, args[] }  return value is written to result.status
#   dump       : [ { sheet, out } ]  UsedRange values written as UTF-8 TSV
#   saveAs     : path to save as .xlsm ("" = do not save)
#   result     : path of the result JSON
param([Parameter(Mandatory=$true)][string]$Spec)
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false
$utf8 = New-Object System.Text.UTF8Encoding $false
$res = [ordered]@{ ok=$false; status=$null; error=$null; modules=@() }
$x = $null; $wb = $null

function Add-ModuleFromFile($wb, $path, $type) {
  $raw = [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8)
  $name = $null
  $body = New-Object System.Collections.Generic.List[string]
  foreach ($line in ($raw -split "`r?`n")) {
    if ($line -cmatch '^Attribute VB_Name = "([^"]+)"') { $name = $Matches[1]; continue }
    if ($line -cmatch '^(VERSION 1\.0 CLASS|BEGIN|  MultiUse|END\s*$|Attribute )') { continue }
    $body.Add($line)
  }
  $comp = $wb.VBProject.VBComponents.Add([int]$type)
  $comp.Name = $name
  $comp.CodeModule.AddFromString(($body -join "`r`n"))
  return $name
}

try {
  $cfg = Get-Content -Raw -LiteralPath $Spec -Encoding UTF8 | ConvertFrom-Json
  $x = New-Object -ComObject Excel.Application
  $x.Visible = $false
  $x.DisplayAlerts = $false
  $x.EnableEvents = $false
  $x.AskToUpdateLinks = $false

  if ($cfg.open) { $wb = $x.Workbooks.Open([string]$cfg.book) }
  else { $wb = $x.Workbooks.Add() }

  foreach ($m in $cfg.modules) {
    $res.modules += (Add-ModuleFromFile $wb ([string]$m.path) ([int]$m.type))
  }
  $tmp = $null
  if ($cfg.testCode) {
    $tmp = $wb.VBProject.VBComponents.Add(1)
    $tmp.Name = "modTestDriver"
    $tmp.CodeModule.AddFromString([string]$cfg.testCode)
  }
  if ($cfg.init) { $x.Run([string]$cfg.init) | Out-Null }

  if ($cfg.call) {
    $a = @($cfg.call.args)
    switch ($a.Count) {
      0 { $ret = $x.Run([string]$cfg.call.macro) }
      1 { $ret = $x.Run([string]$cfg.call.macro, $a[0]) }
      2 { $ret = $x.Run([string]$cfg.call.macro, $a[0], $a[1]) }
      3 { $ret = $x.Run([string]$cfg.call.macro, $a[0], $a[1], $a[2]) }
      4 { $ret = $x.Run([string]$cfg.call.macro, $a[0], $a[1], $a[2], $a[3]) }
      5 { $ret = $x.Run([string]$cfg.call.macro, $a[0], $a[1], $a[2], $a[3], $a[4]) }
      default { throw "too many args" }
    }
    $res.status = [string]$ret
  }
  $x.CalculateFull()

  foreach ($d in $cfg.dump) {
    $ws = $wb.Worksheets.Item([string]$d.sheet)
    $ur = $ws.UsedRange
    $r0 = $ur.Row; $c0 = $ur.Column
    $nr = $ur.Rows.Count; $nc = $ur.Columns.Count
    $vals = $ur.Value2
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("#origin`t$r0`t$c0`n")
    for ($i = 1; $i -le $nr; $i++) {
      $cells = @()
      for ($j = 1; $j -le $nc; $j++) {
        if ($nr -eq 1 -and $nc -eq 1) { $v = $vals } else { $v = $vals[$i, $j] }
        if ($v -eq $null) { $s = "" } else { $s = [string]$v }
        $cells += ($s -replace "[`t`r`n]", " ")
      }
      [void]$sb.Append(($cells -join "`t") + "`n")
    }
    [IO.File]::WriteAllText([string]$d.out, $sb.ToString(), $utf8)
  }

  if ($tmp -ne $null) { $wb.VBProject.VBComponents.Remove($tmp) }
  foreach ($nm in $cfg.remove) { $wb.VBProject.VBComponents.Remove($wb.VBProject.VBComponents.Item([string]$nm)) }
  if ($cfg.saveAs) { $wb.SaveAs([string]$cfg.saveAs, 52) }
  $res.ok = $true
}
catch { $res.error = $_.Exception.Message + " @line=" + $_.InvocationInfo.ScriptLineNumber }
finally {
  if ($wb -ne $null) { try { $wb.Close($false) } catch {} }
  if ($x  -ne $null) { try { $x.Quit() } catch {} }
}
$json = $res | ConvertTo-Json -Compress -Depth 5
if ($cfg -and $cfg.result) { [IO.File]::WriteAllText([string]$cfg.result, $json, $utf8) }
$json
