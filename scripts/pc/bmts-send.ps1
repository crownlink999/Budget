# BMTS: отправляет на приёмник Google Apps Script (-> GitHub input/) файлы из папки files (и её подпапок).
# Рядом нужен send.txt: 1-я строка = адрес веб-приложения (https://script.google.com/macros/s/.../exec), 2-я = KEY.
# Уже отправленные и не изменившиеся файлы повторно не шлются (sent.json).
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
try { [Net.WebRequest]::DefaultWebProxy.Credentials = [Net.CredentialCache]::DefaultCredentials } catch {}
$Dir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$Log   = Join-Path $Dir 'log.txt'
$Src   = Join-Path $Dir 'files'
$State = Join-Path $Dir 'sent.json'
function Log($m) { $l = '{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $m; Add-Content $Log $l -Encoding UTF8; Write-Host $l }
try {
    $cfg = Get-Content -LiteralPath (Join-Path $Dir 'send.txt') -Encoding UTF8 | Where-Object { $_.Trim() }
    $url = $cfg[0].Trim(); $key = $cfg[1].Trim()
    $sent = @{}
    if (Test-Path $State) { (Get-Content $State -Raw -Encoding UTF8 | ConvertFrom-Json).PSObject.Properties | ForEach-Object { $sent[$_.Name] = $_.Value } }
    $list = @(); $sig = @{}

    if (Test-Path -LiteralPath $Src) {
        foreach ($f in (Get-ChildItem -LiteralPath $Src -File -Recurse | Where-Object { $_.Extension -match '^\.(xlsx|xlsm|xls)$' })) {
            $s = '{0}|{1}' -f $f.Length, $f.LastWriteTimeUtc.Ticks
            if ($sent[$f.Name] -eq $s) { continue }
            $list += @{ name = $f.Name; b64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($f.FullName)) }
            $sig[$f.Name] = $s
        }
    }

    if (-not $list.Count) { Log 'send: nothing new'; exit 0 }
    $json = @{ key = $key; files = $list } | ConvertTo-Json -Depth 4 -Compress
    $bytes = [Text.Encoding]::UTF8.GetBytes($json)
    $r = Invoke-RestMethod -Uri $url -Method Post -Body $bytes -ContentType 'application/json; charset=utf-8' -TimeoutSec 300
    if (-not $r.ok) { throw ('server: ' + $r.error) }
    foreach ($k in $sig.Keys) { $sent[$k] = $sig[$k] }
    $sent | ConvertTo-Json | Set-Content $State -Encoding UTF8
    Log ('send: OK, files=' + $list.Count + ' (' + (($list | ForEach-Object { $_.name }) -join '; ') + ')')
}
catch { Log ('SEND ERROR: ' + $_.Exception.Message); exit 1 }
