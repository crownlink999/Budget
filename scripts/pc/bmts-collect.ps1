# BMTS: берёт самый свежий (по дате изменения) подходящий файл ВМЕСТЕ СО ВСЕМИ ПОДПАПКАМИ (список в sources.txt)
# и копирует в локальную папку files рядом со скриптом (для каждого источника своя подпапка src1, src2...).
# Никуда ничего не отправляет.
# sources.txt: по строке на источник, поля через |  :  ПУТЬ | МАСКА | СЛОВА-ИСКЛЮЧЕНИЯ | дата
#   В ПУТИ можно писать {ГОД}: S:\...\{ГОД} год\Отчеты\ME2N  (подставится текущий год по дате ПК, 2027 подхватится сам в январе)
#   МАСКА (необязательно): например *me2n*  (* = любые символы; регистр не важен)
#   СЛОВА-ИСКЛЮЧЕНИЯ (необязательно): через запятую; файлы с этими словами в имени пропускаются
#   дата (необязательно): слово «дата» = брать только файлы, у которых в имени есть дата (16.11.2026, 16_11_26...)
#   Строка, которая начинается с + , добавляет ещё одну папку к ПРЕДЫДУЩЕМУ источнику: файл ищется во всех папках
#   источника, и берётся самый свежий по дате изменения из всех.
# Строки с # игнорируются.
$ErrorActionPreference = 'Stop'
$Dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Log = Join-Path $Dir 'log.txt'
$Out = Join-Path $Dir 'files'
$Cfg = Join-Path $Dir 'sources.txt'

function Log($m) {
    $line = '{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $m
    Add-Content -Path $Log -Value $line -Encoding UTF8
    Write-Host $line
}

# самый свежий подходящий файл для одной папки источника (или $null)
function Find-Best($m) {
    $cands = @($m.Folder)
    if ($m.Folder -match '(?i)\{(\u0413\u041E\u0414|YEAR)\}') {
        $cands = @($m.Folder -replace '(?i)\{(\u0413\u041E\u0414|YEAR)\}', ('' + (Get-Date).Year))
    }
    $excl = $m.Excl
    $needDate = $m.NeedDate
    $found = $null
    foreach ($cand in $cands) {
        $p = $cand
        if ($p -match '[*?]') {
            $res = @(Resolve-Path -Path $p -ErrorAction SilentlyContinue | Sort-Object Path -Descending)
            if (-not $res.Count) { Log ('folder not found: ' + $p); continue }
            $p = $res[0].ProviderPath
        }
        if (-not (Test-Path -LiteralPath $p)) { Log ('folder not reachable: ' + $p); continue }
        $found = Get-ChildItem -LiteralPath $p -File -Recurse -Filter $m.Mask -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -match '^\.(xlsx|xlsm|xls)$' -and $_.Name -notlike '~$*' } |
            Where-Object { $fn = $_.Name; -not ($excl | Where-Object { $fn.IndexOf($_, [StringComparison]::OrdinalIgnoreCase) -ge 0 }) } |
            Where-Object { (-not $needDate) -or ($_.BaseName -match '\d{1,2}[._ -]\d{1,2}[._ -]\d{2,4}') } |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($found) { break }
    }
    return $found
}

try {
    if ((Test-Path $Log) -and (Get-Item $Log).Length -gt 500KB) { Remove-Item $Log -Force }
    if (-not (Test-Path -LiteralPath $Cfg)) { throw 'sources.txt not found next to the script' }
    $lines = Get-Content -LiteralPath $Cfg -Encoding UTF8 | Where-Object { $_.Trim() -and -not $_.Trim().StartsWith('#') }
    if (-not $lines) { throw 'sources.txt is empty' }
    if (-not (Test-Path -LiteralPath $Out)) { New-Item -ItemType Directory -Path $Out | Out-Null }

    # разбор sources.txt: источник = одна или несколько папок (строки с + добавляются к предыдущему источнику)
    $groups = New-Object System.Collections.ArrayList
    foreach ($ln in $lines) {
        $t = $ln.Trim()
        $isExtra = $t.StartsWith('+')
        if ($isExtra) { $t = $t.Substring(1).Trim() }
        $parts = $t.Split('|')
        $m = @{ Folder = $parts[0].Trim(); Mask = '*'; Excl = @(); NeedDate = $false }
        if ($parts.Count -ge 2 -and $parts[1].Trim()) { $m.Mask = $parts[1].Trim() }
        if ($parts.Count -ge 3 -and $parts[2].Trim()) { $m.Excl = @($parts[2].Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
        if ($parts.Count -ge 4 -and $parts[3].Trim() -match '(?i)^(\u0434\u0430\u0442\u0430|date)$') { $m.NeedDate = $true }
        if ($isExtra -and $groups.Count -gt 0) { [void]$groups[$groups.Count - 1].Add($m) }
        else {
            $g = New-Object System.Collections.ArrayList
            [void]$g.Add($m)
            [void]$groups.Add($g)
        }
    }

    $n = 0; $idx = 0
    foreach ($g in $groups) {
        $idx++
        $sub = Join-Path $Out ('src' + $idx)
        $where = (($g | ForEach-Object { $_.Folder }) -join ' | ')
        try {
            $best = $null
            foreach ($m in $g) {
                $b = Find-Best $m
                if ($b -and ((-not $best) -or ($b.LastWriteTime -gt $best.LastWriteTime))) { $best = $b }
            }
            if (-not $best) { Log ('no suitable Excel file in: ' + $where + ' (searched all subfolders)'); continue }
            if (-not (Test-Path -LiteralPath $sub)) { New-Item -ItemType Directory -Path $sub | Out-Null }
            Get-ChildItem -LiteralPath $sub -Filter '*.part' -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
            $name = $best.Name
            $dst  = Join-Path $sub $name
            $part = $dst + '.' + [guid]::NewGuid().ToString('N').Substring(0, 6) + '.part'
            try {
                $src = [IO.File]::Open($best.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
                try { $f = [IO.File]::Create($part); try { $src.CopyTo($f) } finally { $f.Dispose() } } finally { $src.Dispose() }
                Move-Item -LiteralPath $part -Destination $dst -Force
                (Get-Item -LiteralPath $dst).LastWriteTime = $best.LastWriteTime
            }
            catch {
                Log ('cannot copy ' + $best.FullName + ': ' + $_.Exception.Message)
                if (Test-Path -LiteralPath $part) { Remove-Item -LiteralPath $part -Force -ErrorAction SilentlyContinue }
                continue
            }
            # remove older files of this source from its local subfolder
            Get-ChildItem -LiteralPath $sub -File | Where-Object { $_.Name -ne $name } | Remove-Item -Force
            Log ('saved: ' + $name + ' (' + $best.LastWriteTime.ToString('yyyy-MM-dd HH:mm') + ') from ' + $best.DirectoryName)
            $n++
        }
        catch { Log ('ERROR [' + $where + ']: ' + $_.Exception.Message) }
    }
    Log ('done: ' + $n + ' file(s) in ' + $Out)
}
catch {
    Log ('ERROR: ' + $_.Exception.Message)
    exit 1
}
