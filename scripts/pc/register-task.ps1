$d = Split-Path -Parent $MyInvocation.MyCommand.Path

# запускающий файл: собрать файлы и отправить (ошибка сбора не мешает отправке)
$runner = Join-Path $d 'bmts-run.ps1'
$lines = @(
  '$d = Split-Path -Parent $MyInvocation.MyCommand.Path',
  'try { & (Join-Path $d "bmts-collect.ps1") } catch {}',
  'try { & (Join-Path $d "bmts-send.ps1") } catch {}'
)
Set-Content -LiteralPath $runner -Value $lines -Encoding UTF8

$arg = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $runner + '"'
$a = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $arg
# время запусков: поменяй под себя. Несколько попыток в день нужны на случай, если файл на сервере появился позже:
# уже отправленный файл повторно не шлётся.
$times = '07:35','08:30','10:00','14:00','16:00','18:00'
$t = $times | ForEach-Object { New-ScheduledTaskTrigger -Daily -At $_ }
$s = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew
Register-ScheduledTask -TaskName 'BMTS upload' -Action $a -Trigger $t -Settings $s -Force | Out-Null
'OK registered at ' + ($times -join ', ')
