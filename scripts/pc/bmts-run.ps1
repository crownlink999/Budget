$d = Split-Path -Parent $MyInvocation.MyCommand.Path
try { & (Join-Path $d "bmts-collect.ps1") } catch {}
try { & (Join-Path $d "bmts-send.ps1") } catch {}
