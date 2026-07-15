<#
  embed_export.ps1 — 운세 PPTX에 글꼴을 임베드하고 슬라이드를 PNG로 내보낸다.
  (PowerPoint COM 사용. 이 PC에 본문 글꼴 "a아시아명조M"이 설치돼 있어야 임베드됨.)

  사용법 (이 폴더에서 실행, 코드는 문자열로 — 앞자리 0 보존):
    powershell -ExecutionPolicy Bypass -File embed_export.ps1 0709 0710 0711
    또는 세션에서:  & .\embed_export.ps1 '0709','0710'

  각 코드에 대해:
    ../4_PPTX/CODE.pptx  → 글꼴 임베드하여 같은 파일로 저장(SaveAs, EmbedTrueTypeFonts)
    ../5_PNG/CODE_png/slide_01.png ~ slide_13.png  (1500x2362) 로 내보내기

  견고화: COM 호출이 "Call was rejected by callee"(RPC_E_CALL_REJECTED)로 거부되면
  잠시 쉬고 재시도한다. deck별로 try/catch 격리하여 하나가 실패해도 나머지는 계속 처리.
#>
param([Parameter(Mandatory=$true)][string[]]$Codes)

$base    = $PSScriptRoot
$pptxDir = Join-Path $base "..\4_PPTX"
$pngDir  = Join-Path $base "..\5_PNG"

# COM 호출 재시도 (서버가 바쁠 때 RPC_E_CALL_REJECTED 등 대비)
function Invoke-Com {
  param([scriptblock]$Action, [int]$Tries = 8, [int]$DelayMs = 600)
  for ($k = 1; $k -le $Tries; $k++) {
    try { return & $Action }
    catch {
      if ($k -eq $Tries) { throw }
      Start-Sleep -Milliseconds $DelayMs
    }
  }
}

$pp = New-Object -ComObject PowerPoint.Application
$ok = 0; $fail = 0
try {
  $pp.DisplayAlerts = 1
  foreach ($c in $Codes) {
    $path = (Resolve-Path (Join-Path $pptxDir "$c.pptx") -ErrorAction SilentlyContinue)
    if (-not $path) { Write-Output "$c : ../4_PPTX/$c.pptx 없음 — 건너뜀"; $fail++; continue }
    $path = $path.Path
    try {
      # 1) embed fonts (SaveAs: 24=ppSaveAsOpenXMLPresentation, -1=msoTrue)
      $pres = Invoke-Com { $pp.Presentations.Open($path, $false, $false, $false) }
      Invoke-Com { $pres.SaveAs($path, 24, -1) }
      Start-Sleep -Milliseconds 400
      # 2) export PNG
      $out = Join-Path $pngDir "${c}_png"
      if (-not (Test-Path $out)) { New-Item -ItemType Directory $out | Out-Null }
      Get-ChildItem "$out\slide_*.png" -ErrorAction SilentlyContinue | Remove-Item -Force
      $n = Invoke-Com { $pres.Slides.Count }
      for ($i = 1; $i -le $n; $i++) {
        $f = Join-Path $out ("slide_{0:D2}.png" -f $i)
        Invoke-Com { $pres.Slides.Item($i).Export($f, "PNG", 1500, 2362) }
      }
      Invoke-Com { $pres.Close() }
      Start-Sleep -Milliseconds 300
      Write-Output "$c : 글꼴 임베드 + PNG $n장 완료"
      $ok++
    } catch {
      Write-Output "$c : 실패 — $($_.Exception.Message)"
      $fail++
      try { Invoke-Com { $pres.Close() } -Tries 3 } catch {}
    }
  }
} finally {
  try { Invoke-Com { $pp.Quit() } } catch {}
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null
}
Write-Output "=== 완료 $ok / 실패 $fail ==="
