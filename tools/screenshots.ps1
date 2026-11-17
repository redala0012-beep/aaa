# README / presentation screenshots via headless Microsoft Edge (no install needed)
# Usage: start a local server that serves this repo at http://localhost:8765, then
#   powershell -ExecutionPolicy Bypass -File tools\screenshots.ps1
param([string]$Base = 'http://localhost:8765')
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$out = Join-Path $root 'docs\screenshots'
New-Item -ItemType Directory -Force $out | Out-Null
$edge = @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $edge) { throw 'Microsoft Edge not found' }
$shots = @(
  @{ f = '01_who.png';          u = '/site/index.html#/who';                              w = 1280; h = 760 },
  @{ f = '02_timeline.png';     u = '/site/index.html?aud=soldier#/';                     w = 1280; h = 1500 },
  @{ f = '03_detail_full.png';  u = '/site/index.html?aud=officer#/b/DSS';                w = 1280; h = 4200 },
  @{ f = '04_edu_cover.png';    u = '/site/index.html?aud=officer&edu=1#/b/INC';          w = 1280; h = 720 },
  @{ f = '05_edu_decision.png'; u = '/site/index.html?aud=officer&edu=6&reveal=1#/b/DSS'; w = 1280; h = 720 },
  @{ f = '06_edu_quiz.png';     u = '/site/index.html?aud=officer&edu=14&reveal=1#/b/INC';w = 1280; h = 720 },
  @{ f = '07_memorial.png';     u = '/site/index.html?aud=public#/b/YPD';                 w = 1280; h = 800 },
  @{ f = '08_log.png';          u = '/site/index.html?aud=public#/log';                   w = 1280; h = 800 },
  @{ f = '09_converter.png';    u = '/converter/converter.html';                          w = 1280; h = 900 }
)
foreach ($s in $shots) {
  $prof = Join-Path $env:TEMP ('edge-shot-' + [guid]::NewGuid().ToString('N'))   # fresh profile = no saved audience
  $file = Join-Path $out $s.f
  $args = @('--headless=new', '--disable-gpu', '--hide-scrollbars', '--no-first-run', '--no-default-browser-check',
            "--user-data-dir=$prof", '--virtual-time-budget=4000', "--window-size=$($s.w),$($s.h)",
            "--screenshot=$file", ($Base + $s.u))
  $p = Start-Process -FilePath $edge -ArgumentList $args -PassThru -WindowStyle Hidden
  if (-not $p.WaitForExit(60000)) { $p.Kill() }
  Remove-Item -Recurse -Force $prof -ErrorAction SilentlyContinue
  if (Test-Path $file) { '{0}  {1:N0}KB' -f $s.f, ((Get-Item $file).Length / 1KB) } else { "$($s.f) FAILED" }
}
