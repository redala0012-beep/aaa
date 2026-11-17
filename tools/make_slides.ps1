# 발표자료(docs/발표자료.pptx, .pdf)를 PowerPoint 자동화로 생성합니다.
# 사용: powershell -ExecutionPolicy Bypass -File tools\make_slides.ps1 [-QaDir 폴더]
# 필요: Microsoft PowerPoint (설치만 되어 있으면 됨). 이 파일은 UTF-8(BOM)으로 저장해야 한글이 깨지지 않습니다.
param([string]$QaDir = '')
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$shots = Join-Path $root 'docs\screenshots'
$photos = Join-Path $root 'site\images\full'
$outPptx = Join-Path $root 'docs\발표자료.pptx'
$outPdf = Join-Path $root 'docs\발표자료.pdf'

# ---- 색 (사이트의 '차분한 화면'과 같은 계열) ----
$PAL = @{ red='8E2A2F'; redD='6F1F24'; gold='B8963A'; goldL='F7F1E1'; ink='26262A'; sub='66666E'; bg='F7F6F2'; line='E2DDD2'; white='FFFFFF'; blue='3D5A91'; blueL='EEF2FA'; green='2E7D4F'; greenL='E8F3EC' }
function C([string]$hex) { $r = [Convert]::ToInt32($hex.Substring(0,2),16); $g = [Convert]::ToInt32($hex.Substring(2,2),16); $b = [Convert]::ToInt32($hex.Substring(4,2),16); return $r + 256*$g + 65536*$b }
$FONT = 'Malgun Gothic'

function TB($sl, [double]$x, [double]$y, [double]$w, [double]$h, [string]$text, [double]$size, [string]$hex, [bool]$bold = $false, [int]$align = 1, [int]$anchor = 1) {
  $s = $sl.Shapes.AddTextbox(1, $x, $y, $w, $h)
  $tf = $s.TextFrame; $tf.WordWrap = -1; $tf.AutoSize = 0
  $tf.MarginLeft = 0; $tf.MarginRight = 0; $tf.MarginTop = 0; $tf.MarginBottom = 0; $tf.VerticalAnchor = $anchor
  $r = $tf.TextRange; $r.Text = $text
  $r.Font.Name = $FONT; $r.Font.NameFarEast = $FONT; $r.Font.Size = $size
  $r.Font.Bold = $(if ($bold) { -1 } else { 0 }); $r.Font.Color.RGB = (C $hex)
  $r.ParagraphFormat.Alignment = $align; $r.ParagraphFormat.SpaceWithin = 1.1
  # 한글을 음절이 아닌 단어 단위로 줄바꿈 ("한글 단어 잘림 허용" 끄기)
  $r.ParagraphFormat.WordWrap = -1
  $s.TextFrame2.TextRange.ParagraphFormat.WordWrap = -1
  return $s
}
function BX($sl, [int]$type, [double]$x, [double]$y, [double]$w, [double]$h, [string]$fill, [string]$lineHex = '') {
  $s = $sl.Shapes.AddShape($type, $x, $y, $w, $h)
  $s.Fill.Solid(); $s.Fill.ForeColor.RGB = (C $fill)
  if ($lineHex) { $s.Line.Visible = -1; $s.Line.ForeColor.RGB = (C $lineHex); $s.Line.Weight = 1 } else { $s.Line.Visible = 0 }
  $s.Shadow.Visible = 0
  if ($type -eq 5) { $s.Adjustments.Item(1) = 0.06 }
  if ($s.HasTextFrame) { $s.TextFrame.TextRange.Text = '' }
  return $s
}
function PIC($sl, [string]$path, [double]$x, [double]$y, [double]$w, [double]$h, [bool]$cover = $false, [string]$border = '') {
  $p = $sl.Shapes.AddPicture($path, 0, -1, $x, $y, -1, -1)
  $p.LockAspectRatio = -1
  $pw = $p.Width; $ph = $p.Height; $want = $w / $h; $have = $pw / $ph
  if ($cover) {
    if ($have -gt $want) { $cut = $pw - $ph * $want; $p.PictureFormat.CropLeft = $cut / 2; $p.PictureFormat.CropRight = $cut / 2 }
    else { $cut = $ph - $pw / $want; $p.PictureFormat.CropTop = $cut * 0.25; $p.PictureFormat.CropBottom = $cut * 0.75 }
    $p.Width = $w; $p.Left = $x; $p.Top = $y
  } else {
    if ($have -gt $want) { $p.Width = $w } else { $p.Height = $h }
    $p.Left = $x + ($w - $p.Width) / 2; $p.Top = $y + ($h - $p.Height) / 2
  }
  if ($border) { $p.Line.Visible = -1; $p.Line.ForeColor.RGB = (C $border); $p.Line.Weight = 1 }
  return $p
}
function NOTES($sl, [string]$text) { $sl.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = $text }
# 본문 슬라이드 공통 머리: 작은 구분 라벨 + 제목 (모든 본문 슬라이드에서 같은 위치)
function HEAD($sl, [string]$kicker, [string]$title, [int]$no) {
  $null = TB $sl 48 30 864 18 $kicker 12 $PAL.red $true
  $null = TB $sl 48 50 864 44 $title 28 $PAL.ink $true
  $null = TB $sl 48 512 600 16 '해병대 전사(戰史) 디지털 연표' 10 $PAL.sub
  $null = TB $sl 812 512 100 16 "$no / 10" 10 $PAL.sub $false 3
}
function BULLETS($shape, [double]$after = 6) {
  $r = $shape.TextFrame.TextRange
  for ($i = 1; $i -le $r.Paragraphs().Count; $i++) {
    $p = $r.Paragraphs($i); $p.ParagraphFormat.Bullet.Visible = -1; $p.ParagraphFormat.Bullet.Character = 8226
    $p.ParagraphFormat.SpaceAfter = $after; $p.IndentLevel = 1
  }
  $shape.TextFrame.Ruler.Levels.Item(1).FirstMargin = 0; $shape.TextFrame.Ruler.Levels.Item(1).LeftMargin = 14
}
function BOLDLINE($shape, [int]$para, [string]$hex) { $p = $shape.TextFrame.TextRange.Paragraphs($para); $p.Font.Bold = -1; if ($hex) { $p.Font.Color.RGB = (C $hex) } }

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Add(0)
$pres.PageSetup.SlideWidth = 960; $pres.PageSetup.SlideHeight = 540
function NEW([int]$i) { $sl = $pres.Slides.Add($i, 12); $sl.FollowMasterBackground = 0; $sl.Background.Fill.Solid(); $sl.Background.Fill.ForeColor.RGB = (C $PAL.white); return $sl }

# ===== 1. 표지 =====
$s = NEW 1
$s.Background.Fill.ForeColor.RGB = (C $PAL.redD)
$null = PIC $s (Join-Path $photos 'INC-001.jpg') 500 0 460 540 $true
$null = TB $s 48 120 420 20 'LLM 코딩 에이전트 활용 업무 자동화 실습' 13 'D9B65C' $true
$null = TB $s 48 150 430 110 "해병대 전사(戰史)`r디지털 연표" 40 $PAL.white $true
$null = TB $s 48 275 420 60 '엑셀 한 장으로 만드는 오프라인 전투사 연표와 교육 슬라이드' 17 'F2E6E6'
$null = TB $s 48 430 420 18 '해병대 군사연구소 · 2026. 11.' 12 'D8C3C3'
$null = TB $s 48 452 420 18 'github.com/redala0012-beep/aaa' 12 'D8C3C3'
$null = TB $s 512 512 440 16 '사진: 인천상륙작전(1950.9.15), 미 해군역사유산사령부 NH 96876' 9 'EEEEEE' $false 3
NOTES $s '해병대 군사연구소가 디지털화하고 있는 전사 자료를 후배들이 쉽게 보도록, 엑셀만으로 연표 사이트와 교육 슬라이드를 만드는 도구를 LLM 코딩 에이전트(Claude Code)와 함께 하루 만에 만들었습니다.'

# ===== 2. 기획 의도 =====
$s = NEW 2
HEAD $s '01 기획 의도' '자료는 있는데, 후배들에게 닿지 않았다' 2
$stats = @(@('200여 점','디지털화한 전사 자료 (스캔 진행 중)'), @('2,000명','현장 교육 인원 (장병·교육생)'), @('10만 명+','개인 PC·휴대폰 열람 목표'))
for ($i = 0; $i -lt 3; $i++) {
  $x = 48 + $i * 296
  $null = BX $s 5 $x 118 272 112 $PAL.bg
  $null = TB $s ($x + 20) 136 232 44 $stats[$i][0] 32 $PAL.red $true
  $null = TB $s ($x + 20) 186 232 30 $stats[$i][1] 13 $PAL.sub
}
$null = BX $s 5 48 256 420 230 $PAL.white $PAL.line
$null = TB $s 68 274 380 22 '문제' 16 $PAL.ink $true
$t = TB $s 68 304 380 170 "스캔 자료가 흩어져 있고 형태가 딱딱해`v후배들이 찾아보지 않음`r6·25전쟁부터 연평도까지 전투가`v하나의 흐름으로 이어지지 않음`r연구관 3명이 웹 개발 없이`v자료를 계속 더할 방법이 없음" 14 $PAL.ink
BULLETS $t 8
$null = BX $s 5 492 256 420 230 $PAL.goldL
$null = TB $s 512 274 380 22 '목표' 16 $PAL.ink $true
$t = TB $s 512 304 380 170 "한 전투의 흐름(왜, 어떻게, 결과와 의의)을 3분 안에 이해`r같은 자료로 교육장 강의까지 바로 진행`r엑셀만으로 갱신 — 로그인·서버·인터넷 없이 동작" 14 $PAL.ink
BULLETS $t 8
NOTES $s '연구소장 인터뷰에서 나온 실제 문제입니다. 자료 200여 점은 있지만 전달이 안 되고, 목표 열람자는 현장 2천 명, 온라인 10만 명입니다. 그래서 목표를 3분 이해, 강의 겸용, 엑셀만으로 갱신으로 잡았습니다.'

# ===== 3. 사용자와 시나리오 =====
$s = NEW 3
HEAD $s '02 사용자와 핵심 시나리오' '한 번 입력하면, 읽기와 강의 두 곳에 쓰인다' 3
$cards = @(
  @('편', '연구관 3명', '작성자', "엑셀과 사진 폴더를 정리하고`r변환기에 끌어다 놓는다", '사이트(zip)가 자동으로 만들어짐'),
  @('읽', '장병 · 방문객', '열람자', "휴대폰으로 연표를 넘기다`r전투를 눌러 3분 동안 읽고,`r퀴즈를 푼다", '전투의 흐름과 의미'),
  @('강', '교관', '교육 진행', "같은 전투를 교육 모드로 띄워`r강의하고, 토의 뒤`r실제 결심을 공개한다", '별도 PPT 없이 바로 강의'))
for ($i = 0; $i -lt 3; $i++) {
  $x = 48 + $i * 296
  $null = BX $s 5 $x 118 272 300 $PAL.bg
  $o = BX $s 9 ($x + 20) 138 52 52 $PAL.red
  $null = TB $s ($x + 20) 138 52 52 $cards[$i][0] 20 $PAL.white $true 2 3
  $null = TB $s ($x + 86) 142 170 24 $cards[$i][1] 18 $PAL.ink $true
  $null = TB $s ($x + 86) 168 170 18 $cards[$i][2] 12 $PAL.sub
  $null = TB $s ($x + 20) 214 232 90 $cards[$i][3] 14 $PAL.ink
  $null = BX $s 5 ($x + 20) 330 232 68 $PAL.white
  $null = TB $s ($x + 34) 342 204 14 '얻는 것' 10 $PAL.red $true
  $null = TB $s ($x + 34) 360 204 30 $cards[$i][4] 13.5 $PAL.ink $true
}
$null = BX $s 5 48 438 864 50 $PAL.goldL
$null = TB $s 68 438 824 50 '데이터는 한 번만 입력 → 같은 내용이 휴대폰 읽기 화면과 교육장 전체화면 슬라이드에 함께 쓰임' 14 $PAL.ink $true 1 3
NOTES $s '사용자는 세 부류입니다. 연구관은 엑셀만 다루고, 열람자는 휴대폰으로 읽고, 교관은 같은 데이터를 교육 모드로 씁니다. 데이터를 한 번만 입력하게 한 것이 핵심 설계입니다.'

# ===== 4. MVP 설계 =====
$s = NEW 4
HEAD $s '03 MVP 설계' '기능은 다섯, 뺀 것은 분명하게' 4
$feat = @(
  @('F1','v0.1','전투 연표',"시간순 세로 연표 · 대표 사진`r전쟁 구분 필터"),
  @('F2','v0.1','상세 + 교육 모드',"같은 데이터로 읽기 화면과`r전체화면 슬라이드"),
  @('F3','v0.1','엑셀 → 사이트 변환기',"입력 검사 · 축소 · 워터마크`r비공개 제외 · zip"),
  @('F4','v0.2','대상자별 보기 + 퀴즈',"장병·간부·일반별 도입과 교훈`r퀴즈 16문항"),
  @('F5','v0.3','살아 있는 기록',"개정 이력 · 근거·신뢰 수준`r'지휘관이라면?' 결심"))
for ($i = 0; $i -lt 5; $i++) {
  $col = $i % 3; $row = [Math]::Floor($i / 3)
  $x = 48 + $col * 296; $y = 116 + $row * 150
  $null = BX $s 5 $x $y 272 132 $(if ($i -ge 3) { $PAL.goldL } else { $PAL.bg })
  $null = TB $s ($x + 18) ($y + 16) 60 18 $feat[$i][0] 14 $PAL.red $true
  $null = TB $s ($x + 194) ($y + 17) 60 16 $feat[$i][1] 11 $PAL.sub $false 3
  $null = TB $s ($x + 18) ($y + 42) 236 26 $feat[$i][2] 17 $PAL.ink $true
  $null = TB $s ($x + 18) ($y + 74) 236 46 $feat[$i][3] 12.5 $PAL.sub
}
$x = 640; $y = 266
$null = BX $s 5 $x $y 272 132 $PAL.white $PAL.line
$null = TB $s ($x + 18) ($y + 16) 236 18 '만들지 않은 것' 14 $PAL.ink $true
$t = TB $s ($x + 18) ($y + 44) 236 76 "로그인 · 결제 · 서버 · DB`r웹 편집 · 인터넷 게시(보류)" 13 $PAL.sub
$t.TextFrame2.TextRange.Font.Strike = 1
$null = TB $s 48 436 864 50 '엑셀 입력 방식을 고르자 로그인, 서버, DB가 모두 필요 없어졌습니다. 편집 권한은 엑셀 파일을 가진 연구관에게만 있습니다.' 14 $PAL.ink
NOTES $s 'v0.1에서 기능 3개로 시작해, 연구소장의 개선 요청에 따라 v0.2에서 대상자별 보기와 퀴즈, v0.3에서 살아 있는 기록을 더해 5개가 되었습니다. 로그인과 서버를 빼는 결정이 하루 개발을 가능하게 했습니다.'

# ===== 5. 개선 과정 =====
$s = NEW 5
HEAD $s '04 개선 과정' '읽는 자료에서, 참여하고 갱신되는 자료로' 5
$vers = @(
  @('v0.1','11.16','"바로 기능을 정하지 말고 하나씩 질문해 줘"','PRD → 연표 · 교육 모드 · 변환기 · 실제 사진 9장'),
  @('v0.2','11.17','"딱딱하지 않게, 대상자별로, 퀴즈로"','대상자 선택 · 퀴즈와 완주 배지 · 연평도 추모 톤'),
  @('v0.3','11.17','"일방향이고 변화를 담지 못한다"','개정 이력 · 근거·신뢰 수준 · 결심 시나리오'))
$null = BX $s 1 72 128 2 300 $PAL.line
for ($i = 0; $i -lt 3; $i++) {
  $y = 116 + $i * 122
  $null = BX $s 9 58 $y 30 30 $(if ($i -eq 2) { $PAL.red } else { $PAL.white }) $PAL.red
  $null = TB $s 58 $y 30 30 ([string]($i + 1)) 12 $(if ($i -eq 2) { $PAL.white } else { $PAL.red }) $true 2 3
  $null = TB $s 104 ($y + 2) 120 20 ($vers[$i][0] + '  ' + $vers[$i][1]) 15 $PAL.ink $true
  $null = TB $s 104 ($y + 28) 400 20 ('소장 요청  ' + $vers[$i][2]) 12.5 $PAL.red
  $null = TB $s 104 ($y + 52) 400 40 $vers[$i][3] 13.5 $PAL.ink
}
$null = PIC $s (Join-Path $shots '03b_decision_quiz.png') 540 116 372 360 $false $PAL.line
$null = TB $s 540 482 372 16 '결심 카드("지휘관이라면?")와 퀴즈가 경과 사이에 들어간 상세 화면' 10 $PAL.sub
NOTES $s '버전마다 연구소장의 한 문장이 출발점이었습니다. LLM은 바로 구현하지 않고 대상자 구분, 퀴즈 방식, 문체 범위 같은 결정 포인트를 권고안과 함께 되물었고, OK를 받은 뒤 PRD를 고치고 구현했습니다.'

# ===== 6. 시스템 구조 =====
$s = NEW 6
HEAD $s '05 시스템 구조' '외부 라이브러리 0개, 인터넷 연결 0회' 6
$flow = @(
  @('연구관','전사자료.xlsx`r+ 사진 폴더'),
  @('변환기 (브라우저)','XLSX 읽기 · 입력 검사`r축소 · 워터마크 · ZIP'),
  @('site.zip','index.html · data.js`r· images'),
  @('열람자 (오프라인)','대상자 선택 · 연표`r퀴즈 · 결심 · 교육 모드'))
for ($i = 0; $i -lt 4; $i++) {
  $x = 48 + $i * 222
  $fill = $(if ($i -eq 1) { $PAL.red } else { $PAL.bg })
  $fg = $(if ($i -eq 1) { $PAL.white } else { $PAL.ink })
  $null = BX $s 5 $x 122 186 120 $fill
  $null = TB $s ($x + 14) 140 158 22 $flow[$i][0] 15 $fg $true 2
  $null = TB $s ($x + 14) 172 158 56 ($flow[$i][1] -replace '`r', "`r") 12.5 $fg $false 2
  if ($i -lt 3) { $a = BX $s 33 ($x + 192) 168 24 26 $PAL.gold }
}
$pts = @(
  @('외부 라이브러리 0개',"엑셀(.xlsx)은 ZIP 속 XML`r브라우저 내장 기능으로 직접 읽고`r결과 ZIP도 직접 만듦"),
  @('더블클릭으로 실행',"file://에서는 JSON을 읽을 수 없음`r→ data.js를 <script>로 불러와`r더블클릭만으로 열림"),
  @('원본 보호',"결과물에는 1600px로 줄이고`r워터마크를 넣은 사진만 포함`r(원본 고해상도 미포함)"))
for ($i = 0; $i -lt 3; $i++) {
  $x = 48 + $i * 296
  $null = BX $s 5 $x 272 272 150 $PAL.white $PAL.line
  $null = TB $s ($x + 20) 290 232 22 $pts[$i][0] 16 $PAL.red $true
  $null = TB $s ($x + 20) 322 232 92 $pts[$i][1] 13 $PAL.ink
}
$null = TB $s 48 446 864 40 '군 PC는 프로그램 설치와 인터넷이 제한되는 경우가 많아, 브라우저 하나만으로 변환부터 열람까지 끝나도록 설계했습니다.' 14 $PAL.ink
NOTES $s '군 PC는 설치와 인터넷이 제한되므로, 엑셀 해석과 ZIP 생성을 직접 구현해 외부 라이브러리를 하나도 쓰지 않았습니다. 결과 사이트는 index.html 더블클릭으로 열립니다.'

# ===== 7. LLM 활용 과정 =====
$s = NEW 7
HEAD $s '06 LLM 활용 과정' '코딩보다 질문을 먼저 — 결정은 사람, 실행과 검증은 LLM' 7
$steps = @(
  @('요구사항 인터뷰','7개 항목에 답하고 결정',"하나씩 질문`r결정 포인트 D1~D14"),
  @('PRD','검토 후 "OK"',"v0.1 → v0.3 작성`r완료 기준 19개"),
  @('구현','"개발 시작"',"사이트 · 변환기`r빌드 스크립트"),
  @('검증','결과 확인',"브라우저 캡처`r화면 크기별 측정"),
  @('자료 수집','사진 9장 승인',"공개 라이선스`r사진 후보 조사"))
for ($i = 0; $i -lt 5; $i++) {
  $x = 48 + $i * 176
  $null = BX $s 9 ($x + 66) 112 36 36 $PAL.red
  $null = TB $s ($x + 66) 112 36 36 ([string]($i + 1)) 14 $PAL.white $true 2 3
  if ($i -lt 4) { $null = BX $s 1 ($x + 108) 129 128 2 $PAL.line }
  $null = TB $s $x 158 168 22 $steps[$i][0] 15 $PAL.ink $true 2
  $null = BX $s 5 ($x + 4) 188 160 66 $PAL.goldL
  $null = TB $s ($x + 14) 196 140 14 '사람' 10 '7A6320' $true
  $null = TB $s ($x + 14) 212 140 38 $steps[$i][1] 12 $PAL.ink
  $null = BX $s 5 ($x + 4) 262 160 82 $PAL.blueL
  $null = TB $s ($x + 14) 270 140 14 'LLM (Claude Code)' 10 $PAL.blue $true
  $null = TB $s ($x + 14) 286 140 54 $steps[$i][2] 12 $PAL.ink
}
$null = BX $s 5 48 366 864 120 $PAL.bg
$null = TB $s 72 380 200 18 '가장 효과가 컸던 프롬프트' 12 $PAL.red $true
$null = TB $s 72 402 820 80 "'지금부터 너는 해병대 군사연구소장으로서 역사 전문가야.`r나한테 바로 기능을 정하라고 하지 말고,`r배경 · 사용자 · 시나리오 · 데이터 · 화면 · 제약 · 완료 기준을 하나씩 질문해줘.`r내가 OK하면 PRD로 만들어줘.'" 13.5 $PAL.ink
NOTES $s '사람은 결정과 승인, LLM은 질문과 초안, 구현, 검증을 맡았습니다. 기능부터 정하지 않고 질문을 먼저 받은 덕분에, 엑셀 입력이면 로그인이 필요 없다는 결론이 자연스럽게 나왔습니다.'

# ===== 8. 디버깅과 문제 해결 =====
$s = NEW 8
HEAD $s '07 디버깅과 문제 해결' '"완료" 보고 뒤의 버그는 직접 재어 봐야 보인다' 8
$cases = @(
  @('사진 다운로드가 전부 429 차단','스크립트 요청을 봇 트래픽으로 판단','승인 범위 안에서 브라우저 경유로 받기'),
  @("세로 사진이 글자를`r화면 밖으로 밀어냄",'CSS Grid 자식의 기본 min-width:auto','min-width:0 · 대표 사진 4:3 고정'),
  @("'사진'으로 승인받은 자료가`r실제로는 홍보 전단",'메타데이터만 보고 후보를 고름','미리보기로 전수 확인 후 "문서"로 재분류'),
  @('검사 코드를 넣기만 하고 실행 안 함','정상 데이터로만 시험',"일부러 틀린 엑셀 3개로`r오류 26종 검출 확인"))
for ($i = 0; $i -lt 4; $i++) {
  $col = $i % 2; $row = [Math]::Floor($i / 2)
  $x = 48 + $col * 300; $y = 114 + $row * 190
  $null = BX $s 5 $x $y 284 174 $PAL.bg
  $t = TB $s ($x + 18) ($y + 16) 248 44 $cases[$i][0] 14.5 $PAL.ink $true
  $null = TB $s ($x + 18) ($y + 66) 248 14 '원인' 10 $PAL.sub $true
  $null = TB $s ($x + 18) ($y + 82) 248 32 $cases[$i][1] 12 $PAL.sub
  $null = TB $s ($x + 18) ($y + 118) 248 14 '해결' 10 $PAL.green $true
  $null = TB $s ($x + 18) ($y + 134) 248 32 $cases[$i][2] 12.5 $PAL.ink $true
}
$big = @(@('23건','문제 해결 기록 (README)'), @('26종','입력 오류 검사 모두 검출'), @('0장','화면 넘침 (3가지 화면, 66장)'))
for ($i = 0; $i -lt 3; $i++) {
  $y = 114 + $i * 124
  $null = BX $s 5 664 $y 248 112 $PAL.white $PAL.line
  $null = TB $s 684 ($y + 18) 208 44 $big[$i][0] 32 $PAL.red $true
  $null = TB $s 684 ($y + 70) 208 30 $big[$i][1] 12.5 $PAL.sub
}
NOTES $s 'LLM의 완료 보고만 믿지 않고 브라우저 캡처와 수치 측정으로 직접 확인했습니다. 홍보 전단 재분류처럼, 확인 과정에서 내용 오류도 찾았습니다. 전체 23건은 README 문제 해결 표에 있습니다.'

# ===== 9. 실행 결과 =====
$s = NEW 9
HEAD $s '08 실행 결과' '엑셀 1개 + 사진 9장 → 3MB 오프라인 사이트' 9
$grid = @(
  @('03_detail.png',"전투 상세 (간부 모드)`r이야기 상자`r신뢰 수준 표시"),
  @('04_edu_cover.png',"교육 모드 표지`r전체화면 슬라이드"),
  @('05_edu_decision.png',"교육 모드 결심`r토의 후`r실제 결심 공개"),
  @('10_converter_errors.jpg',"변환기 입력 검사`r행 번호와 함께`r오류 표시"))
for ($i = 0; $i -lt 4; $i++) {
  $col = $i % 2; $row = [Math]::Floor($i / 2)
  $x = 48 + $col * 440; $y = 108 + $row * 200
  $null = PIC $s (Join-Path $shots $grid[$i][0]) $x $y 296 170 $true $PAL.line
  $null = TB $s ($x + 306) ($y + 6) 118 160 $grid[$i][1] 12 $PAL.ink
}
$null = TB $s 48 492 864 16 '전투 4 · 퀴즈 16 · 결심 4 · 근거 8 · 개정 6 · 용어 9   |   원본 사진 32MB → 웹용 3MB   |   인터넷 없이 index.html로 실행' 11 $PAL.sub
NOTES $s '실제 실행 화면입니다. 원본 사진 32MB가 워터마크를 넣은 웹용 3MB로 줄어든 사이트가 만들어지고, 같은 데이터가 상세 화면과 교육 모드에 쓰입니다. 변환기는 잘못 입력한 행을 정확히 짚어 줍니다.'

# ===== 10. 한계와 다음 단계 =====
$s = NEW 10
$s.Background.Fill.ForeColor.RGB = (C $PAL.redD)
$null = TB $s 48 30 864 18 '09 한계와 다음 단계' 12 'D9B65C' $true
$null = TB $s 48 50 864 44 '도구는 완성, 이제 내용을 채울 차례' 28 $PAL.white $true
$null = TB $s 48 124 400 22 '한계' 16 'D9B65C' $true
$t = TB $s 48 154 400 200 "전투 문구·퀴즈·근거는 LLM 초안 — 연구소장 검수 필요`r도솔산·짜빈동은 전투 당시 사진이 없어 같은 시기 사진으로 대체`r제보·증언 창구는 운영 정책이 정해진 뒤 구현" 14 'F2E6E6'
BULLETS $t 10
$null = TB $s 500 124 412 22 '다음 단계' 16 'D9B65C' $true
$t = TB $s 500 154 412 200 "연구소 소장 스캔본 200여 점 반영, 전투 확대`r검수 완료 후 인터넷 게시 (Vercel 등)`r전시물 QR코드 → ?aud= 주소로 대상자별 바로 열기`r제보·증언 창구 → 열람자와 연구소의 순환 구조" 14 'F2E6E6'
BULLETS $t 10
$null = BX $s 5 48 400 864 88 'FFFFFF'
$null = TB $s 72 416 400 16 'GitHub 저장소 (소스 코드 · 샘플 데이터 · README)' 11 $PAL.red $true
$null = TB $s 72 438 820 36 'github.com/redala0012-beep/aaa' 24 $PAL.ink $true
NOTES $s '도구는 완성되었고, 남은 일은 내용 검수와 확대입니다. 검수가 끝나면 인터넷에 게시하고, 대상자별 QR코드와 제보·증언 창구로 열람자와 연구소가 이어지는 구조를 만들 계획입니다.'

# ---- 저장 ----
if (Test-Path $outPptx) { Remove-Item $outPptx }
if (Test-Path $outPdf) { Remove-Item $outPdf }
$pres.SaveAs($outPptx, 24)
$pres.SaveAs($outPdf, 32)
if ($QaDir) {
  New-Item -ItemType Directory -Force $QaDir | Out-Null
  for ($i = 1; $i -le $pres.Slides.Count; $i++) { $pres.Slides.Item($i).Export((Join-Path $QaDir ("slide-{0:D2}.png" -f $i)), 'PNG', 1600, 900) }
}
$n = $pres.Slides.Count
$pres.Close(); $ppt.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
"slides: $n"
"pptx: {0:N0}KB / pdf: {1:N0}KB" -f ((Get-Item $outPptx).Length/1KB), ((Get-Item $outPdf).Length/1KB)
