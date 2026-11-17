# 발표자료(docs/발표자료.pptx, .pdf)와 갤러리 표지(docs/갤러리_표지.png)를 PowerPoint 자동화로 생성합니다.
# 사용: powershell -ExecutionPolicy Bypass -File tools\make_slides.ps1 [-QaDir 폴더]
# 필요: Microsoft PowerPoint (설치만 되어 있으면 됨). 이 파일은 UTF-8(BOM)으로 저장해야 한글이 깨지지 않습니다.
param([string]$QaDir = '')
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$shots = Join-Path $root 'docs\screenshots'
$photos = Join-Path $root 'site\images\full'
$outPptx = Join-Path $root 'docs\발표자료.pptx'
$outPdf = Join-Path $root 'docs\발표자료.pdf'
$outCover = Join-Path $root 'docs\갤러리_표지.png'
$DECK = '알기쉽게 배우는 해병대 디지털 전사'
$SITE = 'rokmc-war-history-timeline.vercel.app'
$TOTAL = 11

# ---- 색: SKKU 바이브코딩 갤러리(skku-vibecoding-gallery.vercel.app)와 같은 계열 ----
#  바탕 #0D1D16 · 카드 #183128/#12261D · 강조 연두 #8DC63F · 글자 흰색/흰색 70%
$PAL = @{
  bg='0D1D16'; card='183128'; card2='12261D'; tint='22402F'; line='2E4D3F'
  txt='F2F5F3'; sub='A9B5AE'; acc='8DC63F'; accD='0D1D16'; llm='1C3640'; llmT='8FC9D9'
}
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
  if ($type -eq 5) { $s.Adjustments.Item(1) = 0.08 }
  if ($s.HasTextFrame) { $s.TextFrame.TextRange.Text = '' }
  return $s
}
function PIC($sl, [string]$path, [double]$x, [double]$y, [double]$w, [double]$h, [bool]$cover = $false, [string]$border = '', [bool]$round = $false) {
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
  if ($round) { $p.AutoShapeType = 5; $p.Adjustments.Item(1) = 0.05; $p.LockAspectRatio = 0; $p.Height = $h; $p.Width = $w; $p.Left = $x; $p.Top = $y }
  if ($border) { $p.Line.Visible = -1; $p.Line.ForeColor.RGB = (C $border); $p.Line.Weight = 1.25 }
  return $p
}
function CHIP($sl, [double]$x, [double]$y, [double]$w, [string]$text, [double]$size = 11.5) {
  $b = BX $sl 5 $x $y $w 26 $PAL.tint
  $b.Adjustments.Item(1) = 0.5
  $null = TB $sl $x $y $w 26 $text $size $PAL.acc $true 2 3
}
function NOTES($sl, [string]$text) { $sl.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = $text }
# 본문 슬라이드 공통 머리: 연두 점 + 구분 라벨 + 제목 (갤러리 사이트의 섹션 머리와 같은 형식)
function HEAD($sl, [string]$kicker, [string]$title, [int]$no) {
  $null = BX $sl 9 48 35 7 7 $PAL.acc
  $null = TB $sl 62 30 850 18 $kicker 12 $PAL.acc $true
  $null = TB $sl 48 50 864 44 $title 27 $PAL.txt $true
  $null = BX $sl 1 48 505 864 0.75 $PAL.line
  $null = TB $sl 48 514 600 16 $DECK 10 $PAL.sub
  $null = TB $sl 812 514 100 16 "$no / $TOTAL" 10 $PAL.sub $false 3
}
function BULLETS($shape, [double]$after = 6) {
  $r = $shape.TextFrame.TextRange
  for ($i = 1; $i -le $r.Paragraphs().Count; $i++) {
    $p = $r.Paragraphs($i); $p.ParagraphFormat.Bullet.Visible = -1; $p.ParagraphFormat.Bullet.Character = 8226
    $p.ParagraphFormat.Bullet.Font.Color.RGB = (C $PAL.acc)
    $p.ParagraphFormat.SpaceAfter = $after; $p.IndentLevel = 1
  }
  $shape.TextFrame.Ruler.Levels.Item(1).FirstMargin = 0; $shape.TextFrame.Ruler.Levels.Item(1).LeftMargin = 14
}

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Add(0)
$pres.PageSetup.SlideWidth = 960; $pres.PageSetup.SlideHeight = 540
function NEW([int]$i) { $sl = $pres.Slides.Add($i, 12); $sl.FollowMasterBackground = 0; $sl.Background.Fill.Solid(); $sl.Background.Fill.ForeColor.RGB = (C $PAL.bg); return $sl }

# ===== 1. 표지 (갤러리 썸네일 겸용: 핵심 내용은 가운데 16:10 영역 안에) =====
$s = NEW 1
$null = PIC $s (Join-Path $photos 'INC-001.jpg') 520 64 392 290 $true $PAL.line $true
$null = PIC $s (Join-Path $shots '02_timeline.png') 476 278 236 176 $true $PAL.acc $true
$null = TB $s 724 362 188 16 '인천상륙작전(1950) · 연표 화면' 10 $PAL.sub $false 3
$null = TB $s 48 124 440 52 '알기쉽게 배우는' 38 $PAL.txt $true
$null = TB $s 48 176 460 60 '해병대 디지털 전사' 44 $PAL.acc $true
$null = TB $s 48 252 410 60 "엑셀 한 장으로 만드는 전투사 연표 · 퀴즈 · 교육 슬라이드,`r그리고 보는 사람 눈높이의 AI 3줄 설명" 14.5 $PAL.sub
$chips = @(@('대상자별 보기', 104), @('퀴즈 · 결심', 90), @('교육 모드', 80), @('AI 3줄 설명', 96))
$cx = 48
foreach ($c in $chips) { CHIP $s $cx 334 $c[1] $c[0]; $cx += $c[1] + 8 }
$null = TB $s 48 420 400 18 '해병대 군사연구소 · 2026. 11.' 12 $PAL.sub
$null = TB $s 48 442 420 22 $SITE 15 $PAL.acc $true
NOTES $s '"알기쉽게 배우는 해병대 디지털 전사"입니다. 해병대 군사연구소가 디지털화하고 있는 전사 자료를 장병과 방문객이 쉽게 보도록, 엑셀만으로 연표 사이트와 교육 슬라이드를 만들고 AI가 눈높이에 맞춰 풀어 주는 도구를 LLM 코딩 에이전트(Claude Code)와 함께 만들었습니다.'

# ===== 2. 한눈에 보기 (갤러리 소개 카드 형식) =====
$s = NEW 2
HEAD $s '00 한눈에 보기' '엑셀로 쌓고, 연표로 읽고, 교육장에서 바로 쓴다' 2
$null = BX $s 5 48 108 520 384 $PAL.card $PAL.line
$null = TB $s 70 126 480 16 '한 줄 소개' 11 $PAL.acc $true
$null = TB $s 70 146 480 46 "6·25전쟁부터 연평도까지, 해병대 전투사를`r보는 사람 눈높이에 맞춰 3분 만에 이해하는 디지털 연표" 15 $PAL.txt $true
$null = TB $s 70 206 480 16 '소개' 11 $PAL.acc $true
$null = TB $s 70 226 480 92 "연구관이 엑셀과 사진 폴더만 정리하면 변환기가 워터마크를 넣은 오프라인 사이트를 만들어 줍니다. 장병 · 간부 · 방문객별 이야기와 퀴즈, '지휘관이라면?' 결심, 근거와 개정 이력을 담았고, 같은 자료가 교육장 슬라이드로도 쓰입니다." 12.5 $PAL.sub
$null = TB $s 70 330 480 16 '주요 특징' 11 $PAL.acc $true
$f1 = @(@('대상자별 보기', 104), @('퀴즈 16문항', 96), @('결심 시나리오', 104), @('교육 모드', 80))
$f2 = @(@('근거 · 신뢰 수준', 118), @('개정 이력', 80), @('AI 3줄 설명 (Gemini)', 150))
$cx = 70; foreach ($c in $f1) { CHIP $s $cx 352 $c[1] $c[0] 11; $cx += $c[1] + 8 }
$cx = 70; foreach ($c in $f2) { CHIP $s $cx 386 $c[1] $c[0] 11; $cx += $c[1] + 8 }
$null = TB $s 70 426 480 50 "하이라이트 · 외부 라이브러리 0개, 더블클릭만으로 인터넷 없이 동작하고,`r온라인판에서는 AI가 전투 본문만 근거로 눈높이 설명을 만듭니다." 11.5 $PAL.txt
$null = PIC $s (Join-Path $shots '01_who.png') 592 108 320 214 $true $PAL.line $true
$null = BX $s 5 592 338 320 72 $PAL.card $PAL.line
$null = TB $s 610 350 284 14 '사이트 열기 ↗' 11 $PAL.acc $true
$null = TB $s 610 370 284 30 $SITE 12.5 $PAL.txt $true
$null = BX $s 5 592 420 320 72 $PAL.card $PAL.line
$null = TB $s 610 432 284 14 'GitHub (소스 · 샘플 데이터 · README)' 11 $PAL.acc $true
$null = TB $s 610 452 284 30 'github.com/redala0012-beep/aaa' 12.5 $PAL.txt $true
NOTES $s '갤러리 소개 카드와 같은 형식으로 정리했습니다. 엑셀로 쌓고, 연표로 읽고, 교육장에서 바로 쓰는 도구입니다. 온라인판에서는 Gemini가 대상자 눈높이의 3줄 설명을 만들어 줍니다.'

# ===== 3. 기획 의도 =====
$s = NEW 3
HEAD $s '01 기획 의도' '자료는 있는데, 후배들에게 닿지 않았다' 3
$stats = @(@('200여 점','디지털화한 전사 자료 (스캔 진행 중)'), @('2,000명','현장 교육 인원 (장병·교육생)'), @('10만 명+','개인 PC·휴대폰 열람 목표'))
for ($i = 0; $i -lt 3; $i++) {
  $x = 48 + $i * 296
  $null = BX $s 5 $x 118 272 112 $PAL.card
  $null = TB $s ($x + 20) 136 232 44 $stats[$i][0] 32 $PAL.acc $true
  $null = TB $s ($x + 20) 186 232 30 $stats[$i][1] 13 $PAL.sub
}
$null = BX $s 5 48 256 420 230 $PAL.card2 $PAL.line
$null = TB $s 68 274 380 22 '문제' 16 $PAL.txt $true
$t = TB $s 68 304 380 170 "스캔 자료가 흩어져 있고 형태가 딱딱해`v후배들이 찾아보지 않음`r6·25전쟁부터 연평도까지 전투가`v하나의 흐름으로 이어지지 않음`r연구관 3명이 웹 개발 없이`v자료를 계속 더할 방법이 없음" 14 $PAL.txt
BULLETS $t 8
$null = BX $s 5 492 256 420 230 $PAL.tint
$null = TB $s 512 274 380 22 '목표' 16 $PAL.acc $true
$t = TB $s 512 304 380 170 "한 전투의 흐름(왜, 어떻게, 결과와 의의)을 3분 안에 이해`r같은 자료로 교육장 강의까지 바로 진행`r엑셀만으로 갱신 — 로그인·DB 없이, 오프라인에서도 동작" 14 $PAL.txt
BULLETS $t 8
NOTES $s '연구소장 인터뷰에서 나온 실제 문제입니다. 자료 200여 점은 있지만 전달이 안 되고, 목표 열람자는 현장 2천 명, 온라인 10만 명입니다. 그래서 목표를 3분 이해, 강의 겸용, 엑셀만으로 갱신으로 잡았습니다.'

# ===== 4. 사용자와 시나리오 =====
$s = NEW 4
HEAD $s '02 사용자와 핵심 시나리오' '한 번 입력하면, 읽기와 강의 두 곳에 쓰인다' 4
$cards = @(
  @('편', '연구관 3명', '작성자', "엑셀과 사진 폴더를 정리하고`r변환기에 끌어다 놓는다", '사이트(zip)가 자동으로 만들어짐'),
  @('읽', '장병 · 방문객', '열람자', "휴대폰으로 연표를 넘기다`r전투를 눌러 3분 동안 읽고,`r퀴즈를 푼다", '전투의 흐름과 의미'),
  @('강', '교관', '교육 진행', "같은 전투를 교육 모드로 띄워`r강의하고, 토의 뒤`r실제 결심을 공개한다", '별도 PPT 없이 바로 강의'))
for ($i = 0; $i -lt 3; $i++) {
  $x = 48 + $i * 296
  $null = BX $s 5 $x 118 272 300 $PAL.card
  $null = BX $s 9 ($x + 20) 138 52 52 $PAL.acc
  $null = TB $s ($x + 20) 138 52 52 $cards[$i][0] 20 $PAL.accD $true 2 3
  $null = TB $s ($x + 86) 142 170 24 $cards[$i][1] 18 $PAL.txt $true
  $null = TB $s ($x + 86) 168 170 18 $cards[$i][2] 12 $PAL.sub
  $null = TB $s ($x + 20) 214 232 90 $cards[$i][3] 14 $PAL.txt
  $null = BX $s 5 ($x + 20) 330 232 68 $PAL.card2
  $null = TB $s ($x + 34) 342 204 14 '얻는 것' 10 $PAL.acc $true
  $null = TB $s ($x + 34) 360 204 30 $cards[$i][4] 13.5 $PAL.txt $true
}
$null = BX $s 5 48 438 864 50 $PAL.tint
$null = TB $s 68 438 824 50 '데이터는 한 번만 입력 → 같은 내용이 휴대폰 읽기 화면과 교육장 전체화면 슬라이드에 함께 쓰임' 14 $PAL.txt $true 1 3
NOTES $s '사용자는 세 부류입니다. 연구관은 엑셀만 다루고, 열람자는 휴대폰으로 읽고, 교관은 같은 데이터를 교육 모드로 씁니다. 데이터를 한 번만 입력하게 한 것이 핵심 설계입니다.'

# ===== 5. MVP 설계 =====
$s = NEW 5
HEAD $s '03 MVP 설계' '기능은 여섯, 뺀 것은 분명하게' 5
$feat = @(
  @('F1','v0.1','전투 연표',"시간순 세로 연표 · 대표 사진`r전쟁 구분 필터"),
  @('F2','v0.1','상세 + 교육 모드',"같은 데이터로 읽기 화면과`r전체화면 슬라이드"),
  @('F3','v0.1','엑셀 → 사이트 변환기',"입력 검사 · 축소 · 워터마크`r비공개 제외 · zip"),
  @('F4','v0.2','대상자별 보기 + 퀴즈',"장병·간부·일반별 도입과 교훈`r퀴즈 16문항"),
  @('F5','v0.3','살아 있는 기록',"개정 이력 · 근거·신뢰 수준`r'지휘관이라면?' 결심"),
  @('F6','v0.4','AI 쉬운 설명',"Gemini가 본문만 근거로`r대상자 눈높이 3줄 설명"))
for ($i = 0; $i -lt 6; $i++) {
  $col = $i % 3; $row = [Math]::Floor($i / 3)
  $x = 48 + $col * 296; $y = 116 + $row * 150
  $null = BX $s 5 $x $y 272 132 $(if ($i -ge 3) { $PAL.tint } else { $PAL.card })
  $null = TB $s ($x + 18) ($y + 16) 60 18 $feat[$i][0] 14 $PAL.acc $true
  $null = TB $s ($x + 194) ($y + 17) 60 16 $feat[$i][1] 11 $PAL.sub $false 3
  $null = TB $s ($x + 18) ($y + 42) 236 26 $feat[$i][2] 17 $PAL.txt $true
  $null = TB $s ($x + 18) ($y + 74) 236 46 $feat[$i][3] 12.5 $PAL.sub
}
$null = TB $s 48 432 864 18 '만들지 않은 것' 12 $PAL.acc $true
$null = TB $s 48 452 864 40 '로그인 · 결제 · DB · 웹 편집 — 엑셀 입력을 고르자 모두 필요 없어졌습니다. 편집 권한은 엑셀 파일을 가진 연구관에게만 있고, 서버는 AI 호출 하나에만 씁니다.' 13.5 $PAL.txt
NOTES $s 'v0.1에서 기능 3개로 시작해 v0.2 대상자별 보기와 퀴즈, v0.3 살아 있는 기록, v0.4 AI 쉬운 설명을 더해 6개가 되었습니다. 로그인과 DB를 빼는 결정이 짧은 개발을 가능하게 했고, 서버는 API 키를 숨기기 위한 AI 호출 하나에만 씁니다.'

# ===== 6. 개선 과정 =====
$s = NEW 6
HEAD $s '04 개선 과정' '읽는 자료에서, 참여하고 갱신되는 자료로' 6
$vers = @(
  @('v0.1','11.16','"바로 기능을 정하지 말고 하나씩 질문해 줘"','PRD → 연표 · 교육 모드 · 변환기 · 실제 사진 9장'),
  @('v0.2','11.17','"딱딱하지 않게, 대상자별로, 퀴즈로"','대상자 선택 · 퀴즈와 완주 배지 · 연평도 추모 톤'),
  @('v0.3','11.17','"일방향이고 변화를 담지 못한다"','개정 이력 · 근거·신뢰 수준 · 결심 시나리오'),
  @('v0.4','11.17','"AI 기능 하나 붙여줘, 호출은 서버에서만"','대상자별 AI 3줄 설명 · Vercel 배포'))
$null = BX $s 1 72 128 2 290 $PAL.line
for ($i = 0; $i -lt 4; $i++) {
  $y = 112 + $i * 96
  $last = ($i -eq 3)
  $null = BX $s 9 58 $y 30 30 $(if ($last) { $PAL.acc } else { $PAL.bg }) $PAL.acc
  $null = TB $s 58 $y 30 30 ([string]($i + 1)) 12 $(if ($last) { $PAL.accD } else { $PAL.acc }) $true 2 3
  $null = TB $s 104 ($y + 2) 120 20 ($vers[$i][0] + '  ' + $vers[$i][1]) 15 $PAL.txt $true
  $null = TB $s 104 ($y + 26) 410 20 ('소장 요청  ' + $vers[$i][2]) 12.5 $PAL.acc
  $null = TB $s 104 ($y + 48) 410 36 $vers[$i][3] 13.5 $PAL.txt
}
$null = PIC $s (Join-Path $shots '03b_decision_quiz.png') 540 112 372 364 $false $PAL.line
$null = TB $s 540 482 372 16 '결심 카드("지휘관이라면?")와 퀴즈가 경과 사이에 들어간 상세 화면' 10 $PAL.sub
NOTES $s '버전마다 연구소장의 한 문장이 출발점이었습니다. LLM은 바로 구현하지 않고 대상자 구분, 퀴즈 방식, AI 기능의 종류 같은 결정 포인트를 권고안과 함께 되물었고, OK를 받은 뒤 구현했습니다.'

# ===== 7. 시스템 구조 =====
$s = NEW 7
HEAD $s '05 시스템 구조' '외부 라이브러리 0개, 오프라인에서도 그대로 동작' 7
$flow = @(
  @('연구관','전사자료.xlsx`r+ 사진 폴더'),
  @('변환기 (브라우저)','XLSX 읽기 · 입력 검사`r축소 · 워터마크 · ZIP'),
  @('site.zip','index.html · data.js`r· images'),
  @('열람자','대상자 선택 · 연표`r퀴즈 · 결심 · 교육 모드'))
for ($i = 0; $i -lt 4; $i++) {
  $x = 48 + $i * 222
  $hi = ($i -eq 1)
  $null = BX $s 5 $x 116 186 112 $(if ($hi) { $PAL.acc } else { $PAL.card })
  $fg = $(if ($hi) { $PAL.accD } else { $PAL.txt })
  $null = TB $s ($x + 14) 132 158 22 $flow[$i][0] 15 $fg $true 2
  $null = TB $s ($x + 14) 162 158 56 ($flow[$i][1] -replace '`r', "`r") 12.5 $fg $false 2
  if ($i -lt 3) { $null = BX $s 33 ($x + 192) 160 24 26 $PAL.acc }
}
$pts = @(
  @('외부 라이브러리 0개',"엑셀(.xlsx)은 ZIP 속 XML`r브라우저 내장 기능으로 직접 읽고`r결과 ZIP도 직접 만듦"),
  @('더블클릭으로 실행',"file://에서는 JSON을 읽을 수 없음`r→ data.js를 <script>로 불러와`r인터넷 없이 열림"),
  @('AI는 서버에서만',"온라인판의 api/explain.js가`r본문만 근거로 Gemini 호출`r키는 서버 환경변수에만 보관"))
for ($i = 0; $i -lt 3; $i++) {
  $x = 48 + $i * 296
  $null = BX $s 5 $x 252 272 160 $PAL.card2 $PAL.line
  $null = TB $s ($x + 20) 270 232 22 $pts[$i][0] 16 $PAL.acc $true
  $null = TB $s ($x + 20) 302 232 100 $pts[$i][1] 13 $PAL.txt
}
$null = TB $s 48 432 864 56 '군 PC는 프로그램 설치와 인터넷이 제한되는 경우가 많아, 브라우저 하나만으로 변환부터 열람까지 끝나도록 설계했습니다. 인터넷이 되는 곳(Vercel 배포판)에서만 AI 버튼이 나타납니다.' 14 $PAL.txt
NOTES $s '군 PC는 설치와 인터넷이 제한되므로 엑셀 해석과 ZIP 생성을 직접 구현해 외부 라이브러리를 하나도 쓰지 않았습니다. AI 기능은 온라인판에서만 쓰이며, 브라우저는 전투 ID와 대상자만 보내고 키는 서버에만 있습니다.'

# ===== 8. LLM 활용 과정 =====
$s = NEW 8
HEAD $s '06 LLM 활용 과정' '코딩보다 질문을 먼저 — 결정은 사람, 실행과 검증은 LLM' 8
$steps = @(
  @('요구사항 인터뷰','7개 항목에 답하고 결정',"하나씩 질문`r결정 포인트 D1~D14"),
  @('PRD','검토 후 "OK"',"v0.1 → v0.3 작성`r완료 기준 19개"),
  @('구현','"개발 시작"',"사이트 · 변환기`r빌드 · AI 서버 함수"),
  @('검증','결과 확인',"브라우저 캡처`r화면 크기별 측정"),
  @('배포','키 입력 · 로그인 승인',"Vercel 배포`r키는 화면에 미출력"))
for ($i = 0; $i -lt 5; $i++) {
  $x = 48 + $i * 176
  $null = BX $s 9 ($x + 66) 112 36 36 $PAL.acc
  $null = TB $s ($x + 66) 112 36 36 ([string]($i + 1)) 14 $PAL.accD $true 2 3
  if ($i -lt 4) { $null = BX $s 1 ($x + 108) 129 128 2 $PAL.line }
  $null = TB $s $x 158 168 22 $steps[$i][0] 15 $PAL.txt $true 2
  $null = BX $s 5 ($x + 4) 188 160 66 $PAL.tint
  $null = TB $s ($x + 14) 196 140 14 '사람' 10 $PAL.acc $true
  $null = TB $s ($x + 14) 212 140 38 $steps[$i][1] 12 $PAL.txt
  $null = BX $s 5 ($x + 4) 262 160 82 $PAL.llm
  $null = TB $s ($x + 14) 270 140 14 'LLM (Claude Code)' 10 $PAL.llmT $true
  $null = TB $s ($x + 14) 286 140 54 $steps[$i][2] 12 $PAL.txt
}
$null = BX $s 5 48 366 864 122 $PAL.card
$null = TB $s 72 380 200 18 '가장 효과가 컸던 프롬프트' 12 $PAL.acc $true
$null = TB $s 72 402 820 80 "'지금부터 너는 해병대 군사연구소장으로서 역사 전문가야.`r나한테 바로 기능을 정하라고 하지 말고,`r배경 · 사용자 · 시나리오 · 데이터 · 화면 · 제약 · 완료 기준을 하나씩 질문해줘.`r내가 OK하면 PRD로 만들어줘.'" 13.5 $PAL.txt
NOTES $s '사람은 결정과 승인, LLM은 질문과 초안, 구현, 검증, 배포를 맡았습니다. 기능부터 정하지 않고 질문을 먼저 받은 덕분에, 엑셀 입력이면 로그인이 필요 없다는 결론이 자연스럽게 나왔습니다. 비밀 키는 사람이 직접 넣고 LLM은 값을 출력하지 않았습니다.'

# ===== 9. 디버깅과 문제 해결 =====
$s = NEW 9
HEAD $s '07 디버깅과 문제 해결' '"완료" 보고 뒤의 버그는 직접 재어 봐야 보인다' 9
$cases = @(
  @('사진 다운로드가 전부 429 차단','스크립트 요청을 봇 트래픽으로 판단','승인 범위 안에서 브라우저 경유로 받기'),
  @("세로 사진이 글자를`r화면 밖으로 밀어냄",'CSS Grid 자식의 기본 min-width:auto','min-width:0 · 대표 사진 4:3 고정'),
  @("'사진'으로 승인받은 자료가`r실제로는 홍보 전단",'메타데이터만 보고 후보를 고름','미리보기로 전수 확인 후 "문서"로 재분류'),
  @("배포판에서만 AI 설명이`r0.4초 만에 실패",'키 값 앞에 보이지 않는 BOM 문자','오류 로그로 원인 확인 · 서버에서 제거'))
for ($i = 0; $i -lt 4; $i++) {
  $col = $i % 2; $row = [Math]::Floor($i / 2)
  $x = 48 + $col * 300; $y = 114 + $row * 190
  $null = BX $s 5 $x $y 284 174 $PAL.card
  $null = TB $s ($x + 18) ($y + 16) 248 44 $cases[$i][0] 14.5 $PAL.txt $true
  $null = TB $s ($x + 18) ($y + 66) 248 14 '원인' 10 $PAL.sub $true
  $null = TB $s ($x + 18) ($y + 82) 248 32 $cases[$i][1] 12 $PAL.sub
  $null = TB $s ($x + 18) ($y + 118) 248 14 '해결' 10 $PAL.acc $true
  $null = TB $s ($x + 18) ($y + 134) 248 32 $cases[$i][2] 12.5 $PAL.txt $true
}
$big = @(@('25건','문제 해결 기록 (README)'), @('26종','입력 오류 검사 모두 검출'), @('0장','화면 넘침 (3가지 화면, 66장)'))
for ($i = 0; $i -lt 3; $i++) {
  $y = 114 + $i * 124
  $null = BX $s 5 664 $y 248 112 $PAL.card2 $PAL.line
  $null = TB $s 684 ($y + 18) 208 44 $big[$i][0] 32 $PAL.acc $true
  $null = TB $s 684 ($y + 70) 208 30 $big[$i][1] 12.5 $PAL.sub
}
NOTES $s 'LLM의 완료 보고만 믿지 않고 브라우저 캡처, 수치 측정, 서버 로그로 직접 확인했습니다. 홍보 전단 재분류처럼 내용 오류도 찾았고, 배포판에서만 생긴 AI 오류는 로그의 문자 코드 하나로 원인을 찾았습니다. 전체 25건은 README 문제 해결 표에 있습니다.'

# ===== 10. 실행 결과 =====
$s = NEW 10
HEAD $s '08 실행 결과' '엑셀 1개 + 사진 9장 → 3MB 사이트, 온라인에서는 AI까지' 10
$grid = @(
  @('03_detail.png',"전투 상세 (간부 모드)`r이야기 상자`r신뢰 수준 표시"),
  @('04_edu_cover.png',"교육 모드 표지`r전체화면 슬라이드"),
  @('05_edu_decision.png',"교육 모드 결심`r토의 후`r실제 결심 공개"),
  @('10_converter_errors.jpg',"변환기 입력 검사`r행 번호와 함께`r오류 표시"))
for ($i = 0; $i -lt 4; $i++) {
  $col = $i % 2; $row = [Math]::Floor($i / 2)
  $x = 48 + $col * 440; $y = 108 + $row * 196
  $null = PIC $s (Join-Path $shots $grid[$i][0]) $x $y 296 168 $true $PAL.line
  $null = TB $s ($x + 306) ($y + 6) 118 156 $grid[$i][1] 12 $PAL.txt
}
$null = TB $s 48 482 864 16 "전투 4 · 퀴즈 16 · 결심 4 · 근거 8 · 개정 6 · 용어 9   |   원본 사진 32MB → 웹용 3MB   |   $SITE" 11 $PAL.sub
NOTES $s '실제 실행 화면입니다. 원본 사진 32MB가 워터마크를 넣은 웹용 3MB로 줄어든 사이트가 만들어지고, 같은 데이터가 상세 화면과 교육 모드에 쓰입니다. 변환기는 잘못 입력한 행을 정확히 짚어 줍니다. 온라인판 주소에서 AI 3줄 설명도 직접 눌러 볼 수 있습니다.'

# ===== 11. 한계와 다음 단계 =====
$s = NEW 11
HEAD $s '09 한계와 다음 단계' '도구는 완성, 이제 내용을 채울 차례' 11
$null = BX $s 5 48 108 420 270 $PAL.card2 $PAL.line
$null = TB $s 68 124 380 22 '한계' 16 $PAL.acc $true
$t = TB $s 68 156 380 210 "전투 문구·퀴즈·근거는 LLM 초안 — 연구소장 검수 필요`rAI 설명은 요약일 뿐 — 공식 서술은 본문과 근거 기준`r도솔산·짜빈동은 같은 시기 사진으로 대체`r무료 AI 키는 분당 요청 수 제한" 13.5 $PAL.txt
BULLETS $t 10
$null = BX $s 5 492 108 420 270 $PAL.tint
$null = TB $s 512 124 380 22 '다음 단계' 16 $PAL.acc $true
$t = TB $s 512 156 380 210 "연구소 소장 스캔본 200여 점 반영, 전투 확대`r검수 완료 후 온라인판 공개 범위 확정`r전시물 QR코드 → ?aud= 주소로 대상자별 바로 열기`r제보·증언 창구 → 열람자와 연구소의 순환 구조" 13.5 $PAL.txt
BULLETS $t 10
$null = BX $s 5 48 396 864 96 $PAL.acc
$null = TB $s 72 410 400 16 '직접 눌러 보기' 11 $PAL.accD $true
$null = TB $s 72 430 820 30 $SITE 22 $PAL.accD $true
$null = TB $s 72 464 820 18 'GitHub  github.com/redala0012-beep/aaa' 12 $PAL.accD
NOTES $s '도구는 완성되었고, 남은 일은 내용 검수와 확대입니다. 검수가 끝나면 공개 범위를 정하고, 대상자별 QR코드와 제보·증언 창구로 열람자와 연구소가 이어지는 구조를 만들 계획입니다.'

# ---- 저장 ----
if (Test-Path $outPptx) { Remove-Item $outPptx }
if (Test-Path $outPdf) { Remove-Item $outPdf }
$pres.SaveAs($outPptx, 24)
$pres.SaveAs($outPdf, 32)
$pres.Slides.Item(1).Export($outCover, 'PNG', 1920, 1080)
if ($QaDir) {
  New-Item -ItemType Directory -Force $QaDir | Out-Null
  for ($i = 1; $i -le $pres.Slides.Count; $i++) { $pres.Slides.Item($i).Export((Join-Path $QaDir ("slide-{0:D2}.png" -f $i)), 'PNG', 1600, 900) }
}
$n = $pres.Slides.Count
$pres.Close(); $ppt.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
"slides: $n"
"pptx: {0:N0}KB / pdf: {1:N0}KB / cover: {2:N0}KB" -f ((Get-Item $outPptx).Length/1KB), ((Get-Item $outPdf).Length/1KB), ((Get-Item $outCover).Length/1KB)
