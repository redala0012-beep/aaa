# Build helper (ASCII only; Korean text lives in content.json / site-template.html)
#   -Template      : converter/site-template.js  from src/site-template.html
#   -Workbooks     : data/xlsx (filled sample) + converter/template xlsx
#   -Placeholders  : temporary images for the sample workbook
param([switch]$Template, [switch]$Workbooks, [switch]$Placeholders, [switch]$Tests)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$utf8 = New-Object System.Text.UTF8Encoding($false)
$c = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'content.json'), $utf8) | ConvertFrom-Json
# Korean folder/file names, built from code points so this file stays ASCII
function K([int[]]$cp) { -join ($cp | ForEach-Object { [char]$_ }) }
$dataDirName = K 0xC804,0xC0AC,0xC790,0xB8CC                          # jeonsajaryo
$tmplName    = (K 0xC804,0xC0AC,0xC790,0xB8CC) + '_' + (K 0xC591,0xC2DD) + '.xlsx'
$dataDir = Join-Path $root $dataDirName
$imgDir  = Join-Path $dataDir 'images'

if ($Template) {
  $html = [IO.File]::ReadAllText((Join-Path $root 'src\site-template.html'), $utf8)
  $json = $html | ConvertTo-Json -Compress
  [IO.File]::WriteAllText((Join-Path $root 'converter\site-template.js'), "window.SITE_TEMPLATE = $json;`n", $utf8)
  'site-template.js written'
}

function Esc([string]$s) { [System.Security.SecurityElement]::Escape($s) }
function SheetXml($headers, $rows) {
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
  [void]$sb.Append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
  [void]$sb.Append('<cols>')
  for ($i = 1; $i -le $headers.Count; $i++) {
    $h = [string]$headers[$i-1]
    # long text columns: gyeonggwa, yoyak, uiui, seolmyeong, doip, gyohun, munje, haeseol, puli, chumo, bogi
    $longKeys = @((K 0xACBD,0xACFC), (K 0xC694,0xC57D), (K 0xC758,0xC758), (K 0xC124,0xBA85), (K 0xB3C4,0xC785),
                  (K 0xAD50,0xD6C8), (K 0xBB38,0xC81C), (K 0xD574,0xC124), (K 0xD480,0xC774), (K 0xCD94,0xBAA8), (K 0xBCF4,0xAE30),
                  (K 0xC0C1,0xD669), (K 0xADFC,0xAC70), (K 0xB0B4,0xC6A9), (K 0xACB0,0xACFC), (K 0xC120,0xD0DD))
    $long = $false; foreach ($lk in $longKeys) { if ($h.Contains($lk)) { $long = $true } }
    $w = if ($long) { 48 } else { 16 }
    [void]$sb.Append("<col min=""$i"" max=""$i"" width=""$w"" style=""2"" customWidth=""1""/>")
  }
  [void]$sb.Append('</cols><sheetData>')
  $all = @(,@($headers)) + @($rows)
  for ($r = 0; $r -lt $all.Count; $r++) {
    $rn = $r + 1
    [void]$sb.Append("<row r=""$rn"">")
    $row = $all[$r]
    for ($k = 0; $k -lt $row.Count; $k++) {
      $ref = [string][char](65 + $k) + $rn
      $v = $row[$k]
      $s = if ($r -eq 0) { 1 } else { 2 }
      if ($v -is [int] -or $v -is [long] -or $v -is [double]) {
        [void]$sb.Append("<c r=""$ref"" s=""$s""><v>$v</v></c>")
      } elseif ("$v" -ne '') {
        [void]$sb.Append("<c r=""$ref"" s=""$s"" t=""inlineStr""><is><t xml:space=""preserve"">$(Esc "$v")</t></is></c>")
      }
    }
    [void]$sb.Append('</row>')
  }
  [void]$sb.Append('</sheetData></worksheet>')
  $sb.ToString()
}

function WriteXlsx($path, $sheets) {
  Add-Type -AssemblyName System.IO.Compression
  if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path }
  $fs = [IO.File]::Open($path, 'CreateNew')
  $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
  function Put($name, $text) {
    $e = $zip.CreateEntry($name); $st = $e.Open()
    $b = $utf8.GetBytes($text); $st.Write($b, 0, $b.Length); $st.Close()
  }
  $ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
  $wbs = ''; $rels = ''
  for ($i = 1; $i -le $sheets.Count; $i++) {
    $ct += "<Override PartName=""/xl/worksheets/sheet$i.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>"
    $wbs += "<sheet name=""$(Esc $sheets[$i-1].name)"" sheetId=""$i"" r:id=""rId$i""/>"
    $rels += "<Relationship Id=""rId$i"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet$i.xml""/>"
  }
  $n = $sheets.Count + 1
  $rels += "<Relationship Id=""rId$n"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/>"
  Put '[Content_Types].xml' ($ct + '</Types>')
  Put '_rels/.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
  Put 'xl/workbook.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + $wbs + '</sheets></workbook>')
  Put 'xl/_rels/workbook.xml.rels' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + $rels + '</Relationships>')
  Put 'xl/styles.xml' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Malgun Gothic"/></font><font><b/><sz val="11"/><name val="Malgun Gothic"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFDE9C2"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="49" fontId="1" fillId="2" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1"/><xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf></cellXfs></styleSheet>'
  for ($i = 1; $i -le $sheets.Count; $i++) { Put "xl/worksheets/sheet$i.xml" $sheets[$i-1].xml }
  $zip.Dispose(); $fs.Dispose()
}

# optional sheets: quiz, term, source(geungeo), decision(gyeolsim), revision(gaejeong)
function OptionalSheets($q, $t, $src = $c.sources, $dec = $c.decisions, $rev = $c.revisions, [switch]$Empty) {
  if ($Empty) { $q = @(); $t = @(); $src = @(); $dec = @(); $rev = @() }
  @(
    @{ name = (K 0xD034,0xC988); xml = (SheetXml $c.quizHeaders     @($q)) },
    @{ name = (K 0xC6A9,0xC5B4); xml = (SheetXml $c.termHeaders     @($t)) },
    @{ name = (K 0xADFC,0xAC70); xml = (SheetXml $c.sourceHeaders   @($src)) },
    @{ name = (K 0xACB0,0xC2EC); xml = (SheetXml $c.decisionHeaders @($dec)) },
    @{ name = (K 0xAC1C,0xC815); xml = (SheetXml $c.revisionHeaders @($rev)) })
}

if ($Workbooks) {
  New-Item -ItemType Directory -Force $dataDir | Out-Null
  $sB = K 0xC804,0xD22C   # battle sheet
  $sP = K 0xC790,0xB8CC   # material sheet
  $full = @(
    @{ name = $sB; xml = (SheetXml $c.battleHeaders $c.battles) },
    @{ name = $sP; xml = (SheetXml $c.photoHeaders  $c.photos) }) + (OptionalSheets $c.quizzes $c.terms)
  WriteXlsx (Join-Path $dataDir "$dataDirName.xlsx") $full
  $empty = @(
    @{ name = $sB; xml = (SheetXml $c.battleHeaders @()) },
    @{ name = $sP; xml = (SheetXml $c.photoHeaders  @()) }) + (OptionalSheets @() @() -Empty)
  WriteXlsx (Join-Path $root "converter\$tmplName") $empty
  'workbooks written'
}

if ($Tests) {
  # tests/cases.json -> tests/*.xlsx (same battles/photos as the sample, quiz/term sheets replaced)
  $t = [IO.File]::ReadAllText((Join-Path $root 'tests\cases.json'), $utf8) | ConvertFrom-Json
  $sB = K 0xC804,0xD22C; $sP = K 0xC790,0xB8CC; $sQ = K 0xD034,0xC988; $sT = K 0xC6A9,0xC5B4
  foreach ($case in $t.cases) {
    $sheets = @(
      @{ name = $sB; xml = (SheetXml $c.battleHeaders $c.battles) },
      @{ name = $sP; xml = (SheetXml $c.photoHeaders  $c.photos) })
    if (-not $case.noQuizSheets) {
      $q = if ($case.quizzes) { @($case.quizzes) } else { $c.quizzes }
      $tm = if ($case.terms) { @($case.terms) } else { $c.terms }
      $src = if ($case.sources) { @($case.sources) } else { $c.sources }
      $dec = if ($case.decisions) { @($case.decisions) } else { $c.decisions }
      $rev = if ($case.revisions) { @($case.revisions) } else { $c.revisions }
      $sheets += OptionalSheets $q $tm $src $dec $rev
    }
    WriteXlsx (Join-Path $root ('tests\' + $case.file)) $sheets
  }
  'test workbooks written: ' + $t.cases.Count
}

if ($Placeholders) {
  Add-Type -AssemblyName System.Drawing
  New-Item -ItemType Directory -Force $imgDir | Out-Null
  $names = @{}; foreach ($b in $c.battles) { $names[$b[0]] = $b[1] }
  $hues = @{ INC = @(24,46,74); DSS = @(52,64,40); TBD = @(70,52,30); YPD = @(60,30,36) }
  $codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object MimeType -eq 'image/jpeg'
  $ep = New-Object System.Drawing.Imaging.EncoderParameters(1)
  $ep.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality, [long]88)
  foreach ($p in $c.photos) {
    if ($p[10] -ne 'O') { continue }
    $W = 2400; $H = 1600
    $bmp = New-Object System.Drawing.Bitmap($W, $H)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'; $g.TextRenderingHint = 'AntiAliasGridFit'
    $hu = $hues[$p[4]]
    $c1 = [System.Drawing.Color]::FromArgb($hu[0], $hu[1], $hu[2]); $c2 = [System.Drawing.Color]::FromArgb(14, 14, 16)
    $br = New-Object System.Drawing.Drawing2D.LinearGradientBrush((New-Object System.Drawing.Rectangle(0, 0, $W, $H)), $c1, $c2, 60.0)
    $g.FillRectangle($br, 0, 0, $W, $H)
    $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(40, 255, 255, 255), 3)
    for ($x = -$H; $x -lt $W; $x += 90) { $g.DrawLine($pen, $x, $H, $x + $H, 0) }
    $sf = New-Object System.Drawing.StringFormat; $sf.Alignment = 'Center'; $sf.LineAlignment = 'Center'
    $f1 = New-Object System.Drawing.Font('Malgun Gothic', 120, [System.Drawing.FontStyle]::Bold, 'Pixel')
    $f2 = New-Object System.Drawing.Font('Malgun Gothic', 64, [System.Drawing.FontStyle]::Regular, 'Pixel')
    $f3 = New-Object System.Drawing.Font('Malgun Gothic', 48, [System.Drawing.FontStyle]::Regular, 'Pixel')
    $white = [System.Drawing.Brushes]::White
    $gold = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(242, 183, 5))
    $g.DrawString($names[$p[4]], $f1, $white, (New-Object System.Drawing.RectangleF(0, 480, $W, 200)), $sf)
    $g.DrawString(($p[2] -replace '\s*\(.*\)$', ''), $f2, $gold, (New-Object System.Drawing.RectangleF(0, 700, $W, 120)), $sf)
    $g.DrawString($c.placeholderSub, $f3, $white, (New-Object System.Drawing.RectangleF(0, 880, $W, 100)), $sf)
    $g.Dispose()
    $bmp.Save((Join-Path $imgDir $p[1]), $codec, $ep); $bmp.Dispose()
  }
  'placeholders written: ' + (Get-ChildItem $imgDir).Count
}
