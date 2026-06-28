# 高校さがしレポート PowerPoint generator (PowerPoint COM automation)
$ErrorActionPreference = "Stop"

function CRGB([string]$hex){
  $r=[Convert]::ToInt32($hex.Substring(0,2),16)
  $g=[Convert]::ToInt32($hex.Substring(2,2),16)
  $b=[Convert]::ToInt32($hex.Substring(4,2),16)
  return ($r + $g*256 + $b*65536)
}

# Palette
$NAVY="1E2761"; $BG="F4F6FB"; $INK="222B45"; $MUTED="6B7280"
$WHITE="FFFFFF"; $CORAL="F96167"; $ICE="CADCFC"; $CARD="FFFFFF"; $LINE="D8DEEC"
$FONT="Meiryo"

$msoShapeRect=1; $msoShapeRound=5; $msoShapeOval=9; $msoShapeTri=7; $msoShapeTrap=3
$msoTrue=-1; $msoFalse=0
$alnL=1; $alnC=2; $alnR=3
$anTop=1; $anMid=3

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Add(0)   # WithWindow = msoFalse
$pres.PageSetup.SlideWidth = 960
$pres.PageSetup.SlideHeight = 540

function New-Slide(){
  $s = $pres.Slides.Add($pres.Slides.Count+1, 12)  # ppLayoutBlank
  return $s
}
function Set-Bg($s,$hex){
  $s.FollowMasterBackground = $msoFalse
  $s.Background.Fill.Solid()
  $s.Background.Fill.ForeColor.RGB = (CRGB $hex)
}
function Add-Rect($s,$type,$l,$t,$w,$h,$fillHex,$lineHex,$lineW){
  $sp = $s.Shapes.AddShape($type,$l,$t,$w,$h)
  if($fillHex){ $sp.Fill.Solid(); $sp.Fill.ForeColor.RGB=(CRGB $fillHex) } else { $sp.Fill.Visible=$msoFalse }
  if($lineHex){ $sp.Line.Visible=$msoTrue; $sp.Line.ForeColor.RGB=[int](CRGB $lineHex); $sp.Line.Weight=[single]$lineW } else { $sp.Line.Visible=$msoFalse }
  $sp.Shadow.Visible=$msoFalse
  return $sp
}
function Add-Text($s,$l,$t,$w,$h,$text,$size,$colorHex,$bold,$align,$anchor){
  $sp = $s.Shapes.AddTextbox(1,$l,$t,$w,$h)
  $tf = $sp.TextFrame
  $tf.WordWrap=$msoTrue
  $tf.MarginLeft=2; $tf.MarginRight=2; $tf.MarginTop=1; $tf.MarginBottom=1
  $tf.VerticalAnchor=$anchor
  $tr = $tf.TextRange
  $tr.Text = $text
  $tr.Font.Size=[single]$size
  $tr.Font.Name=$FONT
  $tr.Font.Color.RGB=[int](CRGB $colorHex)
  if($bold){ $tr.Font.Bold=$msoTrue } else { $tr.Font.Bold=$msoFalse }
  $tr.ParagraphFormat.Alignment=$align
  return $sp
}
function Add-Link($s,$l,$t,$w,$h,$text,$size,$colorHex,$url){
  $sp = Add-Text $s $l $t $w $h $text $size $colorHex $false $alnL $anTop
  $sp.TextFrame.TextRange.ActionSettings.Item(1).Hyperlink.Address = $url
  return $sp
}

function Add-Uniform($s,$cx,$topY,$blazerHex,$skirtHex,$tieHex){
  # skirt (behind), trapezoid wide bottom
  Add-Rect $s $msoShapeTrap ($cx-62) ($topY+118) 124 80 $skirtHex $null 0 | Out-Null
  # blazer torso
  Add-Rect $s $msoShapeRound ($cx-55) $topY 110 122 $blazerHex $null 0 | Out-Null
  # white shirt V (inverted triangle)
  $shirt = Add-Rect $s $msoShapeTri ($cx-24) ($topY-3) 48 48 $WHITE $null 0
  $shirt.Rotation=[single]180
  # tie
  Add-Rect $s $msoShapeRect ($cx-6) ($topY+16) 12 54 $tieHex $null 0 | Out-Null
  # head
  Add-Rect $s $msoShapeOval ($cx-20) ($topY-46) 40 40 "F1D2B0" $null 0 | Out-Null
}

# ============ SLIDE 1 : TITLE ============
$s = New-Slide
Set-Bg $s $NAVY
Add-Rect $s $msoShapeOval 760 -120 420 420 "2A3A78" $null 0 | Out-Null
Add-Rect $s $msoShapeOval -140 360 360 360 "2A3A78" $null 0 | Out-Null
Add-Rect $s $msoShapeOval 60 60 70 70 $CORAL $null 0 | Out-Null
Add-Text $s 60 56 600 40 "五郎レポート" 16 $ICE $false $alnL $anTop | Out-Null
Add-Text $s 60 170 840 130 "紗衣の 高校さがし" 60 $WHITE $true $alnL $anTop | Out-Null
Add-Text $s 64 300 840 60 "名古屋市内・公立・共学で さがした候補4校" 24 $ICE $false $alnL $anTop | Out-Null
Add-Rect $s $msoShapeRound 64 380 560 56 "2A3A78" $null 0 | Out-Null
Add-Text $s 84 384 540 48 "偏差値50前後／バドミントン部／部活が活発／評判◎" 16 $WHITE $false $alnL $anMid | Out-Null
Add-Text $s 60 470 840 40 "作成：五郎（AI秘書）　2026年6月　ご自宅PCにて" 13 "9DB0E8" $false $alnL $anTop | Out-Null

# ============ SLIDE 2 : CONDITIONS ============
$s = New-Slide
Set-Bg $s $BG
Add-Text $s 40 34 880 50 "さがした条件（6つ ＋ エリア）" 32 $NAVY $true $alnL $anTop | Out-Null
Add-Text $s 42 86 880 30 "お父さんのメモをもとに、五郎が整理しました" 15 $MUTED $false $alnL $anTop | Out-Null
$conds = @(
 @("1","公立","私立ではなく公立高校"),
 @("2","共学","男女いっしょに学ぶ学校"),
 @("3","偏差値 50前後","学力・内申の目安ライン"),
 @("4","バドミントン部","やりたい部活があること"),
 @("5","行事・部活が活発","楽しい高校生活を送れる"),
 @("6","評判がいい","通ってよかったと思える")
)
$cardW=280; $cardH=120; $gx=20; $gy=20; $x0=40; $y0=130
for($i=0;$i -lt 6;$i++){
  $col=$i%3; $row=[int]([math]::Floor($i/3))
  $l=$x0+$col*($cardW+$gx); $t=$y0+$row*($cardH+$gy)
  Add-Rect $s $msoShapeRound $l $t $cardW $cardH $CARD $LINE 1 | Out-Null
  $badge = Add-Rect $s $msoShapeOval ($l+18) ($t+20) 44 44 $CORAL $null 0
  Add-Text $s ($l+18) ($t+20) 44 44 $conds[$i][0] 22 $WHITE $true $alnC $anMid | Out-Null
  Add-Text $s ($l+74) ($t+22) ($cardW-90) 40 $conds[$i][1] 19 $NAVY $true $alnL $anMid | Out-Null
  Add-Text $s ($l+22) ($t+72) ($cardW-40) 36 $conds[$i][2] 14 $INK $false $alnL $anTop | Out-Null
}
Add-Rect $s $msoShapeRound 40 408 880 70 "E9EEFA" $null 0 | Out-Null
Add-Text $s 60 420 880 46 "エリア：名古屋市内（ご自宅＝北区鳩岡）から通えること。北区を中心に近いエリアでさがしました。" 16 $NAVY $true $alnL $anMid | Out-Null

# ============ SLIDE 3 : 4 CANDIDATES TABLE ============
$s = New-Slide
Set-Bg $s $BG
Add-Text $s 40 34 880 50 "候補は4校（北区から通えるエリア）" 32 $NAVY $true $alnL $anTop | Out-Null
$cols = @("学校","区・市","偏差値","バド部","評判","通学")
$cw = @(250,150,110,110,120,140)
$rows = @(
 @("名古屋市立 北高校","北区","50","あり","2.9","◎ 同じ北区"),
 @("名古屋市立 山田高校","西区","49","◎人気","3.5","〇 上小田井"),
 @("愛知県立 中村高校","中村区","50","県大会","--","〇 中村区"),
 @("愛知県立 春日井南高校","春日井市","54","要確認","3.9","△ 市外/電車")
)
$accents = @("0E8C8B","E8633F","7B5EA7","2E8B57")
$tx=40; $ty=100; $rh=64; $hh=42
$cx=$tx
for($c=0;$c -lt 6;$c++){
  Add-Rect $s $msoShapeRect $cx $ty $cw[$c] $hh $NAVY $null 0 | Out-Null
  Add-Text $s $cx $ty $cw[$c] $hh $cols[$c] 15 $WHITE $true $alnC $anMid | Out-Null
  $cx += $cw[$c]
}
for($r=0;$r -lt 4;$r++){
  $rt = $ty+$hh+$r*$rh
  $rowbg = if($r%2 -eq 0){"FFFFFF"}else{"EEF2FB"}
  $cx=$tx
  for($c=0;$c -lt 6;$c++){
    Add-Rect $s $msoShapeRect $cx $rt $cw[$c] $rh $rowbg $LINE 0.75 | Out-Null
    $val=$rows[$r][$c]
    if($c -eq 0){
      Add-Rect $s $msoShapeOval ($cx+10) ($rt+$rh/2-7) 14 14 $accents[$r] $null 0 | Out-Null
      Add-Text $s ($cx+30) $rt ($cw[$c]-34) $rh $val 14 $NAVY $true $alnL $anMid | Out-Null
    } elseif($c -eq 2){
      Add-Text $s $cx $rt $cw[$c] $rh $val 20 $accents[$r] $true $alnC $anMid | Out-Null
    } else {
      Add-Text $s $cx $rt $cw[$c] $rh $val 14 $INK $false $alnC $anMid | Out-Null
    }
    $cx += $cw[$c]
  }
}
Add-Text $s 40 430 880 70 "※偏差値・評判の数値は2026年6月時点の進学情報サイト等を参照（みんなの高校情報ほか）。倍率・募集人数は年度で変わります。最終確認は必ず各校の公式サイト・募集要項で。" 12 $MUTED $false $alnL $anTop | Out-Null

# ============ SCHOOL DETAIL SLIDES ============
function School-Slide($name,$area,$accent,$hensa,$scale,$rows,$blazer,$skirt,$tie,$linkText,$linkUrl){
  $s = New-Slide
  Set-Bg $s $WHITE
  Add-Rect $s $msoShapeRound 40 32 650 56 $accent $null 0 | Out-Null
  Add-Text $s 60 32 620 56 $name 27 $WHITE $true $alnL $anMid | Out-Null
  Add-Text $s 44 96 640 28 $area 15 $MUTED $false $alnL $anTop | Out-Null
  # hensa callout
  Add-Rect $s $msoShapeRound 712 32 208 110 "F2F5FC" $accent 1.25 | Out-Null
  Add-Text $s 712 42 208 24 "偏差値（目安）" 13 $MUTED $false $alnC $anTop | Out-Null
  Add-Text $s 712 60 208 56 $hensa 48 $accent $true $alnC $anTop | Out-Null
  Add-Text $s 712 118 208 20 $scale 11 $MUTED $false $alnC $anTop | Out-Null
  # info rows
  $ix=44; $iy=140; $rh=46; $labW=150; $valW=460
  for($i=0;$i -lt $rows.Count;$i++){
    $t=$iy+$i*$rh
    Add-Rect $s $msoShapeRound $ix $t ($labW+$valW+16) ($rh-8) "F7F9FD" $null 0 | Out-Null
    Add-Text $s ($ix+12) $t $labW ($rh-8) $rows[$i][0] 14 $accent $true $alnL $anMid | Out-Null
    Add-Text $s ($ix+12+$labW) $t $valW ($rh-8) $rows[$i][1] 14 $INK $false $alnL $anMid | Out-Null
  }
  # uniform figure
  Add-Uniform $s 805 210 $blazer $skirt $tie
  Add-Text $s 700 405 220 22 "▲ 制服イメージ図" 12 $MUTED $false $alnC $anTop | Out-Null
  Add-Text $s 700 425 220 36 "※色・デザインは目安。実物は下のリンクで確認を" 10 $MUTED $false $alnC $anTop | Out-Null
  # link
  Add-Rect $s $msoShapeRound 44 500 876 34 "EEF2FB" $null 0 | Out-Null
  Add-Link $s 56 500 864 34 $linkText 13 "1A56C4" $linkUrl | Out-Null
}

# Slide 4: 北高校
School-Slide "名古屋市立 北高校（きた）" "名古屋市北区如来町50　ご自宅と同じ北区＝いちばん近い" "0E8C8B" "50" "市内・公立で中堅" @(
 @("学科・コース","普通科（特進コース／国際理解コース）"),
 @("バドミントン部","あり。部活では音楽部が全国レベルで有名"),
 @("評判","★2.9 / 5（口コミは賛否あり・課題が多めとの声）"),
 @("進学実績","約7割が大学進学。地元国公立（愛知教育大・名市大・県立大）も"),
 @("通学","◎ 自宅と同じ北区。最寄りはJR城北線『比良』＋市バス"),
 @("ひとこと","地元で通いやすさは一番。落ち着いた校風")
) "23304E" "6E7480" "F96167" "▶ くわしい学校情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）" "https://www.minkou.jp/hischool/school/3391/"

# Slide 5: 山田高校
School-Slide "名古屋市立 山田高校（やまだ）" "名古屋市西区二方町19-1　最寄り『上小田井』駅" "E8633F" "49" "市内・公立で中堅" @(
 @("学科・コース","普通科"),
 @("バドミントン部","あり。部員が多く人気の部活。部活全般の評価が高い"),
 @("評判","★3.5 / 5（明るい校風・楽しい学校との声・制服も人気）"),
 @("進学実績","同志社・立命館・関西・愛知工大など私大中心"),
 @("通学","〇 上小田井（地下鉄鶴舞線／名鉄犬山線）で北区から乗換で"),
 @("ひとこと","行事・部活が活発。倍率は約2.0倍と人気高め")
) "23304E" "2E3A66" "C0C4CC" "▶ くわしい学校情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）" "https://www.minkou.jp/hischool/school/3402/"

# Slide 6: 中村高校
School-Slide "愛知県立 中村高校（なかむら）" "名古屋市中村区　自由な校風（1953年開校）" "7B5EA7" "50" "市内・公立で中堅" @(
 @("学科・コース","普通科"),
 @("バドミントン部","あり。県大会出場。弓道・陸上・体操なども熱心"),
 @("評判","自由でのびのびした校風。部活に打ち込みやすい"),
 @("進学実績","私立大学中心に幅広く進学"),
 @("通学","〇 中村区。地下鉄・市バスで北区から通学圏"),
 @("ひとこと","部活が活発で、自由な雰囲気を求める子に合う")
) "23304E" "41506E" "7B5EA7" "▶ くわしい学校情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）" "https://www.minkou.jp/hischool/school/123/"

# Slide 7: 春日井南高校
School-Slide "愛知県立 春日井南高校（チャレンジ校）" "春日井市　偏差値は少し上。北区から電車で通学圏" "2E8B57" "54" "やや上のレベル" @(
 @("学科・コース","普通科"),
 @("バドミントン部","要確認（運動部14・文化部10と部活は豊富。ハンドボール全国級）"),
 @("評判","★3.9 / 5（県内35位/222校）校舎がきれい・自習室が夜まで"),
 @("進学実績","国公立に毎年20名前後。中部大・名城大・中京大ほか"),
 @("通学","△ 市外。JR中央線／名鉄小牧線方面で北区から通学"),
 @("ひとこと","学習環境◎。少しがんばる『チャレンジ校』候補")
) "23304E" "3A4A38" "2E8B57" "▶ くわしい学校情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）" "https://www.minkou.jp/hischool/school/52/"

# ============ SLIDE 8 : COMPARISON ============
$s = New-Slide
Set-Bg $s $BG
Add-Text $s 40 34 880 50 "4校をくらべると（早見表）" 32 $NAVY $true $alnL $anTop | Out-Null
$crit = @("偏差値","バドミントン部","評判","通学のしやすさ","おすすめポイント")
$schoolNames = @("北高校","山田高校","中村高校","春日井南")
$accents2 = @("0E8C8B","E8633F","7B5EA7","2E8B57")
$data = @(
 @("50","49","50","54"),
 @("あり","◎人気","県大会","要確認"),
 @("★2.9","★3.5","--","★3.9"),
 @("◎ 北区","〇 西区","〇 中村","△ 市外"),
 @("地元で近い","部活活発・評判◎","自由な校風","環境◎・上狙い")
)
$lx=40; $ly=100; $lw=170; $colw=170; $rh=64; $hh=44
# header
Add-Rect $s $msoShapeRect $lx $ly $lw $hh $NAVY $null 0 | Out-Null
Add-Text $s $lx $ly $lw $hh "項目" 14 $WHITE $true $alnC $anMid | Out-Null
for($c=0;$c -lt 4;$c++){
  $cx=$lx+$lw+$c*$colw
  Add-Rect $s $msoShapeRect $cx $ly $colw $hh $accents2[$c] $null 0 | Out-Null
  Add-Text $s $cx $ly $colw $hh $schoolNames[$c] 14 $WHITE $true $alnC $anMid | Out-Null
}
for($r=0;$r -lt 5;$r++){
  $rt=$ly+$hh+$r*$rh
  Add-Rect $s $msoShapeRect $lx $rt $lw $rh "E4E9F5" $LINE 0.75 | Out-Null
  Add-Text $s ($lx+10) $rt ($lw-16) $rh $crit[$r] 13 $NAVY $true $alnL $anMid | Out-Null
  for($c=0;$c -lt 4;$c++){
    $cx=$lx+$lw+$c*$colw
    $rowbg = if($r%2 -eq 0){"FFFFFF"}else{"EEF2FB"}
    Add-Rect $s $msoShapeRect $cx $rt $colw $rh $rowbg $LINE 0.75 | Out-Null
    $val=$data[$r][$c]
    if($r -eq 0){ Add-Text $s $cx $rt $colw $rh $val 19 $accents2[$c] $true $alnC $anMid | Out-Null }
    else { Add-Text $s ($cx+6) $rt ($colw-12) $rh $val 12.5 $INK $false $alnC $anMid | Out-Null }
  }
}

# ============ SLIDE 9 : NEXT STEPS ============
$s = New-Slide
Set-Bg $s $NAVY
Add-Text $s 40 40 880 50 "次にやること（おすすめの進め方）" 32 $WHITE $true $alnL $anTop | Out-Null
$steps = @(
 @("1","学校説明会・体験入学に行く","パンフより『行ってみた感じ』が大事。親子で雰囲気を確かめる"),
 @("2","バドミントン部を見学する","部の強さ・人数・雰囲気を実際に見る。入りたい部かを確認"),
 @("3","通学を実際に試す","北区の自宅から朝の時間帯に1往復。乗換・所要時間を体感"),
 @("4","倍率・募集要項を公式で確認","この資料の数値は目安。最新の倍率・日程は必ず公式サイトで")
)
for($i=0;$i -lt 4;$i++){
  $t=110+$i*86
  Add-Rect $s $msoShapeRound 40 $t 880 74 "2A3A78" $null 0 | Out-Null
  Add-Rect $s $msoShapeOval 58 ($t+17) 40 40 $CORAL $null 0 | Out-Null
  Add-Text $s 58 ($t+17) 40 40 $steps[$i][0] 20 $WHITE $true $alnC $anMid | Out-Null
  Add-Text $s 116 ($t+12) 790 30 $steps[$i][1] 18 $WHITE $true $alnL $anTop | Out-Null
  Add-Text $s 116 ($t+42) 790 26 $steps[$i][2] 13 $ICE $false $alnL $anTop | Out-Null
}
Add-Text $s 40 470 880 50 "五郎より：まずは『北高校』と『山田高校』の説明会から。気になる順に一緒に動きましょう。応援しています。" 14 "9DB0E8" $false $alnL $anTop | Out-Null

# Save
$out = "C:\Users\山本　史朗\OneDrive\Desktop\新しいフォルダー\高校さがしレポート_五郎.pptx"
if(Test-Path $out){ Remove-Item $out -Force }
$pres.SaveAs($out, 24)   # ppSaveAsOpenXMLPresentation
$pres.Close()
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
[GC]::Collect()
Write-Output ("SAVED: " + $out)
Write-Output ("SIZE: " + ((Get-Item $out).Length) + " bytes")
