# 高校さがし 公立編・私立編 v4 : 各校2ページ(基本／学校生活)・部活/校風くわしく版
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing

function CRGB([string]$hex){ $r=[Convert]::ToInt32($hex.Substring(0,2),16);$g=[Convert]::ToInt32($hex.Substring(2,2),16);$b=[Convert]::ToInt32($hex.Substring(4,2),16); return ($r+$g*256+$b*65536) }
$NAVY="1E2761"; $BG="F4F6FB"; $INK="222B45"; $MUTED="6B7280"
$WHITE="FFFFFF"; $CORAL="F96167"; $ICE="CADCFC"; $LINE="D8DEEC"
$GREENBG="E7F6EC"; $GREENINK="1E7A43"; $REDBG="FBEAEA"; $REDINK="B23B3B"
$BLUEBG="EAF2FB"; $PURPBG="F3EFFA"; $PURPINK="6B46A8"
$FONT="Meiryo"
$msoShapeRect=1; $msoShapeRound=5; $msoShapeOval=9
$msoTrue=-1; $msoFalse=0; $alnL=1; $alnC=2; $anTop=1; $anMid=3

$ppt = New-Object -ComObject PowerPoint.Application
function Set-Bg($s,$hex){ $s.FollowMasterBackground=$msoFalse; $s.Background.Fill.Solid(); $s.Background.Fill.ForeColor.RGB=[int](CRGB $hex) }
function Add-Rect($s,$type,$l,$t,$w,$h,$fillHex,$lineHex,$lineW){
  $sp=$s.Shapes.AddShape($type,$l,$t,$w,$h)
  if($fillHex){ $sp.Fill.Solid(); $sp.Fill.ForeColor.RGB=[int](CRGB $fillHex) } else { $sp.Fill.Visible=$msoFalse }
  if($lineHex){ $sp.Line.Visible=$msoTrue; $sp.Line.ForeColor.RGB=[int](CRGB $lineHex); $sp.Line.Weight=[single]$lineW } else { $sp.Line.Visible=$msoFalse }
  $sp.Shadow.Visible=$msoFalse; return $sp
}
function Add-Text($s,$l,$t,$w,$h,$text,$size,$colorHex,$bold,$align,$anchor){
  $sp=$s.Shapes.AddTextbox(1,$l,$t,$w,$h); $tf=$sp.TextFrame
  $tf.WordWrap=$msoTrue; $tf.MarginLeft=3; $tf.MarginRight=3; $tf.MarginTop=1; $tf.MarginBottom=1
  $tf.VerticalAnchor=$anchor; $tr=$tf.TextRange; $tr.Text=$text
  $tr.Font.Size=[single]$size; $tr.Font.Name=$FONT; $tr.Font.Color.RGB=[int](CRGB $colorHex)
  if($bold){$tr.Font.Bold=$msoTrue}else{$tr.Font.Bold=$msoFalse}
  $tr.ParagraphFormat.Alignment=$align; return $sp
}
function Add-Para($s,$l,$t,$w,$h,$lines,$size,$colorHex){ $txt=($lines|ForEach-Object{"・"+$_}) -join "`r"; return (Add-Text $s $l $t $w $h $txt $size $colorHex $false $alnL $anTop) }
function Add-Link($s,$l,$t,$w,$h,$text,$size,$colorHex,$url){ $sp=Add-Text $s $l $t $w $h $text $size $colorHex $false $alnL $anMid; $sp.TextFrame.TextRange.ActionSettings.Item(1).Hyperlink.Address=$url; return $sp }
function Add-Picture($s,$path,$boxX,$boxY,$boxW,$boxH){
  $img=[System.Drawing.Image]::FromFile($path); $iw=$img.Width; $ih=$img.Height; $img.Dispose()
  $scale=[math]::Min($boxW/$iw,$boxH/$ih); $w=$iw*$scale; $h=$ih*$scale
  $x=$boxX+($boxW-$w)/2; $y=$boxY+($boxH-$h)/2
  $pic=$s.Shapes.AddPicture($path,$msoFalse,$msoTrue,$x,$y,$w,$h)
  $pic.Line.Visible=$msoTrue; $pic.Line.ForeColor.RGB=[int](CRGB "FFFFFF"); $pic.Line.Weight=[single]3; return $pic
}

function Detail-Pages($pres,$sc,$isPrivate){
  $acc=$sc.accent
  # ---- PAGE A : 基本 ----
  $s=$pres.Slides.Add($pres.Slides.Count+1,12); Set-Bg $s $WHITE
  Add-Rect $s $msoShapeRound 40 26 624 50 $acc $null 0 | Out-Null
  Add-Text $s 58 26 600 50 ($sc.name+"　①基本") 20 $WHITE $true $alnL $anMid | Out-Null
  Add-Text $s 44 80 624 24 $sc.area 13 $MUTED $false $alnL $anTop | Out-Null
  Add-Rect $s $msoShapeRound 686 26 234 76 "F2F5FC" $acc 1.25 | Out-Null
  Add-Text $s 686 32 234 18 "偏差値（目安）" 12 $MUTED $false $alnC $anTop | Out-Null
  Add-Text $s 686 46 234 42 $sc.hensa 26 $acc $true $alnC $anTop | Out-Null
  Add-Text $s 686 84 234 16 $sc.scale 10 $MUTED $false $alnC $anTop | Out-Null
  $chipbg = if($isPrivate){$GREENINK}else{$acc}
  Add-Rect $s $msoShapeRound 686 108 234 28 $chipbg $null 0 | Out-Null
  Add-Text $s 686 108 234 28 $sc.chip 10 $WHITE $true $alnC $anMid | Out-Null
  Add-Picture $s $sc.photo 686 144 234 176 | Out-Null
  Add-Text $s 686 322 234 26 $sc.psrc 9 $MUTED $false $alnC $anTop | Out-Null
  $ix=40;$iy=116;$rh2=60;$labW=128;$valW=484
  for($i=0;$i -lt $sc.rowsA.Count;$i++){
    $t=$iy+$i*$rh2
    Add-Rect $s $msoShapeRound $ix $t ($labW+$valW+12) ($rh2-7) "F7F9FD" $null 0 | Out-Null
    $lc = if($sc.rowsA[$i][0] -eq "食堂・昼食"){$GREENINK}else{$acc}
    Add-Text $s ($ix+12) $t $labW ($rh2-7) $sc.rowsA[$i][0] 13 $lc $true $alnL $anMid | Out-Null
    Add-Text $s ($ix+12+$labW) $t $valW ($rh2-7) $sc.rowsA[$i][1] 12 $INK $false $alnL $anMid | Out-Null
  }
  Add-Rect $s $msoShapeRound 40 372 880 40 "E9EEFA" $null 0 | Out-Null
  Add-Text $s 54 372 866 40 "▶ 次のページ②に【部活動・校風校則・口コミの良い点/気になる点・2026年度の日程】を掲載しています。" 12 $NAVY $true $alnL $anMid | Out-Null
  Add-Rect $s $msoShapeRound 40 486 880 32 "EEF2FB" $null 0 | Out-Null
  Add-Link $s 54 486 866 32 ("▶ " + $sc.link) 12 "1A56C4" $sc.url | Out-Null

  # ---- PAGE B : 学校生活 ----
  $s=$pres.Slides.Add($pres.Slides.Count+1,12); Set-Bg $s $WHITE
  Add-Rect $s $msoShapeRound 40 26 880 46 $acc $null 0 | Out-Null
  Add-Text $s 58 26 860 46 ($sc.name+"　②部活・校風・口コミ") 20 $WHITE $true $alnL $anMid | Out-Null
  # 部活動
  Add-Rect $s $msoShapeRound 40 84 430 152 $BLUEBG $null 0 | Out-Null
  Add-Text $s 56 92 400 22 "■ 部活動" 14 $NAVY $true $alnL $anTop | Out-Null
  Add-Para $s 56 118 404 112 $sc.bukatsu 11.5 $INK | Out-Null
  # 校風・校則
  Add-Rect $s $msoShapeRound 490 84 430 152 $PURPBG $null 0 | Out-Null
  Add-Text $s 506 92 400 22 "■ 校風・校則" 14 $PURPINK $true $alnL $anTop | Out-Null
  Add-Para $s 506 118 404 112 $sc.kosoku 11.5 $INK | Out-Null
  # 口コミ良い点
  Add-Rect $s $msoShapeRound 40 246 430 150 $GREENBG $null 0 | Out-Null
  Add-Text $s 56 254 400 22 "■ 口コミの良い点" 14 $GREENINK $true $alnL $anTop | Out-Null
  Add-Para $s 56 280 404 110 $sc.pros 11.5 $INK | Out-Null
  # 気になる点
  Add-Rect $s $msoShapeRound 490 246 430 150 $REDBG $null 0 | Out-Null
  Add-Text $s 506 254 400 22 "■ 気になる点" 14 $REDINK $true $alnL $anTop | Out-Null
  Add-Para $s 506 280 404 110 $sc.cons 11.5 $INK | Out-Null
  # 2026 events
  Add-Rect $s $msoShapeRound 40 406 880 70 "FFF4E8" $null 0 | Out-Null
  Add-Text $s 54 410 880 22 "■ 2026年度（令和8年度）の日程 ※最新は公式で要確認" 12 "B5651D" $true $alnL $anTop | Out-Null
  Add-Text $s 54 432 880 22 ("説明会・見学：" + $sc.ev_setsu) 11 $INK $false $alnL $anTop | Out-Null
  Add-Text $s 54 452 880 22 ("文化祭：" + $sc.ev_bunka) 11 $INK $false $alnL $anTop | Out-Null
  Add-Rect $s $msoShapeRound 40 484 880 30 "EEF2FB" $null 0 | Out-Null
  Add-Link $s 54 484 866 30 ("▶ " + $sc.link) 11 "1A56C4" $sc.url | Out-Null
}

function Build-Deck($outPath,$title,$sub,$schools,$isPrivate,$conds,$tableCols,$tableC3green,$compCrit,$steps,$goro){
  $N=$schools.Count
  $pres=$ppt.Presentations.Add(0); $pres.PageSetup.SlideWidth=960; $pres.PageSetup.SlideHeight=540
  # COVER
  $s=$pres.Slides.Add(1,12); Set-Bg $s $NAVY
  Add-Rect $s $msoShapeOval 760 -120 420 420 "2A3A78" $null 0 | Out-Null
  Add-Rect $s $msoShapeOval -140 360 360 360 "2A3A78" $null 0 | Out-Null
  Add-Rect $s $msoShapeOval 60 64 64 64 $CORAL $null 0 | Out-Null
  Add-Text $s 60 60 600 40 "五郎レポート" 16 $ICE $false $alnL $anTop | Out-Null
  Add-Text $s 60 168 860 110 $title 38 $WHITE $true $alnL $anTop | Out-Null
  Add-Text $s 64 300 860 50 $sub 16 $ICE $false $alnL $anTop | Out-Null
  Add-Text $s 60 470 860 40 "作成：五郎（AI秘書）　2026年6月　自宅PC（名古屋市北区鳩岡）にて" 13 "9DB0E8" $false $alnL $anTop | Out-Null
  # CONDITIONS
  $s=$pres.Slides.Add(2,12); Set-Bg $s $BG
  Add-Text $s 40 30 880 50 "さがした条件（6つ ＋ エリア）" 30 $NAVY $true $alnL $anTop | Out-Null
  $cardW=280;$cardH=112;$gx=20;$gy=18;$x0=40;$y0=92
  for($i=0;$i -lt 6;$i++){
    $col=$i%3;$row=[int]([math]::Floor($i/3));$l=$x0+$col*($cardW+$gx);$t=$y0+$row*($cardH+$gy)
    Add-Rect $s $msoShapeRound $l $t $cardW $cardH $WHITE $LINE 1 | Out-Null
    Add-Rect $s $msoShapeOval ($l+18) ($t+18) 42 42 $CORAL $null 0 | Out-Null
    Add-Text $s ($l+18) ($t+18) 42 42 $conds[$i][0] 21 $WHITE $true $alnC $anMid | Out-Null
    Add-Text $s ($l+72) ($t+18) ($cardW-86) 40 $conds[$i][1] 17 $NAVY $true $alnL $anMid | Out-Null
    Add-Text $s ($l+20) ($t+66) ($cardW-36) 34 $conds[$i][2] 13 $INK $false $alnL $anTop | Out-Null
  }
  Add-Rect $s $msoShapeRound 40 372 880 66 "E9EEFA" $null 0 | Out-Null
  Add-Text $s 60 372 850 66 "エリア：自宅＝名古屋市北区鳩岡。地下鉄「黒川」駅(名城線)／名鉄小牧線「上飯田」／市バスはバス停「城北小学校(北)」を起点に通学。" 15 $NAVY $true $alnL $anMid | Out-Null
  Add-Text $s 40 452 880 60 "※各校2ページ構成(①基本・進学・制服 ②部活動・校風校則・口コミ・2026日程)。偏差値・評判・日程は2026年6月時点の目安。最新は各校公式で確認を。" 12 $MUTED $false $alnL $anTop | Out-Null
  # TABLE
  $s=$pres.Slides.Add(3,12); Set-Bg $s $BG
  Add-Text $s 40 28 880 50 $sub 22 $NAVY $true $alnL $anTop | Out-Null
  $cw=@(210,110,120,150,90,130); $tx=40;$ty=88;$rh=58;$hh=38;$cxp=$tx
  for($c=0;$c -lt 6;$c++){ Add-Rect $s $msoShapeRect $cxp $ty $cw[$c] $hh $NAVY $null 0 | Out-Null; Add-Text $s $cxp $ty $cw[$c] $hh $tableCols[$c] 14 $WHITE $true $alnC $anMid | Out-Null; $cxp+=$cw[$c] }
  for($r=0;$r -lt $N;$r++){
    $rt=$ty+$hh+$r*$rh; $rowbg=if($r%2 -eq 0){$WHITE}else{"EEF2FB"}; $cxp=$tx; $row=$schools[$r].cols
    for($c=0;$c -lt 6;$c++){
      Add-Rect $s $msoShapeRect $cxp $rt $cw[$c] $rh $rowbg $LINE 0.75 | Out-Null; $val=$row[$c]
      if($c -eq 0){ Add-Rect $s $msoShapeOval ($cxp+8) ($rt+$rh/2-6) 12 12 $schools[$r].accent $null 0 | Out-Null; Add-Text $s ($cxp+26) $rt ($cw[$c]-30) $rh $val 12 $NAVY $true $alnL $anMid | Out-Null }
      elseif($c -eq 2){ Add-Text $s $cxp $rt $cw[$c] $rh $val 15 $schools[$r].accent $true $alnC $anMid | Out-Null }
      elseif($c -eq 3 -and $tableC3green){ Add-Text $s $cxp $rt $cw[$c] $rh $val 12 $GREENINK $true $alnC $anMid | Out-Null }
      else { Add-Text $s $cxp $rt $cw[$c] $rh $val 12 $INK $false $alnC $anMid | Out-Null }
      $cxp+=$cw[$c]
    }
  }
  Add-Text $s 40 ($ty+$hh+$N*$rh+8) 880 50 "※「評判」はみんなの高校情報の総合評価(5点満点)。所要時間は自宅(北区鳩岡)からの目安。各校くわしくは次ページ以降(1校2ページ)。" 11 $MUTED $false $alnL $anTop | Out-Null
  # DETAILS (2 pages each)
  foreach($sc in $schools){ Detail-Pages $pres $sc $isPrivate }
  # COMPARISON
  $s=$pres.Slides.Add($pres.Slides.Count+1,12); Set-Bg $s $BG
  Add-Text $s 40 28 880 46 ("くらべると（早見表）") 28 $NAVY $true $alnL $anTop | Out-Null
  $lx=40;$ly=86;$lw=150;$colw=152;$rh3=60;$hh3=40
  if($N -le 4){ $colw=176 }
  Add-Rect $s $msoShapeRect $lx $ly $lw $hh3 $NAVY $null 0 | Out-Null; Add-Text $s $lx $ly $lw $hh3 "項目" 13 $WHITE $true $alnC $anMid | Out-Null
  for($c=0;$c -lt $N;$c++){ $cxp=$lx+$lw+$c*$colw; Add-Rect $s $msoShapeRect $cxp $ly $colw $hh3 $schools[$c].accent $null 0 | Out-Null; Add-Text $s $cxp $ly $colw $hh3 $schools[$c].short 13 $WHITE $true $alnC $anMid | Out-Null }
  for($r=0;$r -lt 5;$r++){
    $rt=$ly+$hh3+$r*$rh3
    Add-Rect $s $msoShapeRect $lx $rt $lw $rh3 "E4E9F5" $LINE 0.75 | Out-Null
    Add-Text $s ($lx+8) $rt ($lw-14) $rh3 $compCrit[$r] 12 $NAVY $true $alnL $anMid | Out-Null
    for($c=0;$c -lt $N;$c++){
      $cxp=$lx+$lw+$c*$colw; $rowbg=if($r%2 -eq 0){$WHITE}else{"EEF2FB"}
      Add-Rect $s $msoShapeRect $cxp $rt $colw $rh3 $rowbg $LINE 0.75 | Out-Null; $val=$schools[$c].cmp[$r]
      if($r -eq 0){ Add-Text $s $cxp $rt $colw $rh3 $val 14 $schools[$c].accent $true $alnC $anMid | Out-Null }
      else { Add-Text $s ($cxp+4) $rt ($colw-8) $rh3 $val 11 $INK $false $alnC $anMid | Out-Null }
    }
  }
  # NEXT STEPS
  $s=$pres.Slides.Add($pres.Slides.Count+1,12); Set-Bg $s $NAVY
  Add-Text $s 40 36 880 50 "次にやること（おすすめの進め方）" 30 $WHITE $true $alnL $anTop | Out-Null
  for($i=0;$i -lt 4;$i++){
    $t=104+$i*84
    Add-Rect $s $msoShapeRound 40 $t 880 72 "2A3A78" $null 0 | Out-Null
    Add-Rect $s $msoShapeOval 58 ($t+16) 40 40 $CORAL $null 0 | Out-Null
    Add-Text $s 58 ($t+16) 40 40 $steps[$i][0] 20 $WHITE $true $alnC $anMid | Out-Null
    Add-Text $s 116 ($t+11) 790 30 $steps[$i][1] 18 $WHITE $true $alnL $anTop | Out-Null
    Add-Text $s 116 ($t+41) 790 26 $steps[$i][2] 13 $ICE $false $alnL $anTop | Out-Null
  }
  Add-Text $s 40 466 880 50 $goro 14 "9DB0E8" $false $alnL $anTop | Out-Null
  if(Test-Path $outPath){ Remove-Item $outPath -Force }
  $pres.SaveAs($outPath,24); $pres.Close()
}

$base="C:\Users\山本　史朗\OneDrive\Desktop\新しいフォルダー\"; $U=$base+"uniforms\"

# ===================== PUBLIC =====================
$publicSchools=@(
 @{ name="名古屋市立 北高校"; short="北高校"; area="名古屋市北区如来町／自宅と同じ北区＝最短"; accent="0E8C8B"; hensa="50"; scale="普通科(特進/国際理解)"; chip="国際理解コースあり"; photo=($U+"kita.jpg"); psrc="制服（参考写真）出典：A-制服店.com";
   rowsA=@(
     @("学科・コース","普通科。2年から特進コース／国際理解コース。国際理解は英語・留学プログラムが充実"),
     @("進学・特色","大学進学が約7割。地元国公立(愛知教育大・名市大・愛知県立大)＋私大。音楽部が全国級"),
     @("最寄り駅・自宅から","JR城北線「比良」駅＋市バス。自宅(北区鳩岡)から自転車約15分／バス停『城北小学校』も。市内同区で最短【目安】"),
     @("通学・その他","公立で学費が抑えられる。お弁当＋購買(専用食堂なし)"));
   bukatsu=@("文化部がとても強い。特に音楽部は全国レベルの強豪","運動部は標準的(目立った全国実績は少なめ)","※野球部・吹奏楽部は設置なし");
   kosoku=@("校則は比較的ゆるめとの声","スマホ：朝〜帰りのSTまで電源OFFでカバン保管(行事中はOK)","髪型は巻く/束ねるは自由／アルバイトは禁止","制服は冬服が好評・夏服は不評との声");
   pros=@("地元(同じ北区)で通いやすい","国際理解コースで英語・留学","音楽部など文化部が全国級","公立で学費が抑えられる");
   cons=@("課題・提出物が多めとの声","駅から遠くバス・自転車が前提","先生により対応差との声");
   cols=@("名古屋市立 北高校","北区","50","音楽全国","2.9","◎北区"); cmp=@("50","音楽部 全国級","★2.9","◎北区","国際理解");
   ev_setsu="国際理解コース説明会 8/4(火)10:00／学校見学会 10/10(土)13:30"; ev_bunka="北高祭＝例年9月頃・要確認";
   link="北高校のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/3391/" },
 @{ name="名古屋市立 山田高校"; short="山田高校"; area="名古屋市西区二方町／最寄り 上小田井"; accent="E8633F"; hensa="49"; scale="普通科"; chip="明るい校風・行事が活発"; photo=($U+"yamada.jpg"); psrc="制服（参考写真）出典：A-制服店.com";
   rowsA=@(
     @("学科・コース","普通科"),
     @("進学・特色","同志社・立命館・関西・愛知工大など私大中心。ダンス部が全国常連で行事も活発"),
     @("最寄り駅・自宅から","地下鉄鶴舞線・名鉄犬山線「上小田井」駅 徒歩約12分。黒川→名城線→乗換で約40分／バス『城北小学校』→地下鉄【目安】"),
     @("通学・その他","公立で学費が抑えられる。お弁当＋購買。制服が人気"));
   bukatsu=@("部活動の満足度が高い(口コミ評価◎)","ダンス部が全国常連","運動部・文化部とも活発で行事も盛ん");
   kosoku=@("明るく楽しい雰囲気(説明会から明るいとの声)","服装は厳しめ：スカート丈・ピアスは没収例あり","カラコン没収、メイクは指導(その場で落とす例)","『派手になりすぎ注意』との声も");
   pros=@("とにかく明るく楽しい校風","部活・行事が盛ん","制服がかわいいと人気","公立で学費が抑えられる");
   cons=@("勉強面の手厚さは控えめとの声","服装・メイクの校則は厳しめ");
   cols=@("名古屋市立 山田高校","西区","49","ダンス全国","3.5","〇上小田井"); cmp=@("49","ダンス 全国","★3.5","〇上小田井","明るい・行事");
   ev_setsu="日程は公式サイト・SNSで要確認(例年 夏〜秋に説明会・体験入学)"; ev_bunka="山高祭＝例年9月頃・2日間・要確認";
   link="山田高校のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/3402/" },
 @{ name="愛知県立 中村高校"; short="中村高校"; area="名古屋市中村区／自由な校風（1953開校）"; accent="7B5EA7"; hensa="50"; scale="普通科(ほか美術科)"; chip="自由な校風・伝統校"; photo=($U+"nakamura.jpg"); psrc="制服＝紺セーラー服（参考写真）出典：A-制服店.com";
   rowsA=@(
     @("学科・コース","普通科(ほかに美術科)。制服は紺のセーラー服"),
     @("進学・特色","私立大学中心に幅広く進学。自主自律の伝統校(1953開校)"),
     @("最寄り駅・自宅から","地下鉄東山線「中村公園」ほか。黒川→名城線→栄→東山線で約30〜35分／バス『城北小学校』→黒川【目安】"),
     @("通学・その他","公立で学費が抑えられる。お弁当＋購買"));
   bukatsu=@("運動部だけでなく文化部も充実","陸上部・体操部が県大会出場","吹奏楽部・芸術部・メディア創造部など","『一風変わった部活』もあるのが特徴");
   kosoku=@("『自由でのびのび』との評判","ただしスカート丈・メイク・ピアスは注意指導あり","髪飾りも注意される例／学年で厳しさに差との声");
   pros=@("自由でのびのびした校風","部活に打ち込みやすい(文化部も充実)","伝統と落ち着き","公立で学費が抑えられる");
   cons=@("校舎・設備は年季あり","自由ゆえ自己管理が必要","学年で校則の運用差");
   cols=@("愛知県立 中村高校","中村区","50","文化も充実","--","〇中村区"); cmp=@("50","文化も充実","--","〇中村","自由・伝統");
   ev_setsu="体験入学(普通科) 8/8(金)午前 ※要確認(美術科 7/22)"; ev_bunka="学校祭(体育祭＋文化祭)＝例年9月";
   link="中村高校のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/123/" },
 @{ name="愛知県立 春日井南高校"; short="春日井南"; area="春日井市／文武両道・青春できる学校"; accent="2E8B57"; hensa="54"; scale="普通科(やや上)"; chip="文武両道・行事が盛ん"; photo=($U+"kasugaiminami.jpg"); psrc="制服（参考写真・エンジのタイ）出典：制服市場";
   rowsA=@(
     @("学科・コース","普通科"),
     @("進学・特色","国公立に毎年20名前後。中部大・名城大・中京大ほか。校舎がきれいで自習室が夜まで"),
     @("最寄り駅・自宅から","JR中央線「春日井」駅ほか。自宅から名鉄小牧線(上飯田)・JRで約30〜40分／バス『城北小学校』→上飯田【目安】"),
     @("通学・その他","公立で学費が抑えられる。市外だが通学圏"));
   bukatsu=@("運動部14・文化部10と豊富で打ち込める","ハンドボール部が全国大会の強豪(プロ選手も輩出)","和太鼓部が地域で演奏するなど人気","1年生は原則全員入部(文武両道の方針)");
   kosoku=@("『行事や部活に力／青春できる』校風","春陵祭・体育祭・球技大会が盛り上がる","スマホは昼放課中心に使用可(運用はゆるめ)","頭髪検査なし・服装も『やりすぎなければ』でゆるめ");
   pros=@("学習環境が良い(自習室・きれいな校舎)","評判が高い(県内上位)","部活・行事が活発","校則がゆるめ");
   cons=@("市外で通学はやや遠い","進学校で課題・勉強量は多め","1年は原則全員入部");
   cols=@("愛知県立 春日井南","春日井市","54","ハンド全国","3.9","△市外"); cmp=@("54","ハンド 全国","★3.9","△市外","環境◎");
   ev_setsu="学校説明会 8/18(火) 春日井市民会館／学校見学会 10/11(日)9:00-15:00"; ev_bunka="春陵祭＝例年秋・要確認";
   link="春日井南高校のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/52/" }
)

# ===================== PRIVATE (food) =====================
$privateSchools=@(
 @{ name="名経大 市邨高校"; short="市邨"; area="名古屋市千種区／食堂「さくら」が人気"; accent="D35400"; hensa="44〜57"; scale="コース制で幅あり"; chip="食堂「さくら」定食500円"; photo=($U+"ichimura.jpg"); psrc="制服（参考写真）出典：A-制服店.com";
   rowsA=@(
     @("学科・コース","中高一貫・探究型コース(エクスプローラー/アカデミック等)。コースで偏差値44〜57"),
     @("進学・特色","名経大ほか私大中心。五輪体操選手(黒田真由・寺本明日香ら)を輩出"),
     @("最寄り駅・自宅から","地下鉄「今池/池下」(千種区)。黒川→名城線→栄→東山線で約30分／バス『城北小学校』→黒川【目安】"),
     @("食堂・昼食","ランチルーム「さくら」。日替り定食500円・ランチ450円。人気で保護者も利用可"));
   bukatsu=@("部活の満足度が高い(口コミ評価◎)","バドミントンが全国級","体操で五輪代表を多数輩出した伝統","運動・文化とも幅広く活動");
   kosoku=@("『伸び伸び・主体的に学べる』校風","スマホは校内使用禁止(実際は守る人少との声)","化粧・ピアス・スカート丈・髪色は厳しめ","制服と分かるSNS投稿→即指導／衣替え自由");
   pros=@("食堂「さくら」が人気・安い(定食500円)","探究型でコースを選べる","バド全国級・部活が盛ん","五輪選手を輩出した実績");
   cons=@("校則はやや厳しめ(服装・頭髪・スマホ・SNS)","最寄りから少し歩く");
   cols=@("名経大 市邨","千種区","44-57","さくら食堂","3.2","◎千種"); cmp=@("44-57","さくら◎","★3.2","◎千種","探究・バド");
   ev_setsu="説明会・オープンスクールは公式『高校イベント一覧』で要確認"; ev_bunka="市邨祭(文化祭)＝例年10月・要確認";
   link="市邨高校のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/3385/" },
 @{ name="名経大 高蔵高校"; short="高蔵"; area="名古屋市瑞穂区／「メリア食堂」・制服の評判◎"; accent="1F6F8C"; hensa="49〜55"; scale="特進55/進学49/商業43"; chip="「メリア食堂」充実"; photo=($U+"takakura.jpg"); psrc="制服（参考写真・花文字校章）出典：A-制服店.com";
   rowsA=@(
     @("学科・コース","普通科(特進55/進学49)＋商業科43。中高一貫"),
     @("進学・特色","名経大ほか私大中心。制服の評判が高い(口コミ4.0)。のびのび教育"),
     @("最寄り駅・自宅から","地下鉄「瑞穂運動場西」(名城線)ほか。黒川→名城線でほぼ一本約25〜30分／バス『城北小学校』→黒川【目安】"),
     @("食堂・昼食","「メリア食堂」という広い食堂。うどん・カツカレー・唐揚げなど充実"));
   bukatsu=@("運動系・文化系あわせて約40の部","サッカー部・陸上部が活発(陸上は充実との声)","バスケ部から有名人(岡田麻央)も","文化部が多めの構成");
   kosoku=@("のびのびした教育方針","スマホは放課中も使用禁止(違反は朝指導で1週間預け)","メイク・髪に厳しい／頭髪検査は事前告知あり","アルバイトは申請で許可／登下校の指定ルートあり");
   pros=@("「メリア食堂」が広く充実","制服の評判が高い(口コミ4.0)","特進〜商業まで幅広いコース","部活の種類が多い(約40)");
   cons=@("スマホ・メイク等の校則は厳しめ","コースで偏差値差が大きい","中高一貫で内進生もいる");
   cols=@("名経大 高蔵","瑞穂区","49-55","メリア食堂","-","◎瑞穂"); cmp=@("49-55","メリア◎","★-","◎瑞穂","制服◎・商業");
   ev_setsu="学校説明会 7/21(祝)・10/25(土)・11/29(土) 9:00-12:00(WEB予約)"; ev_bunka="文化祭＝例年10月／体育祭9月";
   link="高蔵高校のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/3386/" },
 @{ name="愛工大 名電高校"; short="名電"; area="名古屋市千種区(砂田橋)／黒川から名城線一本"; accent="8E44AD"; hensa="45〜65"; scale="特進65〜スポーツ45"; chip="校内ランチルーム"; photo=($U+"meiden.jpg"); psrc="制服（参考写真）出典：A-制服店.com";
   rowsA=@(
     @("学科・コース","普通科(特進65/選抜64/普通59/スポーツ45)＋情報科"),
     @("進学・特色","愛知工大ほか。イチロー・浅田真央らの母校。スポーツ・ものづくりに強い"),
     @("最寄り駅・自宅から","地下鉄名城線「砂田橋」/東山線「池下」。黒川→名城線で約15〜20分(ほぼ一本)／バス『城北小学校』→黒川【目安】"),
     @("食堂・昼食","校内にランチルーム(昼食提供)。購買も充実(メロンパン等)。野球部寮に専用食堂"));
   bukatsu=@("全国屈指の部活強豪校","野球部＝甲子園常連(春の選抜優勝歴)","バスケ部＝全国優勝歴／バレー・吹奏楽も全国","スポーツコースもあり本格的");
   kosoku=@("『部活を頑張る人に最適』『真面目で平和』","スマホは特に厳しく禁止／スカート膝下","男子の長髪禁止・月1頭髪検査／メイク厳格","アルバイト禁止・週1見回り。『守れば何も言われない』");
   pros=@("黒川から名城線一本で通いやすい","校内ランチルーム＋購買が充実","全国級の部活(野球・バスケ等)","コースが幅広く選べる");
   cons=@("校則がかなり厳しい(スマホ・バイト・頭髪)","普通コースは偏差値が高め(59)");
   cols=@("愛工大 名電","千種区","45-65","ランチ室","3.3","◎砂田橋"); cmp=@("45-65","ランチ室","★3.3","◎砂田橋","部活強豪");
   ev_setsu="学校説明会・体験入学は公式サイトで要確認(夏〜秋)"; ev_bunka="名電祭(文化祭)＝例年秋・要確認";
   link="名電のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/158/" },
 @{ name="東邦高校"; short="東邦"; area="名古屋市名東区／大学の学食＋キッチンカー"; accent="C0392B"; hensa="52"; scale="普通52/文理特進61"; chip="大学学食＋キッチンカー"; photo=($U+"toho.jpg"); psrc="制服（参考写真・青紺ブレザー）出典：A-制服店.com";
   rowsA=@(
     @("学科・コース","普通科(普通52/文理特進61/人間健康)＋美術科・世界探究科"),
     @("進学・特色","愛知東邦大ほか。野球の名門『春の東邦』。学習重視の校風"),
     @("最寄り駅・自宅から","名東区平和が丘。地下鉄東山線「一社」＋バス。黒川→名城線→栄→東山線で約40〜45分【目安】"),
     @("食堂・昼食","専用食堂なし。昼はキッチンカー＋同キャンパスの愛知東邦大の学食(日替り440円・丼390円・カレー280円)を利用可"));
   bukatsu=@("野球部が全国屈指の名門(『春の東邦』)","近年はサッカーにも注力","文化部・運動部とも盛んに活動");
   kosoku=@("『勉強に集中できる環境』を重視","化粧・染髪・校内スマホは禁止で『ほんとに厳しい』","朝礼時に門前指導・違反は反省文","『キラキラJK生活はない』との声");
   pros=@("愛知東邦大の学食(440円〜)＋キッチンカー","世界探究科など特色コース／美術科もあり","野球の名門で部活が盛ん","学習に集中しやすい環境");
   cons=@("専用学生食堂は無い(大学学食/車を利用)","校則が厳しい(化粧・染髪・スマホ)","名東区東部で北区からやや遠い");
   cols=@("東邦","名東区","52","大学学食","2.9","○名東"); cmp=@("52","大学学食","★2.9","○名東","世界探究");
   ev_setsu="学校説明会 7/28-29・10/4(土)・10/11(土)／文理特進説明会 8/29(土)"; ev_bunka="東邦祭(文化祭)＝例年9月・要確認";
   link="東邦のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/3001/" },
 @{ name="東海学園高校"; short="東海"; area="名古屋市天白区／大学隣接の学生食堂"; accent="2E8B57"; hensa="48〜55"; scale="飛翔55/留学54/明照48"; chip="大学隣接の学生食堂"; photo=($U+"tokaigakuen.jpg"); psrc="制服（参考写真・指定バッグ付）出典：A-制服店.com";
   rowsA=@(
     @("学科・コース","普通科(飛翔55/留学54/明照48)。仏教系・中高一貫"),
     @("進学・特色","東海学園大ほか。留学コースなど国際・進学。『共生(ともいき)』の精神"),
     @("最寄り駅・自宅から","天白区中平。地下鉄鶴舞線「原」徒歩12分/市バス「平針新屋敷」。黒川→名城線→上前津→鶴舞線で約40〜45分【目安】"),
     @("食堂・昼食","同一法人の大学キャンパスに隣接し、学生食堂(麺類・丼物)を利用可。明るい雰囲気"));
   bukatsu=@("運動:サッカー・バスケ・なぎなた・バレー・テニス等","サッカー部・なぎなた部が全国大会出場経験","文化:吹奏楽・演劇・メディアプロダクト・茶道ほか多数","メカトロ部・ダンス同好会も全国経験");
   kosoku=@("『共生』の精神・支え合う校風(個性的な仲間)","スマホはロッカー保管指導(実際は所持多・抜き打ち確認)","服装・頭髪は比較的ゆるめ(巻き髪・メイクの子も)","アルバイトは原則不可だが実際はやる人も");
   pros=@("大学隣接の学生食堂(麺・丼)が使える","留学コースなど進学/国際","部活が多彩(全国経験の部も)","落ち着いた共生の校風");
   cons=@("「原」駅から徒歩約12分","北区からは通学やや遠い","スマホ等は実態とルールに差");
   cols=@("東海学園","天白区","48-55","大学学食","-","△天白"); cmp=@("48-55","大学学食","★-","△天白","留学・進学");
   ev_setsu="オープンスクール(申込〜8/14)／学校説明会 9月〜(申込7/6〜)"; ev_bunka="記念祭(文化祭)＝公式で要確認";
   link="東海学園のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/2759/" }
)

$pubConds=@(@("1","公立","私立ではなく公立校"),@("2","共学","男女いっしょに学ぶ"),@("3","偏差値 50前後","学力・内申の目安"),@("4","部活・行事が活発","打ち込める部活・楽しい行事"),@("5","評判・特色がいい","通ってよかったと思える"),@("6","通いやすい","自宅(北区)から通えること"))
$priConds=@(@("1","私立・共学","私立の男女共学校"),@("2","偏差値 50前後","学力・内申の目安"),@("3","食堂・学食がある","校内や系列大で温かい昼食"),@("4","部活・行事が活発","打ち込める部活・楽しい行事"),@("5","評判・特色がいい","通ってよかったと思える"),@("6","通いやすい","自宅(北区)から通えること"))

$pubSteps=@(@("1","説明会・学校見学会に行く","この資料の2026日程を参考に予約。親子で雰囲気を確認"),@("2","部活・文化祭を見に行く","文化祭は学校の素顔が見える。部活の様子も確認"),@("3","通学を実際に試す","自宅(北区)から朝に。黒川・上飯田・バス(城北小学校)の乗換を体感"),@("4","倍率・募集要項を公式で確認","この資料の数値・日程は目安。最新は必ず各校公式で"))
$priSteps=@(@("1","説明会・オープンキャンパスに行く","この資料の日程を参考に予約。雰囲気・先生・生徒を確認"),@("2","文化祭・部活・食堂を見に行く","学校の素顔が見える。食堂はできれば実食して確かめる"),@("3","通学を実際に試す","自宅(北区)から朝に。黒川・上飯田・バス(城北小学校)の乗換を体感"),@("4","学費・特待・推薦を公式で確認","私立は費用差が大きい。特待生・授業料補助・推薦基準を要チェック"))

Build-Deck ($base+"高校さがし_公立編_五郎.pptx") "紗衣の 高校さがし ＜公立編・くわしく版＞" "公立の候補4校（北区から通えるエリア）" $publicSchools $false $pubConds @("学校","区","偏差値","部活","評判","通学") $false @("偏差値","部活・実績","評判","通学","特色") $pubSteps "五郎より：まずは地元『北高校』(8/4説明会・10/10見学会)から。気になる順に一緒に動きましょう。応援しています。"

Build-Deck ($base+"高校さがし_私立編_五郎.pptx") "紗衣の 高校さがし ＜私立・食堂編・くわしく版＞" "食堂が使える私立・共学5校（北区から通えるエリア）" $privateSchools $true $priConds @("学校","区","偏差値","食堂・昼食","評判","通学") $true @("偏差値","食堂・昼食","評判","通学","特色") $priSteps "五郎より：まずは通いやすい『市邨』『高蔵』『名電』へ。食堂の実食もおすすめ。費用面(特待・補助)も一緒に確認しましょう。"

$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
[GC]::Collect()
Write-Output "DONE"
Get-ChildItem ($base+"高校さがし_*編_五郎.pptx") | ForEach-Object { "{0}  {1} KB" -f $_.Name, [math]::Round($_.Length/1KB) }
