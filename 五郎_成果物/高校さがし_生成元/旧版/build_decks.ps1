# 高校さがしレポート（公立編・私立編）generator - PowerPoint COM
$ErrorActionPreference = "Stop"

function CRGB([string]$hex){
  $r=[Convert]::ToInt32($hex.Substring(0,2),16)
  $g=[Convert]::ToInt32($hex.Substring(2,2),16)
  $b=[Convert]::ToInt32($hex.Substring(4,2),16)
  return ($r + $g*256 + $b*65536)
}

$NAVY="1E2761"; $BG="F4F6FB"; $INK="222B45"; $MUTED="6B7280"
$WHITE="FFFFFF"; $CORAL="F96167"; $ICE="CADCFC"; $LINE="D8DEEC"
$GREENBG="E7F6EC"; $GREENINK="1E7A43"; $REDBG="FBEAEA"; $REDINK="B23B3B"
$FONT="Meiryo"

$msoShapeRect=1; $msoShapeRound=5; $msoShapeOval=9
$msoTrue=-1; $msoFalse=0
$alnL=1; $alnC=2
$anTop=1; $anMid=3

$ppt = New-Object -ComObject PowerPoint.Application

function Set-Bg($s,$hex){
  $s.FollowMasterBackground = $msoFalse
  $s.Background.Fill.Solid()
  $s.Background.Fill.ForeColor.RGB = [int](CRGB $hex)
}
function Add-Rect($s,$type,$l,$t,$w,$h,$fillHex,$lineHex,$lineW){
  $sp = $s.Shapes.AddShape($type,$l,$t,$w,$h)
  if($fillHex){ $sp.Fill.Solid(); $sp.Fill.ForeColor.RGB=[int](CRGB $fillHex) } else { $sp.Fill.Visible=$msoFalse }
  if($lineHex){ $sp.Line.Visible=$msoTrue; $sp.Line.ForeColor.RGB=[int](CRGB $lineHex); $sp.Line.Weight=[single]$lineW } else { $sp.Line.Visible=$msoFalse }
  $sp.Shadow.Visible=$msoFalse
  return $sp
}
function Add-Text($s,$l,$t,$w,$h,$text,$size,$colorHex,$bold,$align,$anchor){
  $sp = $s.Shapes.AddTextbox(1,$l,$t,$w,$h)
  $tf = $sp.TextFrame
  $tf.WordWrap=$msoTrue
  $tf.MarginLeft=3; $tf.MarginRight=3; $tf.MarginTop=1; $tf.MarginBottom=1
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
function Add-Para($s,$l,$t,$w,$h,$lines,$size,$colorHex){
  $txt = ($lines | ForEach-Object { "・" + $_ }) -join "`r"
  $sp = Add-Text $s $l $t $w $h $txt $size $colorHex $false $alnL $anTop
  return $sp
}
function Add-Link($s,$l,$t,$w,$h,$text,$size,$colorHex,$url){
  $sp = Add-Text $s $l $t $w $h $text $size $colorHex $false $alnL $anMid
  $sp.TextFrame.TextRange.ActionSettings.Item(1).Hyperlink.Address = $url
  return $sp
}

function Build-Deck($outPath,$title,$sub,$schools,$isPrivate){
  $pres = $ppt.Presentations.Add(0)
  $pres.PageSetup.SlideWidth = 960
  $pres.PageSetup.SlideHeight = 540

  # ---- SLIDE 1 : COVER ----
  $s = $pres.Slides.Add(1,12)
  Set-Bg $s $NAVY
  Add-Rect $s $msoShapeOval 760 -120 420 420 "2A3A78" $null 0 | Out-Null
  Add-Rect $s $msoShapeOval -140 360 360 360 "2A3A78" $null 0 | Out-Null
  Add-Rect $s $msoShapeOval 60 64 64 64 $CORAL $null 0 | Out-Null
  Add-Text $s 60 60 600 40 "五郎レポート" 16 $ICE $false $alnL $anTop | Out-Null
  Add-Text $s 60 165 860 120 $title 52 $WHITE $true $alnL $anTop | Out-Null
  Add-Text $s 64 300 860 50 $sub 19 $ICE $false $alnL $anTop | Out-Null
  Add-Text $s 60 470 860 40 "作成：五郎（AI秘書）　2026年6月　自宅PC（名古屋市北区鳩岡）にて" 13 "9DB0E8" $false $alnL $anTop | Out-Null

  # ---- SLIDE 2 : CONDITIONS ----
  $s = $pres.Slides.Add(2,12)
  Set-Bg $s $BG
  Add-Text $s 40 30 880 50 "さがした条件（6つ ＋ エリア）" 30 $NAVY $true $alnL $anTop | Out-Null
  $conds = @(
   @("1","公立 or 私立","この資料の対象タイプ"),
   @("2","共学","男女いっしょに学ぶ学校"),
   @("3","偏差値 50前後","学力・内申の目安ライン"),
   @("4","バドミントン部","やりたい部活があること"),
   @("5","行事・部活が活発","楽しい高校生活を送れる"),
   @("6","評判・特色がいい","通ってよかったと思える")
  )
  $cardW=280; $cardH=112; $gx=20; $gy=18; $x0=40; $y0=92
  for($i=0;$i -lt 6;$i++){
    $col=$i%3; $row=[int]([math]::Floor($i/3))
    $l=$x0+$col*($cardW+$gx); $t=$y0+$row*($cardH+$gy)
    Add-Rect $s $msoShapeRound $l $t $cardW $cardH $WHITE $LINE 1 | Out-Null
    Add-Rect $s $msoShapeOval ($l+18) ($t+18) 42 42 $CORAL $null 0 | Out-Null
    Add-Text $s ($l+18) ($t+18) 42 42 $conds[$i][0] 21 $WHITE $true $alnC $anMid | Out-Null
    Add-Text $s ($l+72) ($t+18) ($cardW-86) 40 $conds[$i][1] 18 $NAVY $true $alnL $anMid | Out-Null
    Add-Text $s ($l+20) ($t+66) ($cardW-36) 34 $conds[$i][2] 13 $INK $false $alnL $anTop | Out-Null
  }
  Add-Rect $s $msoShapeRound 40 372 880 66 "E9EEFA" $null 0 | Out-Null
  Add-Text $s 60 372 850 66 "エリア：自宅＝名古屋市北区鳩岡（最寄り 名鉄小牧線「味鋺」駅／JR城北線「味美」）から通えること。" 15 $NAVY $true $alnL $anMid | Out-Null
  Add-Text $s 40 452 880 60 "※偏差値・評判・所要時間は2026年6月時点の進学情報サイト等を参照した目安です。倍率・学費・最新情報は必ず各校の公式サイトと経路アプリでご確認ください。" 12 $MUTED $false $alnL $anTop | Out-Null

  # ---- SLIDE 3 : CANDIDATES TABLE ----
  $s = $pres.Slides.Add(3,12)
  Set-Bg $s $BG
  $tt = if($isPrivate){"私立の候補4校（バドミントン部あり・共学）"}else{"公立の候補4校（北区から通えるエリア）"}
  Add-Text $s 40 30 880 50 $tt 28 $NAVY $true $alnL $anTop | Out-Null
  $cols = @("学校","区・市","偏差値","バド部","評判","通学")
  $cw = @(250,140,130,110,100,150)
  $tx=40; $ty=96; $rh=66; $hh=42
  $cxp=$tx
  for($c=0;$c -lt 6;$c++){
    Add-Rect $s $msoShapeRect $cxp $ty $cw[$c] $hh $NAVY $null 0 | Out-Null
    Add-Text $s $cxp $ty $cw[$c] $hh $cols[$c] 15 $WHITE $true $alnC $anMid | Out-Null
    $cxp += $cw[$c]
  }
  for($r=0;$r -lt 4;$r++){
    $rt = $ty+$hh+$r*$rh
    $rowbg = if($r%2 -eq 0){$WHITE}else{"EEF2FB"}
    $cxp=$tx
    $row = $schools[$r].cols
    for($c=0;$c -lt 6;$c++){
      Add-Rect $s $msoShapeRect $cxp $rt $cw[$c] $rh $rowbg $LINE 0.75 | Out-Null
      $val=$row[$c]
      if($c -eq 0){
        Add-Rect $s $msoShapeOval ($cxp+10) ($rt+$rh/2-7) 14 14 $schools[$r].accent $null 0 | Out-Null
        Add-Text $s ($cxp+30) $rt ($cw[$c]-34) $rh $val 13 $NAVY $true $alnL $anMid | Out-Null
      } elseif($c -eq 2){
        Add-Text $s $cxp $rt $cw[$c] $rh $val 16 $schools[$r].accent $true $alnC $anMid | Out-Null
      } else {
        Add-Text $s $cxp $rt $cw[$c] $rh $val 13 $INK $false $alnC $anMid | Out-Null
      }
      $cxp += $cw[$c]
    }
  }
  Add-Text $s 40 430 880 80 "※「評判」はみんなの高校情報の総合評価（5点満点）。偏差値はコース制の学校はコース幅で表記。所要時間は自宅(北区鳩岡)からの目安です。" 12 $MUTED $false $alnL $anTop | Out-Null

  # ---- SLIDE 4.. : SCHOOL DETAIL ----
  foreach($sc in $schools){
    $s = $pres.Slides.Add($pres.Slides.Count+1,12)
    Set-Bg $s $WHITE
    $acc = $sc.accent
    Add-Rect $s $msoShapeRound 40 26 624 50 $acc $null 0 | Out-Null
    Add-Text $s 58 26 600 50 $sc.name 23 $WHITE $true $alnL $anMid | Out-Null
    Add-Text $s 44 80 624 24 $sc.area 13 $MUTED $false $alnL $anTop | Out-Null
    # hensa callout
    Add-Rect $s $msoShapeRound 686 26 234 78 "F2F5FC" $acc 1.25 | Out-Null
    Add-Text $s 686 32 234 20 "偏差値（目安）" 12 $MUTED $false $alnC $anTop | Out-Null
    Add-Text $s 686 46 150 44 $sc.hensa 30 $acc $true $alnC $anTop | Out-Null
    Add-Text $s 686 86 234 16 $sc.scale 10 $MUTED $false $alnC $anTop | Out-Null
    # bad chip
    Add-Rect $s $msoShapeRound 686 110 234 30 $acc $null 0 | Out-Null
    Add-Text $s 686 110 234 30 ("バドミントン部：" + $sc.bad) 13 $WHITE $true $alnC $anMid | Out-Null
    # info rows (left)
    $ix=40; $iy=110; $rh=50; $labW=128; $valW=486
    for($i=0;$i -lt $sc.rows.Count;$i++){
      $t=$iy+$i*$rh
      Add-Rect $s $msoShapeRound $ix $t ($labW+$valW+12) ($rh-7) "F7F9FD" $null 0 | Out-Null
      Add-Text $s ($ix+12) $t $labW ($rh-7) $sc.rows[$i][0] 13 $acc $true $alnL $anMid | Out-Null
      Add-Text $s ($ix+12+$labW) $t $valW ($rh-7) $sc.rows[$i][1] 12 $INK $false $alnL $anMid | Out-Null
    }
    # pros / cons
    $py=322; $ph=150
    Add-Rect $s $msoShapeRound 40 $py 430 $ph $GREENBG $null 0 | Out-Null
    Add-Text $s 56 ($py+8) 400 24 "■ 口コミの良い点" 14 $GREENINK $true $alnL $anTop | Out-Null
    Add-Para $s 56 ($py+36) 404 ($ph-44) $sc.pros 12 $INK | Out-Null
    Add-Rect $s $msoShapeRound 490 $py 430 $ph $REDBG $null 0 | Out-Null
    Add-Text $s 506 ($py+8) 400 24 "■ 気になる点" 14 $REDINK $true $alnL $anTop | Out-Null
    Add-Para $s 506 ($py+36) 404 ($ph-44) $sc.cons 12 $INK | Out-Null
    # link
    Add-Rect $s $msoShapeRound 40 486 880 34 "EEF2FB" $null 0 | Out-Null
    Add-Link $s 54 486 866 34 ("▶ " + $sc.link) 12 "1A56C4" $sc.url | Out-Null
  }

  # ---- COMPARISON ----
  $s = $pres.Slides.Add($pres.Slides.Count+1,12)
  Set-Bg $s $BG
  Add-Text $s 40 30 880 50 "4校をくらべると（早見表）" 30 $NAVY $true $alnL $anTop | Out-Null
  $crit = @("偏差値","バドミントン部","評判","通学のしやすさ","特色・ポイント")
  $lx=40; $ly=96; $lw=180; $colw=170; $rh=62; $hh=44
  Add-Rect $s $msoShapeRect $lx $ly $lw $hh $NAVY $null 0 | Out-Null
  Add-Text $s $lx $ly $lw $hh "項目" 14 $WHITE $true $alnC $anMid | Out-Null
  for($c=0;$c -lt 4;$c++){
    $cxp=$lx+$lw+$c*$colw
    Add-Rect $s $msoShapeRect $cxp $ly $colw $hh $schools[$c].accent $null 0 | Out-Null
    Add-Text $s $cxp $ly $colw $hh $schools[$c].short 14 $WHITE $true $alnC $anMid | Out-Null
  }
  for($r=0;$r -lt 5;$r++){
    $rt=$ly+$hh+$r*$rh
    Add-Rect $s $msoShapeRect $lx $rt $lw $rh "E4E9F5" $LINE 0.75 | Out-Null
    Add-Text $s ($lx+10) $rt ($lw-16) $rh $crit[$r] 13 $NAVY $true $alnL $anMid | Out-Null
    for($c=0;$c -lt 4;$c++){
      $cxp=$lx+$lw+$c*$colw
      $rowbg = if($r%2 -eq 0){$WHITE}else{"EEF2FB"}
      Add-Rect $s $msoShapeRect $cxp $rt $colw $rh $rowbg $LINE 0.75 | Out-Null
      $val=$schools[$c].cmp[$r]
      if($r -eq 0){ Add-Text $s $cxp $rt $colw $rh $val 16 $schools[$c].accent $true $alnC $anMid | Out-Null }
      else { Add-Text $s ($cxp+6) $rt ($colw-12) $rh $val 12 $INK $false $alnC $anMid | Out-Null }
    }
  }

  # ---- NEXT STEPS ----
  $s = $pres.Slides.Add($pres.Slides.Count+1,12)
  Set-Bg $s $NAVY
  Add-Text $s 40 36 880 50 "次にやること（おすすめの進め方）" 30 $WHITE $true $alnL $anTop | Out-Null
  if($isPrivate){
    $steps = @(
     @("1","説明会・オープンキャンパスに行く","雰囲気・先生・生徒の様子を親子で確認する"),
     @("2","バドミントン部を見学する","強さ・人数・方針（全国志向か楽しむ系か）を確認"),
     @("3","通学を実際に試す","自宅(北区)から朝の時間帯に。スクールバスの有無も確認"),
     @("4","学費・特待・推薦の条件を公式で確認","私立は費用差が大きい。特待生・授業料補助も要チェック")
    )
    $goro = "五郎より：まずは市邨・名城大附属あたりのオープンキャンパスへ。バド部の見学を軸に、費用面も一緒に確認しましょう。"
  } else {
    $steps = @(
     @("1","学校説明会・体験入学に行く","パンフより『行ってみた感じ』が大事。親子で雰囲気を確認"),
     @("2","バドミントン部を見学する","部の強さ・人数・雰囲気を見る。入りたい部かを確認"),
     @("3","通学を実際に試す","自宅(北区)から朝の時間帯に1往復。乗換・所要時間を体感"),
     @("4","倍率・募集要項を公式で確認","この資料の数値は目安。最新の倍率・日程は必ず公式サイトで")
    )
    $goro = "五郎より：まずは『北高校』と『山田高校』の説明会から。気になる順に一緒に動きましょう。応援しています。"
  }
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
  $pres.SaveAs($outPath, 24)
  $pres.Close()
}

# ======== DATA : PUBLIC ========
$publicSchools = @(
 @{ name="名古屋市立 北高校（きた）"; short="北高校"; area="名古屋市北区如来町／自宅と同じ北区＝最短"; accent="0E8C8B"; hensa="50"; scale="市内・公立で中堅"; bad="あり";
    rows=@(
      @("学科・コース","普通科（特進コース／国際理解コース）"),
      @("特色","音楽部が全国レベル。国際理解コースは留学・実践英語が充実"),
      @("最寄り駅・自宅から","JR城北線「比良」駅＋市バス。自宅(北区鳩岡)から自転車約15分、市内同区で最短【目安】"),
      @("進学・通学","大学進学約7割。地元国公立(愛知教育大・名市大・県立大)も／通学◎")
    );
    pros=@("地元で通いやすい(同じ北区)","国際理解コースで英語・国際交流","落ち着いて過ごせる環境");
    cons=@("課題・提出物が多めとの声","駅から遠くバス・自転車が前提");
    cols=@("名古屋市立 北高校","北区","50","あり","2.9","◎北区");
    cmp=@("50","あり","★2.9","◎北区","地元で最短");
    link="北高校のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/3391/" },
 @{ name="名古屋市立 山田高校（やまだ）"; short="山田高校"; area="名古屋市西区二方町／最寄り 上小田井"; accent="E8633F"; hensa="49"; scale="市内・公立で中堅"; bad="◎人気";
    rows=@(
      @("学科・コース","普通科"),
      @("特色","明るく楽しい校風。ダンス部が全国常連、行事・部活が活発。制服も人気"),
      @("最寄り駅・自宅から","地下鉄鶴舞線・名鉄犬山線「上小田井」駅。自宅から味鋺→上飯田→乗換で約40〜50分【目安】"),
      @("進学・通学","同志社・立命館・関西・愛知工大など私大中心／通学〇")
    );
    pros=@("とにかく楽しい・明るい校風","部活・行事が盛ん","制服がかわいいと人気");
    cons=@("勉強面の手厚さは控えめとの声","校則は服装に厳しめ");
    cols=@("名古屋市立 山田高校","西区","49","◎人気","3.5","〇上小田井");
    cmp=@("49","◎人気","★3.5","〇西区","部活・評判◎");
    link="山田高校のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/3402/" },
 @{ name="愛知県立 中村高校（なかむら）"; short="中村高校"; area="名古屋市中村区／自由な校風（1953開校）"; accent="7B5EA7"; hensa="50"; scale="市内・公立で中堅"; bad="県大会";
    rows=@(
      @("学科・コース","普通科"),
      @("特色","伝統校で自主自律。弓道・バドミントン(県大会)・陸上・体操など部活が熱心"),
      @("最寄り駅・自宅から","地下鉄東山線「中村公園」ほか。自宅から平安通→名城線→東山線で約40分【目安】"),
      @("進学・通学","私立大学中心に幅広く進学／通学〇")
    );
    pros=@("自由でのびのびした校風","部活に打ち込みやすい","伝統と落ち着きがある");
    cons=@("校舎・設備は年季あり","自由ゆえ自己管理が必要");
    cols=@("愛知県立 中村高校","中村区","50","県大会","--","〇中村区");
    cmp=@("50","県大会","--","〇中村","自由な校風");
    link="中村高校のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/123/" },
 @{ name="愛知県立 春日井南高校（チャレンジ校）"; short="春日井南"; area="春日井市／少し上のチャレンジ校"; accent="2E8B57"; hensa="54"; scale="やや上のレベル"; bad="要確認";
    rows=@(
      @("学科・コース","普通科"),
      @("特色","校舎がきれいで自習室が夜まで。運動部14・文化部10と部活豊富(ハンドボール全国級)"),
      @("最寄り駅・自宅から","JR中央線「春日井」駅ほか。自宅から味美→勝川→春日井で約30〜40分【目安】"),
      @("進学・通学","国公立に毎年20名前後。中部大・名城大・中京大ほか／通学△〜〇")
    );
    pros=@("学習環境が良い(自習室・きれいな校舎)","評判が高い(県内上位)","部活も活発");
    cons=@("市外で通学はやや遠い","進学校で課題・勉強量は多め");
    cols=@("愛知県立 春日井南","春日井市","54","要確認","3.9","△市外");
    cmp=@("54","要確認","★3.9","△市外","環境◎・上狙い");
    link="春日井南高校のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/52/" }
)

# ======== DATA : PRIVATE ========
$privateSchools = @(
 @{ name="名経大 市邨高校（いちむら）"; short="市邨"; area="名古屋市千種区／バドミントン全国級"; accent="D35400"; hensa="44〜57"; scale="コース制で幅あり"; bad="◎全国級";
    rows=@(
      @("学科・コース","中高一貫・探究型コース(エクスプローラー/アカデミック等)。偏差値44〜57"),
      @("特色","バドミントン部が全国レベルで選手も輩出。部活に本気で打ち込める環境"),
      @("最寄り駅・自宅から","地下鉄「今池/池下」方面(千種区)。自宅から平安通→名城線で約35〜45分【目安】"),
      @("進学・通学","名経大ほか私大中心、コースで幅広い進路／通学〇")
    );
    pros=@("バド部が全国級・本気で打ち込める","コースが選べる(探究型)","部活に強い校風");
    cons=@("校則はやや厳しい(服装・頭髪・スマホ)","最寄りから少し歩く");
    cols=@("名経大 市邨","千種区","44-57","◎全国","3.2","〇千種");
    cmp=@("44-57","◎全国","★3.2","〇千種","バド強豪");
    link="市邨高校のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/3385/" },
 @{ name="名城大学附属高校（めいじょう）"; short="名城大附属"; area="名古屋市中村区／名城大への附属推薦"; accent="1F6F8C"; hensa="56〜"; scale="コースで56〜66"; bad="あり";
    rows=@(
      @("学科・コース","普通科5コース(総合/国際56〜進学61〜特進66)"),
      @("特色","名城大への附属推薦で進学が安心。23の運動部が盛ん(バド部あり)。施設充実"),
      @("最寄り駅・自宅から","名鉄「東枇杷島」駅 徒歩約10分。自宅からJR味美→枇杷島乗換で約35〜45分【目安】"),
      @("進学・通学","名城大進学が最多、国公立・難関私大も／通学〇")
    );
    pros=@("附属推薦で大学進学が安心","部活が盛ん・施設が良い","ある程度自由(メイク等)");
    cons=@("先生により対応に差との声","規模が大きく評価は分かれる");
    cols=@("名城大附属","中村区","56-66","あり","3.3","〇中村");
    cmp=@("56〜","あり","★3.3","〇中村","附属で安心");
    link="名城大附属のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/4711/" },
 @{ name="愛知工大 名電高校（めいでん）"; short="名電"; area="名古屋市千種区／スポーツ・理系に強い"; accent="8E44AD"; hensa="45〜65"; scale="コースで幅あり"; bad="あり";
    rows=@(
      @("学科・コース","普通科(特進65/選抜64/普通59/スポーツ45)ほか"),
      @("特色","部活強豪でバド部も活動。スポーツ・ものづくりに強い(イチロー等の母校)"),
      @("最寄り駅・自宅から","地下鉄名城線「砂田橋」/東山線「池下」(千種区)。自宅から平安通→名城線でほぼ一本約30〜40分【目安】"),
      @("進学・通学","愛知工大ほか、コースで進路が幅広い／通学〇")
    );
    pros=@("部活を頑張る人に最適","コースが幅広く選べる","設備・実績が豊富");
    cons=@("校則が厳しい(スマホ・バイト禁止等)","普通コースは偏差値が高め");
    cols=@("愛知工大 名電","千種区","45-65","あり","3.3","〇千種");
    cmp=@("45〜","あり","★3.3","〇千種","部活強豪");
    link="名電のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/158/" },
 @{ name="中部大 春日丘高校（はるひ）"; short="春日丘"; area="春日井市／スクールバスで通学"; accent="27AE60"; hensa="52〜58"; scale="進学52/創進58"; bad="あり";
    rows=@(
      @("学科・コース","普通科(進学コース52／創進コース58)。中高一貫併設"),
      @("特色","校則がゆるめで部活が盛ん(バド部人気)。全国大会出場の部も多数。中部大へ系列進学"),
      @("最寄り駅・自宅から","JR中央線「神領」駅＋バス。春日井西部ルートのスクールバスあり。自宅から約40〜50分【目安】"),
      @("進学・通学","中部大に多数、名城・中京・南山も／通学△〜〇")
    );
    pros=@("校則がゆるく自由","部活が盛ん・全国出場も","先生が親身との声");
    cons=@("一貫コースは進学プレッシャー","2年で部活の区切り等ルールあり");
    cols=@("中部大 春日丘","春日井市","52-58","あり","-","△春日井");
    cmp=@("52〜58","あり","★-","△春日井","自由・部活盛ん");
    link="春日丘のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/1178/" }
)

$base = "C:\Users\山本　史朗\OneDrive\Desktop\新しいフォルダー\"
Build-Deck ($base + "高校さがし_公立編_五郎.pptx") "紗衣の 高校さがし ＜公立編＞" "名古屋市内・公立・共学／偏差値50前後・バドミントン部・評判◎" $publicSchools $false
Build-Deck ($base + "高校さがし_私立編_五郎.pptx") "紗衣の 高校さがし ＜私立編＞" "名古屋エリア・私立・共学／バドミントン部のある学校4校" $privateSchools $true

$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
[GC]::Collect()
Write-Output "DONE"
Get-ChildItem ($base + "高校さがし_*編_五郎.pptx") | Select-Object Name, Length
