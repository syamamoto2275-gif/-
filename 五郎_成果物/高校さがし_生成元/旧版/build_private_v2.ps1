# 高校さがし 私立編 v2 : 偏差値50前後・通いやすい共学4校／黒川アクセス＋制服写真＋2026イベント
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing

function CRGB([string]$hex){ $r=[Convert]::ToInt32($hex.Substring(0,2),16);$g=[Convert]::ToInt32($hex.Substring(2,2),16);$b=[Convert]::ToInt32($hex.Substring(4,2),16); return ($r+$g*256+$b*65536) }
$NAVY="1E2761"; $BG="F4F6FB"; $INK="222B45"; $MUTED="6B7280"
$WHITE="FFFFFF"; $CORAL="F96167"; $ICE="CADCFC"; $LINE="D8DEEC"
$GREENBG="E7F6EC"; $GREENINK="1E7A43"; $REDBG="FBEAEA"; $REDINK="B23B3B"
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

$base="C:\Users\山本　史朗\OneDrive\Desktop\新しいフォルダー\"; $U=$base+"uniforms\"

$schools=@(
 @{ name="名経大 市邨高校（いちむら）"; short="市邨"; area="名古屋市千種区／通いやすい中央エリア"; accent="D35400"; hensa="44〜57"; scale="コース制で幅あり"; chip="バドミントン全国級"; photo=($U+"ichimura.jpg"); psrc="制服（参考写真）出典：A-制服店.com";
    rows=@(
      @("学科・コース","中高一貫・探究型コース(エクスプローラー/アカデミック等)。偏差値44〜57"),
      @("特色","バドミントン部が全国級。探究型の学びで部活に本気で打ち込める"),
      @("最寄り駅・自宅から","地下鉄「今池/池下」(千種区)。黒川→(名城線)→栄→東山線で約30分【目安】"),
      @("進学・通学","名経大ほか私大中心、コースで幅広い進路／通学〇")
    );
    pros=@("バド部が全国級・部活に本気","コースが選べる(探究型の学び)","中高一貫で面倒見");
    cons=@("校則はやや厳しめ(服装・頭髪・スマホ)","最寄りから少し歩く");
    cols=@("名経大 市邨","千種区","44-57","あり","3.2","〇千種");
    cmp=@("44-57","全国級","★3.2","〇千種","バド・探究");
    ev_setsu="説明会・オープンスクールは公式『高校イベント一覧』で要確認";
    ev_bunka="市邨祭(文化祭)＝例年10月中旬・要確認";
    link="市邨高校のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/3385/" },
 @{ name="名古屋大谷高校（おおたに）"; short="名古屋大谷"; area="名古屋市瑞穂区／コース選択が豊富"; accent="2E8B57"; hensa="45〜52"; scale="特進52/特選47/文理45"; chip="コース選択が豊富"; photo=($U+"otani.jpg"); psrc="制服（参考写真・hers heart）出典：A-制服店.com";
    rows=@(
      @("学科・コース","普通科4コース(特進52/特選47/福祉医療45/文理45)＋商業科"),
      @("特色","コースが豊富で目標に合わせやすい。部活も活発で面倒見のよい指導"),
      @("最寄り駅・自宅から","地下鉄桜通線「瑞穂区役所」駅 徒歩4分。黒川→(名城線)→久屋大通→桜通線で約30分【目安】"),
      @("進学・通学","系列の名経大ほか私大中心／通学〇")
    );
    pros=@("コースが幅広く選べる","部活が活発","面倒見・サポートの声");
    cons=@("口コミ評価は賛否あり(2.8)","指導・校則が厳しめとの声");
    cols=@("名古屋大谷","瑞穂区","45-52","活発","2.8","〇瑞穂");
    cmp=@("45-52","活発","★2.8","〇瑞穂","コース豊富");
    ev_setsu="オープンスクール・学校見学会(要事前申込・6〜7月)";
    ev_bunka="大谷祭(文化祭)＝例年9月・2026日程は公式で要確認";
    link="名古屋大谷のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/3384/" },
 @{ name="名古屋国際高校（こくさい）"; short="名古屋国際"; area="名古屋市昭和区／国際・IB教育に強い"; accent="1F6F8C"; hensa="47〜49"; scale="普通49/国際教養47"; chip="国際・IB認定校"; photo=($U+"kokusai.jpg"); psrc="制服（参考写真・英国調）出典：制服市場";
    rows=@(
      @("学科・コース","普通科(49)／国際教養科(47)。中高一貫"),
      @("特色","東海唯一の国際バカロレア(IB)認定校。英語・国際教育に強い。制服がかわいい・校舎がきれい"),
      @("最寄り駅・自宅から","地下鉄「御器所」駅 徒歩7分。黒川→(名城線)→久屋大通→桜通線で約30〜35分【目安】"),
      @("進学・通学","名古屋商科大ほか、海外・国際系の進路も／通学〇")
    );
    pros=@("制服がかわいい・校舎がきれい","国際・英語教育が充実(IB認定校)","先生がフレンドリーとの声");
    cons=@("口コミ評価は賛否あり(2.7)","入学者の学力に幅があるとの指摘","校則あり");
    cols=@("名古屋国際","昭和区","47-49","活発","2.7","〇昭和");
    cmp=@("47-49","活発","★2.7","〇昭和","国際・IB");
    ev_setsu="オープンキャンパス＝例年6・8月ほか(要予約)／プレテスト 例年10月";
    ev_bunka="光楓祭(学校祭)＝例年9月下旬";
    link="名古屋国際のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/3388/" },
 @{ name="同朋高校（どうほう）"; short="同朋"; area="名古屋市中村区／自由な校風・名古屋駅近"; accent="8E44AD"; hensa="43〜47"; scale="普通47/商業・音楽43"; chip="自由な校風・系列大学"; photo=""; psrc="";
    rows=@(
      @("学科・コース","普通科(47/文理・美術・医療看護等の系統)＋商業科・音楽科"),
      @("特色","私服登校もできる自由な校風。系列大学(同朋・名古屋音大・名古屋造形大)と連携。柔道が強豪"),
      @("最寄り駅・自宅から","地下鉄東山線「岩塚」駅／名古屋駅も近い。黒川→(名城線)→栄→東山線で約35分【目安】"),
      @("進学・通学","系列大学への内部進学や美術・音楽・医療看護系の進路／通学〇(名駅近)")
    );
    pros=@("自由な校風(私服登校もOK)","系列大学と連携(美術・音楽・医療看護)","柔道など部活が盛ん");
    cons=@("偏差値はやや低め(43-47)","口コミ評価は賛否あり","自由ゆえ自己管理が必要");
    cols=@("同朋","中村区","43-47","自由","--","〇名駅近");
    cmp=@("43-47","自由","--","〇名駅","自由・系列大");
    ev_setsu="学校説明会(普通科) 7/25(土)・10/31(土)・11/28(土) ※公式で確認";
    ev_bunka="桜鏡祭(文化祭) 9/19(土)・20(日)";
    link="同朋高校のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/3002/" }
)

$pres=$ppt.Presentations.Add(0); $pres.PageSetup.SlideWidth=960; $pres.PageSetup.SlideHeight=540

# COVER
$s=$pres.Slides.Add(1,12); Set-Bg $s $NAVY
Add-Rect $s $msoShapeOval 760 -120 420 420 "2A3A78" $null 0 | Out-Null
Add-Rect $s $msoShapeOval -140 360 360 360 "2A3A78" $null 0 | Out-Null
Add-Rect $s $msoShapeOval 60 64 64 64 $CORAL $null 0 | Out-Null
Add-Text $s 60 60 600 40 "五郎レポート" 16 $ICE $false $alnL $anTop | Out-Null
Add-Text $s 60 165 860 120 "紗衣の 高校さがし ＜私立編＞" 46 $WHITE $true $alnL $anTop | Out-Null
Add-Text $s 64 300 860 50 "名古屋・私立・共学／偏差値50前後で通いやすい4校（制服写真・2026年度の見学日程つき）" 17 $ICE $false $alnL $anTop | Out-Null
Add-Text $s 60 470 860 40 "作成：五郎（AI秘書）　2026年6月　自宅PC（名古屋市北区鳩岡）にて" 13 "9DB0E8" $false $alnL $anTop | Out-Null

# CONDITIONS
$s=$pres.Slides.Add(2,12); Set-Bg $s $BG
Add-Text $s 40 30 880 50 "さがした条件（6つ ＋ エリア）" 30 $NAVY $true $alnL $anTop | Out-Null
$conds=@(@("1","私立","私立の共学高校"),@("2","共学","男女いっしょに学ぶ学校"),@("3","偏差値 50前後","学力・内申の目安ライン"),@("4","部活・行事が活発","打ち込める部活・楽しい行事"),@("5","評判・特色がいい","通ってよかったと思える"),@("6","通いやすい","自宅(北区)から通えること"))
$cardW=280;$cardH=112;$gx=20;$gy=18;$x0=40;$y0=92
for($i=0;$i -lt 6;$i++){
  $col=$i%3;$row=[int]([math]::Floor($i/3));$l=$x0+$col*($cardW+$gx);$t=$y0+$row*($cardH+$gy)
  Add-Rect $s $msoShapeRound $l $t $cardW $cardH $WHITE $LINE 1 | Out-Null
  Add-Rect $s $msoShapeOval ($l+18) ($t+18) 42 42 $CORAL $null 0 | Out-Null
  Add-Text $s ($l+18) ($t+18) 42 42 $conds[$i][0] 21 $WHITE $true $alnC $anMid | Out-Null
  Add-Text $s ($l+72) ($t+18) ($cardW-86) 40 $conds[$i][1] 18 $NAVY $true $alnL $anMid | Out-Null
  Add-Text $s ($l+20) ($t+66) ($cardW-36) 34 $conds[$i][2] 13 $INK $false $alnL $anTop | Out-Null
}
Add-Rect $s $msoShapeRound 40 372 880 66 "E9EEFA" $null 0 | Out-Null
Add-Text $s 60 372 850 66 "エリア：自宅＝名古屋市北区鳩岡。地下鉄の最寄り「黒川」駅(名城線)／名鉄小牧線は「上飯田」駅を起点に、中央寄りで通いやすい私立を選定。" 15 $NAVY $true $alnL $anMid | Out-Null
Add-Text $s 40 452 880 60 "※バドミントン部の条件は外し、「共学・偏差値50前後・通いやすさ・評判」で選定。私立は学費・特待・推薦の条件も要確認。" 12 $MUTED $false $alnL $anTop | Out-Null

# TABLE
$s=$pres.Slides.Add(3,12); Set-Bg $s $BG
Add-Text $s 40 30 880 50 "私立の候補4校（偏差値50前後・通いやすいエリア）" 26 $NAVY $true $alnL $anTop | Out-Null
$cols=@("学校","区","偏差値","部活","評判","通学"); $cw=@(250,140,120,120,100,150)
$tx=40;$ty=96;$rh=66;$hh=42;$cxp=$tx
for($c=0;$c -lt 6;$c++){ Add-Rect $s $msoShapeRect $cxp $ty $cw[$c] $hh $NAVY $null 0 | Out-Null; Add-Text $s $cxp $ty $cw[$c] $hh $cols[$c] 15 $WHITE $true $alnC $anMid | Out-Null; $cxp+=$cw[$c] }
for($r=0;$r -lt 4;$r++){
  $rt=$ty+$hh+$r*$rh; $rowbg=if($r%2 -eq 0){$WHITE}else{"EEF2FB"}; $cxp=$tx; $row=$schools[$r].cols
  for($c=0;$c -lt 6;$c++){
    Add-Rect $s $msoShapeRect $cxp $rt $cw[$c] $rh $rowbg $LINE 0.75 | Out-Null; $val=$row[$c]
    if($c -eq 0){ Add-Rect $s $msoShapeOval ($cxp+10) ($rt+$rh/2-7) 14 14 $schools[$r].accent $null 0 | Out-Null; Add-Text $s ($cxp+30) $rt ($cw[$c]-34) $rh $val 13 $NAVY $true $alnL $anMid | Out-Null }
    elseif($c -eq 2){ Add-Text $s $cxp $rt $cw[$c] $rh $val 16 $schools[$r].accent $true $alnC $anMid | Out-Null }
    else { Add-Text $s $cxp $rt $cw[$c] $rh $val 13 $INK $false $alnC $anMid | Out-Null }
    $cxp+=$cw[$c]
  }
}
Add-Text $s 40 430 880 80 "※「評判」はみんなの高校情報の総合評価（5点満点・私立は賛否が出やすい点に注意）。所要時間は自宅(北区鳩岡)からの目安。次ページ以降に各校の制服写真・口コミ・2026年度の日程を掲載。" 12 $MUTED $false $alnL $anTop | Out-Null

# DETAIL
foreach($sc in $schools){
  $s=$pres.Slides.Add($pres.Slides.Count+1,12); Set-Bg $s $WHITE; $acc=$sc.accent
  Add-Rect $s $msoShapeRound 40 26 624 50 $acc $null 0 | Out-Null
  Add-Text $s 58 26 600 50 $sc.name 22 $WHITE $true $alnL $anMid | Out-Null
  Add-Text $s 44 80 624 24 $sc.area 13 $MUTED $false $alnL $anTop | Out-Null
  Add-Rect $s $msoShapeRound 686 26 234 76 "F2F5FC" $acc 1.25 | Out-Null
  Add-Text $s 686 32 234 18 "偏差値（目安）" 12 $MUTED $false $alnC $anTop | Out-Null
  Add-Text $s 686 46 234 42 $sc.hensa 26 $acc $true $alnC $anTop | Out-Null
  Add-Text $s 686 84 234 16 $sc.scale 10 $MUTED $false $alnC $anTop | Out-Null
  Add-Rect $s $msoShapeRound 686 108 234 28 $acc $null 0 | Out-Null
  Add-Text $s 686 108 234 28 $sc.chip 12 $WHITE $true $alnC $anMid | Out-Null
  if($sc.photo -ne "" -and (Test-Path $sc.photo)){
    Add-Picture $s $sc.photo 686 144 234 176 | Out-Null
    Add-Text $s 686 322 234 26 $sc.psrc 9 $MUTED $false $alnC $anTop | Out-Null
  } else {
    Add-Rect $s $msoShapeRound 686 144 234 176 "F4ECFA" $acc 1 | Out-Null
    Add-Text $s 696 158 214 150 "制服は標準服あり`r（普段は私服登校もOK＝自由な校風）`r`r※指定制服の写真は省略" 13 $acc $true $alnC $anMid | Out-Null
  }
  $ix=40;$iy=110;$rh2=46;$labW=128;$valW=484
  for($i=0;$i -lt $sc.rows.Count;$i++){
    $t=$iy+$i*$rh2
    Add-Rect $s $msoShapeRound $ix $t ($labW+$valW+12) ($rh2-7) "F7F9FD" $null 0 | Out-Null
    Add-Text $s ($ix+12) $t $labW ($rh2-7) $sc.rows[$i][0] 13 $acc $true $alnL $anMid | Out-Null
    Add-Text $s ($ix+12+$labW) $t $valW ($rh2-7) $sc.rows[$i][1] 12 $INK $false $alnL $anMid | Out-Null
  }
  $py=350;$ph=130
  Add-Rect $s $msoShapeRound 40 $py 430 $ph $GREENBG $null 0 | Out-Null
  Add-Text $s 56 ($py+8) 400 22 "■ 口コミの良い点" 14 $GREENINK $true $alnL $anTop | Out-Null
  Add-Para $s 56 ($py+34) 404 ($ph-40) $sc.pros 12 $INK | Out-Null
  Add-Rect $s $msoShapeRound 490 $py 430 $ph $REDBG $null 0 | Out-Null
  Add-Text $s 506 ($py+8) 400 22 "■ 気になる点" 14 $REDINK $true $alnL $anTop | Out-Null
  Add-Para $s 506 ($py+34) 404 ($ph-40) $sc.cons 12 $INK | Out-Null
  Add-Rect $s $msoShapeRound 40 486 880 32 "EEF2FB" $null 0 | Out-Null
  Add-Link $s 54 486 866 32 ("▶ " + $sc.link) 12 "1A56C4" $sc.url | Out-Null
}

# EVENTS 2026
$s=$pres.Slides.Add($pres.Slides.Count+1,12); Set-Bg $s $BG
Add-Text $s 40 30 880 50 "2026年度（令和8年度）の オープンキャンパス・文化祭" 24 $NAVY $true $alnL $anTop | Out-Null
$ecols=@("学校","2026年度 説明会・オープンキャンパス","文化祭（学園祭）"); $ecw=@(180,460,240)
$ex=40;$ey=92;$erh=78;$ehh=40;$cxp=$ex
for($c=0;$c -lt 3;$c++){ Add-Rect $s $msoShapeRect $cxp $ey $ecw[$c] $ehh $NAVY $null 0 | Out-Null; Add-Text $s $cxp $ey $ecw[$c] $ehh $ecols[$c] 14 $WHITE $true $alnC $anMid | Out-Null; $cxp+=$ecw[$c] }
for($r=0;$r -lt 4;$r++){
  $rt=$ey+$ehh+$r*$erh; $rowbg=if($r%2 -eq 0){$WHITE}else{"EEF2FB"}
  Add-Rect $s $msoShapeRect $ex $rt $ecw[0] $erh $rowbg $LINE 0.75 | Out-Null
  Add-Rect $s $msoShapeOval ($ex+10) ($rt+$erh/2-7) 14 14 $schools[$r].accent $null 0 | Out-Null
  Add-Text $s ($ex+30) $rt ($ecw[0]-34) $erh $schools[$r].short 13 $NAVY $true $alnL $anMid | Out-Null
  Add-Rect $s $msoShapeRect ($ex+$ecw[0]) $rt $ecw[1] $erh $rowbg $LINE 0.75 | Out-Null
  Add-Text $s ($ex+$ecw[0]+10) $rt ($ecw[1]-18) $erh $schools[$r].ev_setsu 12 $INK $false $alnL $anMid | Out-Null
  Add-Rect $s $msoShapeRect ($ex+$ecw[0]+$ecw[1]) $rt $ecw[2] $erh $rowbg $LINE 0.75 | Out-Null
  Add-Text $s ($ex+$ecw[0]+$ecw[1]+10) $rt ($ecw[2]-18) $erh $schools[$r].ev_bunka 12 $INK $false $alnL $anMid | Out-Null
}
Add-Text $s 40 424 880 90 "※2026年6月時点で各校公式サイト等から確認できた情報です。日付のないものは『例年の時期・要確認』。私立はオープンキャンパスに事前予約が要る場合が多く、学費・特待・推薦の条件とあわせて必ず各校公式でご確認ください。" 12 $MUTED $false $alnL $anTop | Out-Null

# COMPARISON
$s=$pres.Slides.Add($pres.Slides.Count+1,12); Set-Bg $s $BG
Add-Text $s 40 30 880 50 "4校をくらべると（早見表）" 30 $NAVY $true $alnL $anTop | Out-Null
$crit=@("偏差値","部活・特色","評判","通学のしやすさ","特色・ポイント")
$lx=40;$ly=96;$lw=190;$colw=170;$rh3=62;$hh3=44
Add-Rect $s $msoShapeRect $lx $ly $lw $hh3 $NAVY $null 0 | Out-Null; Add-Text $s $lx $ly $lw $hh3 "項目" 14 $WHITE $true $alnC $anMid | Out-Null
for($c=0;$c -lt 4;$c++){ $cxp=$lx+$lw+$c*$colw; Add-Rect $s $msoShapeRect $cxp $ly $colw $hh3 $schools[$c].accent $null 0 | Out-Null; Add-Text $s $cxp $ly $colw $hh3 $schools[$c].short 13 $WHITE $true $alnC $anMid | Out-Null }
for($r=0;$r -lt 5;$r++){
  $rt=$ly+$hh3+$r*$rh3
  Add-Rect $s $msoShapeRect $lx $rt $lw $rh3 "E4E9F5" $LINE 0.75 | Out-Null
  Add-Text $s ($lx+10) $rt ($lw-16) $rh3 $crit[$r] 13 $NAVY $true $alnL $anMid | Out-Null
  for($c=0;$c -lt 4;$c++){
    $cxp=$lx+$lw+$c*$colw; $rowbg=if($r%2 -eq 0){$WHITE}else{"EEF2FB"}
    Add-Rect $s $msoShapeRect $cxp $rt $colw $rh3 $rowbg $LINE 0.75 | Out-Null; $val=$schools[$c].cmp[$r]
    if($r -eq 0){ Add-Text $s $cxp $rt $colw $rh3 $val 15 $schools[$c].accent $true $alnC $anMid | Out-Null }
    else { Add-Text $s ($cxp+6) $rt ($colw-12) $rh3 $val 12 $INK $false $alnC $anMid | Out-Null }
  }
}

# NEXT STEPS
$s=$pres.Slides.Add($pres.Slides.Count+1,12); Set-Bg $s $NAVY
Add-Text $s 40 36 880 50 "次にやること（おすすめの進め方）" 30 $WHITE $true $alnL $anTop | Out-Null
$steps=@(
 @("1","説明会・オープンキャンパスに行く","この資料の日程を参考に予約。雰囲気・先生・生徒を親子で確認"),
 @("2","文化祭・部活を見に行く","学校の素顔が見える。入りたい部活や校風を確認"),
 @("3","通学を実際に試す","自宅(北区)から朝の時間帯に。黒川・上飯田からの乗換を体感"),
 @("4","学費・特待・推薦の条件を公式で確認","私立は費用差が大きい。特待生・授業料補助・推薦基準を要チェック")
)
for($i=0;$i -lt 4;$i++){
  $t=104+$i*84
  Add-Rect $s $msoShapeRound 40 $t 880 72 "2A3A78" $null 0 | Out-Null
  Add-Rect $s $msoShapeOval 58 ($t+16) 40 40 $CORAL $null 0 | Out-Null
  Add-Text $s 58 ($t+16) 40 40 $steps[$i][0] 20 $WHITE $true $alnC $anMid | Out-Null
  Add-Text $s 116 ($t+11) 790 30 $steps[$i][1] 18 $WHITE $true $alnL $anTop | Out-Null
  Add-Text $s 116 ($t+41) 790 26 $steps[$i][2] 13 $ICE $false $alnL $anTop | Out-Null
}
Add-Text $s 40 466 880 50 "五郎より：まずは通いやすい『市邨』『名古屋国際』のオープンキャンパスから。費用面（特待・補助）も一緒に確認しましょう。" 14 "9DB0E8" $false $alnL $anTop | Out-Null

$out=$base+"高校さがし_私立編_五郎.pptx"
if(Test-Path $out){ Remove-Item $out -Force }
$pres.SaveAs($out,24); $pres.Close()
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
[GC]::Collect()
Write-Output ("SAVED: "+$out)
Write-Output ("SIZE: "+((Get-Item $out).Length)+" bytes")
