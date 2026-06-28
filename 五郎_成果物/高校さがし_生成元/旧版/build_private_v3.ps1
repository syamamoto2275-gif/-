# 高校さがし 私立・食堂編 v3 : 食堂/学食が使える共学5校／黒川アクセス＋制服写真＋2026イベント
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
 @{ name="名経大 市邨高校（いちむら）"; short="市邨"; area="名古屋市千種区／食堂「さくら」が人気"; accent="D35400"; hensa="44〜57"; scale="コース制で幅あり"; chip="食堂「さくら」定食500円"; photo=($U+"ichimura.jpg"); psrc="制服（参考写真）出典：A-制服店.com";
    rows=@(
      @("学科・コース","中高一貫・探究型コース(エクスプローラー等)。偏差値44〜57"),
      @("食堂・昼食","ランチルーム「さくら」。日替わり定食500円・ランチ450円。人気で保護者も利用可"),
      @("最寄り駅・自宅から","地下鉄「今池/池下」(千種区)。黒川→名城線→東山線で約30分／自宅はバス「城北小学校」→黒川【目安】"),
      @("進学・特色","名経大ほか私大中心。バドミントン部が全国級・部活が盛ん")
    );
    pros=@("食堂「さくら」が人気・安い(定食500円)","探究型でコースを選べる","バド全国級・部活が盛ん");
    cons=@("校則はやや厳しめ(服装・頭髪・スマホ)","最寄りから少し歩く");
    cols=@("名経大 市邨","千種区","44-57","さくら食堂","3.2","◎千種");
    cmp=@("44-57","さくら◎","★3.2","◎千種","探究・バド");
    ev_setsu="説明会・オープンスクールは公式『高校イベント一覧』で要確認";
    ev_bunka="市邨祭(文化祭)＝例年10月・要確認";
    link="市邨高校のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/3385/" },
 @{ name="名経大 高蔵高校（たかくら）"; short="高蔵"; area="名古屋市瑞穂区／「メリア食堂」・制服の評判◎"; accent="1F6F8C"; hensa="49〜55"; scale="特進55/進学49/商業43"; chip="「メリア食堂」充実"; photo=($U+"takakura.jpg"); psrc="制服（参考写真・花文字校章）出典：A-制服店.com";
    rows=@(
      @("学科・コース","普通科(特進55/進学49)＋商業科43。中高一貫"),
      @("食堂・昼食","「メリア食堂」という広い食堂。うどん・カツカレー・唐揚げなどメニュー充実"),
      @("最寄り駅・自宅から","地下鉄「瑞穂運動場西」(名城線)ほか。黒川→名城線でほぼ一本約25〜30分【目安】"),
      @("進学・特色","名経大ほか私大中心。制服の評判が高い(口コミ4.0)")
    );
    pros=@("「メリア食堂」が広く充実","制服の評判が高い(口コミ4.0)","特進〜商業まで幅広いコース");
    cons=@("コースで偏差値差が大きい","中高一貫で内進生もいる");
    cols=@("名経大 高蔵","瑞穂区","49-55","メリア食堂","-","◎瑞穂");
    cmp=@("49-55","メリア◎","★-","◎瑞穂","制服◎・商業");
    ev_setsu="学校説明会 7/21(祝)・10/25(土)・11/29(土) 9:00-12:00(WEB予約)";
    ev_bunka="文化祭＝例年10月／体育祭9月";
    link="高蔵高校のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/3386/" },
 @{ name="愛工大 名電高校（めいでん）"; short="名電"; area="名古屋市千種区(砂田橋)／黒川から名城線一本"; accent="8E44AD"; hensa="45〜65"; scale="特進65〜スポーツ45"; chip="校内ランチルーム"; photo=($U+"meiden.jpg"); psrc="制服（参考写真）出典：A-制服店.com";
    rows=@(
      @("学科・コース","普通科(特進65/選抜64/普通59/スポーツ45)＋情報科"),
      @("食堂・昼食","校内にランチルーム(昼食提供)。購買も充実(メロンパン等)。※野球部寮に専用食堂"),
      @("最寄り駅・自宅から","地下鉄名城線「砂田橋」/東山線「池下」。黒川→名城線で約15〜20分(ほぼ一本)【目安】"),
      @("進学・特色","愛知工大ほか。部活が強豪、コースが幅広い(イチロー等の母校)")
    );
    pros=@("黒川から名城線一本で通いやすい","校内ランチルーム＋購買が充実","部活強豪・コースが幅広い");
    cons=@("校則が厳しめ(スマホ・バイト等)","普通コースは偏差値が高め(59)");
    cols=@("愛工大 名電","千種区","45-65","ランチ室","3.3","◎砂田橋");
    cmp=@("45-65","ランチ室","★3.3","◎砂田橋","部活強豪");
    ev_setsu="学校説明会・体験入学は公式サイトで要確認(夏〜秋)";
    ev_bunka="名電祭(文化祭)＝例年秋・要確認";
    link="名電のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/158/" },
 @{ name="東邦高校（とうほう）"; short="東邦"; area="名古屋市名東区／大学の学食＋キッチンカー"; accent="C0392B"; hensa="52"; scale="普通52/文理特進61"; chip="大学学食＋キッチンカー"; photo=($U+"toho.jpg"); psrc="制服（参考写真・青紺ブレザー）出典：A-制服店.com";
    rows=@(
      @("学科・コース","普通科(普通52/文理特進61/人間健康)＋美術科・世界探究科"),
      @("食堂・昼食","専用食堂は無いが、昼はキッチンカー＋同キャンパスの愛知東邦大の学食(日替り440円・丼390円・カレー280円)を利用可"),
      @("最寄り駅・自宅から","名東区平和が丘。地下鉄東山線「一社」＋バス。黒川→名城線→栄→東山線で約40〜45分【目安】"),
      @("進学・特色","愛知東邦大ほか。世界探究科など特色コース。美術科もあり")
    );
    pros=@("愛知東邦大の学食が使える(440円〜)＋昼はキッチンカー","世界探究科など特色コース","美術科もあり多彩");
    cons=@("専用の学生食堂は無い(大学学食/キッチンカー利用)","名東区東部で北区からやや遠い");
    cols=@("東邦","名東区","52","大学学食","2.9","○名東");
    cmp=@("52","大学学食","★2.9","○名東","世界探究");
    ev_setsu="学校説明会 7/28-29・10/4(土)・10/11(土)／文理特進説明会 8/29(土)";
    ev_bunka="東邦祭(文化祭)＝例年9月・要確認";
    link="東邦のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/3001/" },
 @{ name="東海学園高校（とうかいがくえん）"; short="東海"; area="名古屋市天白区／大学隣接の学生食堂"; accent="2E8B57"; hensa="48〜55"; scale="飛翔55/留学54/明照48"; chip="大学隣接の学生食堂"; photo=($U+"tokaigakuen.jpg"); psrc="制服（参考写真・指定バッグ付）出典：A-制服店.com";
    rows=@(
      @("学科・コース","普通科(飛翔55/留学54/明照48)。仏教系・中高一貫"),
      @("食堂・昼食","同一法人の大学キャンパスに隣接し、学生食堂(麺類・丼物)を利用可。明るい雰囲気"),
      @("最寄り駅・自宅から","天白区中平。地下鉄鶴舞線「原」徒歩12分/市バス「平針新屋敷」。黒川→名城線→上前津→鶴舞線で約40〜45分【目安】"),
      @("進学・特色","東海学園大ほか。留学コースなど国際・進学。落ち着いた校風")
    );
    pros=@("大学隣接の学生食堂(麺・丼)が使える","留学コースなど進学/国際","落ち着いた校風");
    cons=@("「原」駅から徒歩約12分","北区からは通学やや遠い");
    cols=@("東海学園","天白区","48-55","大学学食","-","△天白");
    cmp=@("48-55","大学学食","★-","△天白","留学・進学");
    ev_setsu="オープンスクール(申込〜8/14)／学校説明会 9月〜(申込7/6〜)";
    ev_bunka="記念祭(文化祭)＝公式で要確認";
    link="東海学園のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/2759/" }
)
$N=$schools.Count

$pres=$ppt.Presentations.Add(0); $pres.PageSetup.SlideWidth=960; $pres.PageSetup.SlideHeight=540

# COVER
$s=$pres.Slides.Add(1,12); Set-Bg $s $NAVY
Add-Rect $s $msoShapeOval 760 -120 420 420 "2A3A78" $null 0 | Out-Null
Add-Rect $s $msoShapeOval -140 360 360 360 "2A3A78" $null 0 | Out-Null
Add-Rect $s $msoShapeOval 60 64 64 64 $CORAL $null 0 | Out-Null
Add-Text $s 60 60 600 40 "五郎レポート" 16 $ICE $false $alnL $anTop | Out-Null
Add-Text $s 60 165 860 120 "紗衣の 高校さがし ＜私立・食堂編＞" 40 $WHITE $true $alnL $anTop | Out-Null
Add-Text $s 64 300 860 50 "名古屋・私立・共学／食堂・学食が使える5校（制服写真・2026年度日程つき）" 17 $ICE $false $alnL $anTop | Out-Null
Add-Text $s 60 470 860 40 "作成：五郎（AI秘書）　2026年6月　自宅PC（名古屋市北区鳩岡）にて" 13 "9DB0E8" $false $alnL $anTop | Out-Null

# CONDITIONS
$s=$pres.Slides.Add(2,12); Set-Bg $s $BG
Add-Text $s 40 30 880 50 "さがした条件（6つ ＋ エリア）" 30 $NAVY $true $alnL $anTop | Out-Null
$conds=@(@("1","私立・共学","私立の男女共学校"),@("2","偏差値 50前後","学力・内申の目安"),@("3","食堂・学食がある","校内や系列大で温かい昼食"),@("4","部活・行事が活発","打ち込める部活・楽しい行事"),@("5","評判・特色がいい","通ってよかったと思える"),@("6","通いやすい","自宅(北区)から通えること"))
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
Add-Text $s 40 452 880 60 "※「食堂・学食が使える共学校」を条件に選定(校内食堂のほか、同一キャンパスの系列大学の学食が使える学校も含む)。偏差値・評判は2026年6月時点の目安。" 12 $MUTED $false $alnL $anTop | Out-Null

# TABLE
$s=$pres.Slides.Add(3,12); Set-Bg $s $BG
Add-Text $s 40 28 880 50 "食堂が使える私立・共学5校（北区から通えるエリア）" 24 $NAVY $true $alnL $anTop | Out-Null
$cols=@("学校","区","偏差値","食堂・昼食","評判","通学"); $cw=@(210,110,120,150,90,130)
$tx=40;$ty=88;$rh=58;$hh=38;$cxp=$tx
for($c=0;$c -lt 6;$c++){ Add-Rect $s $msoShapeRect $cxp $ty $cw[$c] $hh $NAVY $null 0 | Out-Null; Add-Text $s $cxp $ty $cw[$c] $hh $cols[$c] 14 $WHITE $true $alnC $anMid | Out-Null; $cxp+=$cw[$c] }
for($r=0;$r -lt $N;$r++){
  $rt=$ty+$hh+$r*$rh; $rowbg=if($r%2 -eq 0){$WHITE}else{"EEF2FB"}; $cxp=$tx; $row=$schools[$r].cols
  for($c=0;$c -lt 6;$c++){
    Add-Rect $s $msoShapeRect $cxp $rt $cw[$c] $rh $rowbg $LINE 0.75 | Out-Null; $val=$row[$c]
    if($c -eq 0){ Add-Rect $s $msoShapeOval ($cxp+8) ($rt+$rh/2-6) 12 12 $schools[$r].accent $null 0 | Out-Null; Add-Text $s ($cxp+26) $rt ($cw[$c]-30) $rh $val 12 $NAVY $true $alnL $anMid | Out-Null }
    elseif($c -eq 2){ Add-Text $s $cxp $rt $cw[$c] $rh $val 15 $schools[$r].accent $true $alnC $anMid | Out-Null }
    elseif($c -eq 3){ Add-Text $s $cxp $rt $cw[$c] $rh $val 12 $GREENINK $true $alnC $anMid | Out-Null }
    else { Add-Text $s $cxp $rt $cw[$c] $rh $val 12 $INK $false $alnC $anMid | Out-Null }
    $cxp+=$cw[$c]
  }
}
Add-Text $s 40 ($ty+$hh+$N*$rh+8) 880 50 "※「評判」はみんなの高校情報の総合評価(5点満点・私立は賛否が出やすい)。所要時間は自宅(北区鳩岡)からの目安。次ページ以降に各校の制服写真・食堂・口コミ・2026日程。" 11 $MUTED $false $alnL $anTop | Out-Null

# DETAIL
foreach($sc in $schools){
  $s=$pres.Slides.Add($pres.Slides.Count+1,12); Set-Bg $s $WHITE; $acc=$sc.accent
  Add-Rect $s $msoShapeRound 40 26 624 50 $acc $null 0 | Out-Null
  Add-Text $s 58 26 600 50 $sc.name 21 $WHITE $true $alnL $anMid | Out-Null
  Add-Text $s 44 80 624 24 $sc.area 13 $MUTED $false $alnL $anTop | Out-Null
  Add-Rect $s $msoShapeRound 686 26 234 76 "F2F5FC" $acc 1.25 | Out-Null
  Add-Text $s 686 32 234 18 "偏差値（目安）" 12 $MUTED $false $alnC $anTop | Out-Null
  Add-Text $s 686 46 234 42 $sc.hensa 26 $acc $true $alnC $anTop | Out-Null
  Add-Text $s 686 84 234 16 $sc.scale 10 $MUTED $false $alnC $anTop | Out-Null
  Add-Rect $s $msoShapeRound 686 108 234 28 $GREENINK $null 0 | Out-Null
  Add-Text $s 686 108 234 28 $sc.chip 10 $WHITE $true $alnC $anMid | Out-Null
  Add-Picture $s $sc.photo 686 144 234 176 | Out-Null
  Add-Text $s 686 322 234 26 $sc.psrc 9 $MUTED $false $alnC $anTop | Out-Null
  $ix=40;$iy=110;$rh2=46;$labW=128;$valW=484
  for($i=0;$i -lt $sc.rows.Count;$i++){
    $t=$iy+$i*$rh2
    Add-Rect $s $msoShapeRound $ix $t ($labW+$valW+12) ($rh2-7) "F7F9FD" $null 0 | Out-Null
    $lc = if($sc.rows[$i][0] -eq "食堂・昼食"){$GREENINK}else{$acc}
    Add-Text $s ($ix+12) $t $labW ($rh2-7) $sc.rows[$i][0] 13 $lc $true $alnL $anMid | Out-Null
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
Add-Text $s 40 26 880 46 "2026年度（令和8年度）の 説明会・オープンキャンパス・文化祭" 22 $NAVY $true $alnL $anTop | Out-Null
$ecols=@("学校","2026年度 説明会・オープンキャンパス","文化祭（学園祭）"); $ecw=@(160,460,260)
$ex=40;$ey=84;$erh=66;$ehh=36;$cxp=$ex
for($c=0;$c -lt 3;$c++){ Add-Rect $s $msoShapeRect $cxp $ey $ecw[$c] $ehh $NAVY $null 0 | Out-Null; Add-Text $s $cxp $ey $ecw[$c] $ehh $ecols[$c] 13 $WHITE $true $alnC $anMid | Out-Null; $cxp+=$ecw[$c] }
for($r=0;$r -lt $N;$r++){
  $rt=$ey+$ehh+$r*$erh; $rowbg=if($r%2 -eq 0){$WHITE}else{"EEF2FB"}
  Add-Rect $s $msoShapeRect $ex $rt $ecw[0] $erh $rowbg $LINE 0.75 | Out-Null
  Add-Rect $s $msoShapeOval ($ex+8) ($rt+$erh/2-6) 12 12 $schools[$r].accent $null 0 | Out-Null
  Add-Text $s ($ex+24) $rt ($ecw[0]-28) $erh $schools[$r].short 12 $NAVY $true $alnL $anMid | Out-Null
  Add-Rect $s $msoShapeRect ($ex+$ecw[0]) $rt $ecw[1] $erh $rowbg $LINE 0.75 | Out-Null
  Add-Text $s ($ex+$ecw[0]+10) $rt ($ecw[1]-18) $erh $schools[$r].ev_setsu 11 $INK $false $alnL $anMid | Out-Null
  Add-Rect $s $msoShapeRect ($ex+$ecw[0]+$ecw[1]) $rt $ecw[2] $erh $rowbg $LINE 0.75 | Out-Null
  Add-Text $s ($ex+$ecw[0]+$ecw[1]+10) $rt ($ecw[2]-18) $erh $schools[$r].ev_bunka 11 $INK $false $alnL $anMid | Out-Null
}
Add-Text $s 40 ($ey+$ehh+$N*$erh+6) 880 60 "※2026年6月時点で公式等から確認できた情報です。「要確認」は未発表/昨年度実績。私立は予約制が多く、学費・特待・推薦の条件とあわせて必ず各校公式でご確認ください。" 11 $MUTED $false $alnL $anTop | Out-Null

# COMPARISON
$s=$pres.Slides.Add($pres.Slides.Count+1,12); Set-Bg $s $BG
Add-Text $s 40 28 880 46 "5校をくらべると（早見表）" 28 $NAVY $true $alnL $anTop | Out-Null
$crit=@("偏差値","食堂・昼食","評判","通学","特色")
$lx=40;$ly=86;$lw=150;$colw=152;$rh3=60;$hh3=40
Add-Rect $s $msoShapeRect $lx $ly $lw $hh3 $NAVY $null 0 | Out-Null; Add-Text $s $lx $ly $lw $hh3 "項目" 13 $WHITE $true $alnC $anMid | Out-Null
for($c=0;$c -lt $N;$c++){ $cxp=$lx+$lw+$c*$colw; Add-Rect $s $msoShapeRect $cxp $ly $colw $hh3 $schools[$c].accent $null 0 | Out-Null; Add-Text $s $cxp $ly $colw $hh3 $schools[$c].short 13 $WHITE $true $alnC $anMid | Out-Null }
for($r=0;$r -lt 5;$r++){
  $rt=$ly+$hh3+$r*$rh3
  Add-Rect $s $msoShapeRect $lx $rt $lw $rh3 "E4E9F5" $LINE 0.75 | Out-Null
  Add-Text $s ($lx+8) $rt ($lw-14) $rh3 $crit[$r] 12 $NAVY $true $alnL $anMid | Out-Null
  for($c=0;$c -lt $N;$c++){
    $cxp=$lx+$lw+$c*$colw; $rowbg=if($r%2 -eq 0){$WHITE}else{"EEF2FB"}
    Add-Rect $s $msoShapeRect $cxp $rt $colw $rh3 $rowbg $LINE 0.75 | Out-Null; $val=$schools[$c].cmp[$r]
    if($r -eq 0){ Add-Text $s $cxp $rt $colw $rh3 $val 14 $schools[$c].accent $true $alnC $anMid | Out-Null }
    elseif($r -eq 1){ Add-Text $s ($cxp+4) $rt ($colw-8) $rh3 $val 11 $GREENINK $true $alnC $anMid | Out-Null }
    else { Add-Text $s ($cxp+4) $rt ($colw-8) $rh3 $val 11 $INK $false $alnC $anMid | Out-Null }
  }
}

# NEXT STEPS
$s=$pres.Slides.Add($pres.Slides.Count+1,12); Set-Bg $s $NAVY
Add-Text $s 40 36 880 50 "次にやること（おすすめの進め方）" 30 $WHITE $true $alnL $anTop | Out-Null
$steps=@(
 @("1","説明会・オープンキャンパスに行く","この資料の日程を参考に予約。雰囲気・先生・生徒を確認"),
 @("2","文化祭・部活・食堂を見に行く","学校の素顔が見える。食堂はできれば実食して確かめる"),
 @("3","通学を実際に試す","自宅(北区)から朝に。黒川・上飯田・バス(城北小学校)の乗換を体感"),
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
Add-Text $s 40 466 880 50 "五郎より：まずは通いやすい『市邨』『高蔵』『名電』へ。食堂の実食もおすすめ。費用面(特待・補助)も一緒に確認しましょう。" 14 "9DB0E8" $false $alnL $anTop | Out-Null

$out=$base+"高校さがし_私立編_五郎.pptx"
if(Test-Path $out){ Remove-Item $out -Force }
$pres.SaveAs($out,24); $pres.Close()
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
[GC]::Collect()
Write-Output ("SAVED: "+$out)
Write-Output ("SIZE: "+((Get-Item $out).Length)+" bytes")
