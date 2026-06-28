# 高校さがし 公立編 v2 : 黒川アクセス＋制服写真＋2026イベント
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing

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
$alnL=1; $alnC=2; $anTop=1; $anMid=3

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
function Add-Para($s,$l,$t,$w,$h,$lines,$size,$colorHex){
  $txt=($lines | ForEach-Object { "・"+$_ }) -join "`r"
  return (Add-Text $s $l $t $w $h $txt $size $colorHex $false $alnL $anTop)
}
function Add-Link($s,$l,$t,$w,$h,$text,$size,$colorHex,$url){
  $sp=Add-Text $s $l $t $w $h $text $size $colorHex $false $alnL $anMid
  $sp.TextFrame.TextRange.ActionSettings.Item(1).Hyperlink.Address=$url; return $sp
}
function Add-Picture($s,$path,$boxX,$boxY,$boxW,$boxH){
  $img=[System.Drawing.Image]::FromFile($path); $iw=$img.Width; $ih=$img.Height; $img.Dispose()
  $scale=[math]::Min($boxW/$iw,$boxH/$ih); $w=$iw*$scale; $h=$ih*$scale
  $x=$boxX+($boxW-$w)/2; $y=$boxY+($boxH-$h)/2
  $pic=$s.Shapes.AddPicture($path,$msoFalse,$msoTrue,$x,$y,$w,$h)
  $pic.Line.Visible=$msoTrue; $pic.Line.ForeColor.RGB=[int](CRGB "FFFFFF"); $pic.Line.Weight=[single]3
  return $pic
}

$base="C:\Users\山本　史朗\OneDrive\Desktop\新しいフォルダー\"
$U=$base+"uniforms\"

$schools=@(
 @{ name="名古屋市立 北高校（きた）"; short="北高校"; area="名古屋市北区如来町／自宅と同じ北区＝最短"; accent="0E8C8B"; hensa="50"; scale="市内・公立で中堅"; bad="あり"; photo=($U+"kita.jpg"); psrc="制服（参考写真）出典：A-制服店.com";
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
    ev_setsu="国際理解コース説明会 8/4(火)10:00 ／ 学校見学会 10/10(土)13:30";
    ev_bunka="北高祭（例年9月頃・日程は公式で要確認）";
    link="北高校のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/3391/" },
 @{ name="名古屋市立 山田高校（やまだ）"; short="山田高校"; area="名古屋市西区二方町／最寄り 上小田井"; accent="E8633F"; hensa="49"; scale="市内・公立で中堅"; bad="◎人気"; photo=($U+"yamada.jpg"); psrc="制服（参考写真）出典：A-制服店.com";
    rows=@(
      @("学科・コース","普通科"),
      @("特色","明るく楽しい校風。ダンス部が全国常連、行事・部活が活発。制服も人気"),
      @("最寄り駅・自宅から","地下鉄鶴舞線・名鉄犬山線「上小田井」駅。黒川→(名城線)→乗換で約40分【目安】"),
      @("進学・通学","同志社・立命館・関西・愛知工大など私大中心／通学〇")
    );
    pros=@("とにかく楽しい・明るい校風","部活・行事が盛ん","制服がかわいいと人気");
    cons=@("勉強面の手厚さは控えめとの声","校則は服装に厳しめ");
    cols=@("名古屋市立 山田高校","西区","49","◎人気","3.5","〇上小田井");
    cmp=@("49","◎人気","★3.5","〇西区","部活・評判◎");
    ev_setsu="日程は公式サイト・SNSで要確認（例年 夏〜秋に説明会・体験入学）";
    ev_bunka="山高祭（例年9月頃・2日間・要確認）";
    link="山田高校のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/3402/" },
 @{ name="愛知県立 中村高校（なかむら）"; short="中村高校"; area="名古屋市中村区／自由な校風（1953開校）"; accent="7B5EA7"; hensa="50"; scale="市内・公立で中堅"; bad="県大会"; photo=($U+"nakamura.jpg"); psrc="制服＝紺セーラー服（参考写真）出典：A-制服店.com";
    rows=@(
      @("学科・コース","普通科（ほかに美術科）"),
      @("特色","伝統校で自主自律。制服は紺のセーラー服。弓道・バドミントン(県大会)・陸上・体操など部活が熱心"),
      @("最寄り駅・自宅から","地下鉄東山線「中村公園」ほか。黒川→(名城線)→栄→東山線で約30〜35分【目安】"),
      @("進学・通学","私立大学中心に幅広く進学／通学〇")
    );
    pros=@("自由でのびのびした校風","部活に打ち込みやすい","伝統と落ち着きがある");
    cons=@("校舎・設備は年季あり","自由ゆえ自己管理が必要");
    cols=@("愛知県立 中村高校","中村区","50","県大会","--","〇中村区");
    cmp=@("50","県大会","--","〇中村","自由な校風");
    ev_setsu="体験入学(普通科) 8/8(金)午前 ※要確認（美術科 7/22）";
    ev_bunka="学校祭（体育祭＋文化祭・例年9月）";
    link="中村高校のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/123/" },
 @{ name="愛知県立 春日井南高校（チャレンジ校）"; short="春日井南"; area="春日井市／少し上のチャレンジ校"; accent="2E8B57"; hensa="54"; scale="やや上のレベル"; bad="要確認"; photo=($U+"kasugaiminami.jpg"); psrc="制服（参考写真・エンジのタイ）出典：制服市場";
    rows=@(
      @("学科・コース","普通科"),
      @("特色","校舎がきれいで自習室が夜まで。制服は紺ブレザー＋エンジのタイ。運動部14・文化部10と部活豊富"),
      @("最寄り駅・自宅から","JR中央線「春日井」駅ほか。自宅から名鉄小牧線(上飯田)・JRで約30〜40分【目安】"),
      @("進学・通学","国公立に毎年20名前後。中部大・名城大・中京大ほか／通学△〜〇")
    );
    pros=@("学習環境が良い(自習室・きれいな校舎)","評判が高い(県内上位)","部活も活発");
    cons=@("市外で通学はやや遠い","進学校で課題・勉強量は多め");
    cols=@("愛知県立 春日井南","春日井市","54","要確認","3.9","△市外");
    cmp=@("54","要確認","★3.9","△市外","環境◎・上狙い");
    ev_setsu="学校説明会 8/18(火) 春日井市民会館 ／ 学校見学会 10/11(日)9:00-15:00";
    ev_bunka="春陵祭（例年秋・日程は公式で要確認）";
    link="春日井南高校のくわしい情報（口コミ・偏差値・制服）はこちら（みんなの高校情報）"; url="https://www.minkou.jp/hischool/school/52/" }
)

$pres=$ppt.Presentations.Add(0)
$pres.PageSetup.SlideWidth=960; $pres.PageSetup.SlideHeight=540

# COVER
$s=$pres.Slides.Add(1,12); Set-Bg $s $NAVY
Add-Rect $s $msoShapeOval 760 -120 420 420 "2A3A78" $null 0 | Out-Null
Add-Rect $s $msoShapeOval -140 360 360 360 "2A3A78" $null 0 | Out-Null
Add-Rect $s $msoShapeOval 60 64 64 64 $CORAL $null 0 | Out-Null
Add-Text $s 60 60 600 40 "五郎レポート" 16 $ICE $false $alnL $anTop | Out-Null
Add-Text $s 60 165 860 120 "紗衣の 高校さがし ＜公立編＞" 46 $WHITE $true $alnL $anTop | Out-Null
Add-Text $s 64 300 860 50 "名古屋市内・公立・共学／偏差値50前後・部活が活発・評判◎（制服写真・2026年度の見学日程つき）" 17 $ICE $false $alnL $anTop | Out-Null
Add-Text $s 60 470 860 40 "作成：五郎（AI秘書）　2026年6月　自宅PC（名古屋市北区鳩岡）にて" 13 "9DB0E8" $false $alnL $anTop | Out-Null

# CONDITIONS
$s=$pres.Slides.Add(2,12); Set-Bg $s $BG
Add-Text $s 40 30 880 50 "さがした条件（6つ ＋ エリア）" 30 $NAVY $true $alnL $anTop | Out-Null
$conds=@(@("1","公立","私立ではなく公立高校"),@("2","共学","男女いっしょに学ぶ学校"),@("3","偏差値 50前後","学力・内申の目安ライン"),@("4","部活・行事が活発","打ち込める部活・楽しい行事"),@("5","評判・特色がいい","通ってよかったと思える"),@("6","通いやすい","自宅(北区)から通えること"))
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
Add-Text $s 60 372 850 66 "エリア：自宅＝名古屋市北区鳩岡。地下鉄の最寄り「黒川」駅(名城線)／名鉄小牧線は「上飯田」駅／JRは「味美」(城北線)を起点に通学。" 15 $NAVY $true $alnL $anMid | Out-Null
Add-Text $s 40 452 880 60 "※バドミントン部の条件は外し、「共学・偏差値50前後・通学・評判」で選定。偏差値・評判・所要時間は2026年6月時点の目安です。" 12 $MUTED $false $alnL $anTop | Out-Null

# TABLE
$s=$pres.Slides.Add(3,12); Set-Bg $s $BG
Add-Text $s 40 30 880 50 "公立の候補4校（北区から通えるエリア）" 28 $NAVY $true $alnL $anTop | Out-Null
$cols=@("学校","区・市","偏差値","部活","評判","通学"); $cw=@(250,140,120,120,100,150)
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
Add-Text $s 40 430 880 80 "※「評判」はみんなの高校情報の総合評価（5点満点）。所要時間は自宅(北区鳩岡)からの目安。次ページ以降に各校の制服写真・口コミ・2026年度の見学日程を掲載。" 12 $MUTED $false $alnL $anTop | Out-Null

# DETAIL (with photo)
foreach($sc in $schools){
  $s=$pres.Slides.Add($pres.Slides.Count+1,12); Set-Bg $s $WHITE; $acc=$sc.accent
  Add-Rect $s $msoShapeRound 40 26 624 50 $acc $null 0 | Out-Null
  Add-Text $s 58 26 600 50 $sc.name 22 $WHITE $true $alnL $anMid | Out-Null
  Add-Text $s 44 80 624 24 $sc.area 13 $MUTED $false $alnL $anTop | Out-Null
  Add-Rect $s $msoShapeRound 686 26 234 76 "F2F5FC" $acc 1.25 | Out-Null
  Add-Text $s 686 32 234 18 "偏差値（目安）" 12 $MUTED $false $alnC $anTop | Out-Null
  Add-Text $s 686 46 234 42 $sc.hensa 28 $acc $true $alnC $anTop | Out-Null
  Add-Text $s 686 84 234 16 $sc.scale 10 $MUTED $false $alnC $anTop | Out-Null
  Add-Rect $s $msoShapeRound 686 108 234 28 $acc $null 0 | Out-Null
  Add-Text $s 686 108 234 28 ("バドミントン部：" + $sc.bad) 12 $WHITE $true $alnC $anMid | Out-Null
  # uniform photo
  Add-Picture $s $sc.photo 686 144 234 176 | Out-Null
  Add-Text $s 686 322 234 26 $sc.psrc 9 $MUTED $false $alnC $anTop | Out-Null
  # info rows
  $ix=40;$iy=110;$rh2=46;$labW=128;$valW=484
  for($i=0;$i -lt $sc.rows.Count;$i++){
    $t=$iy+$i*$rh2
    Add-Rect $s $msoShapeRound $ix $t ($labW+$valW+12) ($rh2-7) "F7F9FD" $null 0 | Out-Null
    Add-Text $s ($ix+12) $t $labW ($rh2-7) $sc.rows[$i][0] 13 $acc $true $alnL $anMid | Out-Null
    Add-Text $s ($ix+12+$labW) $t $valW ($rh2-7) $sc.rows[$i][1] 12 $INK $false $alnL $anMid | Out-Null
  }
  # pros/cons (full width, below)
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
Add-Text $s 40 30 880 50 "2026年度（令和8年度）の 見学・説明会・文化祭" 26 $NAVY $true $alnL $anTop | Out-Null
$ecols=@("学校","2026年度 説明会・見学・体験入学","文化祭（学園祭）"); $ecw=@(180,460,240)
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
Add-Text $s 40 424 880 90 "※2026年6月時点で各校公式サイトから確認できた情報です。「要確認」は未発表または昨年度実績で、変更の可能性があります。申込方法・最新日程・文化祭の公開可否は必ず各校公式でご確認ください。" 12 $MUTED $false $alnL $anTop | Out-Null

# COMPARISON
$s=$pres.Slides.Add($pres.Slides.Count+1,12); Set-Bg $s $BG
Add-Text $s 40 30 880 50 "4校をくらべると（早見表）" 30 $NAVY $true $alnL $anTop | Out-Null
$crit=@("偏差値","部活（バドミントン）","評判","通学のしやすさ","特色・ポイント")
$lx=40;$ly=96;$lw=190;$colw=170;$rh3=62;$hh3=44
Add-Rect $s $msoShapeRect $lx $ly $lw $hh3 $NAVY $null 0 | Out-Null; Add-Text $s $lx $ly $lw $hh3 "項目" 14 $WHITE $true $alnC $anMid | Out-Null
for($c=0;$c -lt 4;$c++){ $cxp=$lx+$lw+$c*$colw; Add-Rect $s $msoShapeRect $cxp $ly $colw $hh3 $schools[$c].accent $null 0 | Out-Null; Add-Text $s $cxp $ly $colw $hh3 $schools[$c].short 14 $WHITE $true $alnC $anMid | Out-Null }
for($r=0;$r -lt 5;$r++){
  $rt=$ly+$hh3+$r*$rh3
  Add-Rect $s $msoShapeRect $lx $rt $lw $rh3 "E4E9F5" $LINE 0.75 | Out-Null
  Add-Text $s ($lx+10) $rt ($lw-16) $rh3 $crit[$r] 13 $NAVY $true $alnL $anMid | Out-Null
  for($c=0;$c -lt 4;$c++){
    $cxp=$lx+$lw+$c*$colw; $rowbg=if($r%2 -eq 0){$WHITE}else{"EEF2FB"}
    Add-Rect $s $msoShapeRect $cxp $rt $colw $rh3 $rowbg $LINE 0.75 | Out-Null; $val=$schools[$c].cmp[$r]
    if($r -eq 0){ Add-Text $s $cxp $rt $colw $rh3 $val 16 $schools[$c].accent $true $alnC $anMid | Out-Null }
    else { Add-Text $s ($cxp+6) $rt ($colw-12) $rh3 $val 12 $INK $false $alnC $anMid | Out-Null }
  }
}

# NEXT STEPS
$s=$pres.Slides.Add($pres.Slides.Count+1,12); Set-Bg $s $NAVY
Add-Text $s 40 36 880 50 "次にやること（おすすめの進め方）" 30 $WHITE $true $alnL $anTop | Out-Null
$steps=@(
 @("1","説明会・学校見学会に行く","この資料の2026日程を参考に予約。親子で雰囲気を確認"),
 @("2","部活・文化祭を見に行く","文化祭は学校の素顔が見えるチャンス。部活の様子も確認"),
 @("3","通学を実際に試す","自宅(北区)から朝の時間帯に。黒川・上飯田からの乗換を体感"),
 @("4","倍率・募集要項を公式で確認","この資料の数値・日程は目安。最新は必ず各校公式で")
)
for($i=0;$i -lt 4;$i++){
  $t=104+$i*84
  Add-Rect $s $msoShapeRound 40 $t 880 72 "2A3A78" $null 0 | Out-Null
  Add-Rect $s $msoShapeOval 58 ($t+16) 40 40 $CORAL $null 0 | Out-Null
  Add-Text $s 58 ($t+16) 40 40 $steps[$i][0] 20 $WHITE $true $alnC $anMid | Out-Null
  Add-Text $s 116 ($t+11) 790 30 $steps[$i][1] 18 $WHITE $true $alnL $anTop | Out-Null
  Add-Text $s 116 ($t+41) 790 26 $steps[$i][2] 13 $ICE $false $alnL $anTop | Out-Null
}
Add-Text $s 40 466 880 50 "五郎より：まずは地元の『北高校』(8/4説明会・10/10見学会)から。気になる順に一緒に動きましょう。応援しています。" 14 "9DB0E8" $false $alnL $anTop | Out-Null

$out=$base+"高校さがし_公立編_五郎.pptx"
if(Test-Path $out){ Remove-Item $out -Force }
$pres.SaveAs($out,24); $pres.Close()
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
[GC]::Collect()
Write-Output ("SAVED: "+$out)
Write-Output ("SIZE: "+((Get-Item $out).Length)+" bytes")
