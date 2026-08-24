<#
  wall_clearance.ps1  -- 壁-壁 内法寸法 外部変形エンジン
  役割:
    Jw_cad が書き出した jwc_temp.txt から、選択された線分を読み取り、
    平行に向かい合う 2 本の内法(clear)寸法を計算し、Jw_cad 標準の
    寸法図形(msg) + 端部(点/矢印)として書き戻す。
  重要な前提(確定事項):
    - #bz なし = 座標は実寸そのまま。倍率は一切掛けない($ScaleDenominator 相当は 1 固定)。
    - jwc_temp.txt は Shift-JIS(CP932)。読み書きとも 932 で扱う。
  選択の使い分け:
    - ちょうど 2 本だけ選択 -> MinGap を無視してその 2 本を計測(手動指定)
    - 3 本以上           -> MinGap 以上/重なりあり/間に線が無い"隣接する対面"だけ自動抽出
#>
param(
  [string]$TempFile = "jwc_temp.txt",
  [double]$MinGap   = 500,   # 内法とみなす最小間隔(mm)。壁厚ペアを除外する閾値
  [double]$AngleTol = 1.0,   # 平行判定の許容角度(度)
  [double]$Offset   = 0,     # 寸法線位置: 重なり区間中央からの壁方向ずらし量(mm)
  [double]$RoundStep= 1,     # 寸法値の丸め(1 / 0.5 / 0.1)
  [string]$Orient   = "para",# para=寸法線と同じ角度 / hv=水平に整える
  [string]$Term     = "dot", # dot=点(Jw標準) / arrow=塗り矢印(三角) / line=線矢印(2本線) / slash=斜線(45度の止め)
  [int]   $Font     = 3,     # 寸法値の文字種(1-10)。TextH>0 のときは無視して任意サイズを使う
  [double]$TextW    = 0,     # 任意サイズ:文字幅(図寸mm)。>0 で有効(cn0を出力)
  [double]$TextH    = 0,     # 任意サイズ:文字高さ(図寸mm)。>0 で有効
  [double]$TextSp   = 0,     # 任意サイズ:文字間隔(図寸mm)
  [int]   $TextColor= 0,     # 任意サイズ:色番号(1-9)。0=1として出力
  [double]$OverlapMin = 300,   # 対面とみなす最小重なり長(mm)
  [switch]$Comma,            # 指定時: 寸法値を3桁カンマ区切り(2,200)にする。Jw本体の寸法設定「カンマ表示」と揃える用
  [switch]$NoFigure          # 指定時: 単線+文字で作図(寸法図形にしない)。既定は寸法図形(msg)+角度付き値
)

$ErrorActionPreference = "Stop"
$inv = [Globalization.CultureInfo]::InvariantCulture
$enc = [Text.Encoding]::GetEncoding(932)
function P([double]$v){ return $v.ToString("0.######",$inv) }

$scriptDir = if($PSScriptRoot){$PSScriptRoot}else{(Get-Location).Path}
$logPath = Join-Path $scriptDir "wall_clearance_log.txt"
$log = New-Object System.Collections.Generic.List[string]
function Log($m){ $log.Add(((Get-Date -Format "HH:mm:ss") + "  " + $m)) }

try {
  # --- temp file 解決 (CWD優先, 無ければスクリプトと同じ場所) ---
  $tp = $TempFile
  if(-not (Test-Path $tp)){ $alt = Join-Path $scriptDir "jwc_temp.txt"; if(Test-Path $alt){ $tp = $alt } }
  if(-not (Test-Path $tp)){ throw "jwc_temp.txt が見つかりません" }
  $tp = (Resolve-Path $tp).Path
  $lines = [IO.File]::ReadAllLines($tp,$enc)
  Log("読込: $tp  行数=$($lines.Count)")

  # --- 設定部 / 図形部 を最初の '#' で分割 ---
  $hashIdx = -1
  for($i=0;$i -lt $lines.Count;$i++){ if($lines[$i].Trim() -eq "#"){ $hashIdx=$i; break } }
  if($hashIdx -lt 0){ throw "区切り '#' が見つかりません" }
  $settings = if($hashIdx -gt 0){ $lines[0..($hashIdx-1)] } else { @() }
  $geom     = if($hashIdx -lt $lines.Count-1){ $lines[($hashIdx+1)..($lines.Count-1)] } else { @() }

  # --- ヘッダから 縮尺 / 文字サイズ / 書込レイヤG を取得 ---
  $hs=@(); $hcw=@(); $hch=@(); $lgIdx=0
  foreach($s in $settings){
    $t=$s.Trim()
    if($t -match '^hs\s+(.+)$'){ $hs = ($matches[1] -split '\s+') | %{ [double]::Parse($_,$inv) } }
    elseif($t -match '^hcw\s+(.+)$'){ $hcw = ($matches[1] -split '\s+') | %{ [double]::Parse($_,$inv) } }
    elseif($t -match '^hch\s+(.+)$'){ $hch = ($matches[1] -split '\s+') | %{ [double]::Parse($_,$inv) } }
    elseif($t -match '^lg([0-9a-fA-F])(\s|$)'){ $lgIdx = [Convert]::ToInt32($matches[1],16) }
  }
  $scale = if($hs.Count -ge 16){ $hs[$lgIdx] } else { 1 }
  if($scale -le 0){ $scale = 1 }
  $fi = [Math]::Max(0,[Math]::Min(9,$Font-1))
  $useCustom = ($TextH -gt 0 -and $TextW -gt 0)
  if($useCustom){
    # 任意サイズ(図寸mm)を指定。位置計算用は実寸に換算(図寸×縮尺分母)
    $charW = $TextW*$scale
    $charH = $TextH*$scale
    $col = if($TextColor -ge 1 -and $TextColor -le 9){ $TextColor } else { 1 }
    # cn0 [文字幅] [文字高さ] [文字間隔] [色番号]  (jwc_temp 仕様の並び)
    $cnLine = "cn0 "+(P $TextW)+" "+(P $TextH)+" "+(P $TextSp)+" "+$col
  } else {
    $charW = if($hcw.Count -ge 10){ $hcw[$fi]*$scale } else { 3*$scale }
    $charH = if($hch.Count -ge 10){ $hch[$fi]*$scale } else { 3*$scale }
    $cnLine = "cn$Font"
  }
  Log("縮尺分母=$scale  文字幅(実寸)=$charW  文字高(実寸)=$charH  cn=$cnLine  書込レイヤG=$lgIdx")

  # --- 図形部から線分だけ収集 (数値4トークンの行) ---
  $L=@()
  foreach($g in $geom){
    $t=$g.Trim()
    if($t -eq "" -or $t -eq "#"){ continue }
    $tok = $t -split '\s+'
    if($tok.Count -ne 4){ continue }
    $nums=@(); $ok=$true
    foreach($k in $tok){ $d=0.0; if([double]::TryParse($k,[Globalization.NumberStyles]::Float,$inv,[ref]$d)){ $nums+=$d } else { $ok=$false; break } }
    if(-not $ok){ continue }
    $dx=$nums[2]-$nums[0]; $dy=$nums[3]-$nums[1]; $len=[Math]::Sqrt($dx*$dx+$dy*$dy)
    if($len -lt 1e-6){ continue }
    $ang=([Math]::Atan2($dy,$dx))*180/[Math]::PI; $ang=(($ang%180)+180)%180
    $L += [pscustomobject]@{x1=$nums[0];y1=$nums[1];x2=$nums[2];y2=$nums[3];ux=$dx/$len;uy=$dy/$len;len=$len;ang=$ang}
  }
  Log("線分数=$($L.Count)")

  function AngClose($a,$b,$tol){ $d=[Math]::Abs($a-$b); if($d -gt 90){$d=180-$d}; return ($d -le $tol) }

  $forceTwo = ($L.Count -eq 2)

  # --- 平行ペア候補を抽出 ---
  $cand = New-Object System.Collections.Generic.List[object]
  for($i=0;$i -lt $L.Count;$i++){
    for($j=$i+1;$j -lt $L.Count;$j++){
      if(-not (AngClose $L[$i].ang $L[$j].ang $AngleTol)){ continue }
      $a=$L[$i]; $nx=-$a.uy; $ny=$a.ux
      $d1=($L[$j].x1-$a.x1)*$nx+($L[$j].y1-$a.y1)*$ny
      $d2=($L[$j].x2-$a.x1)*$nx+($L[$j].y2-$a.y1)*$ny
      $ds=($d1+$d2)/2; $d=[Math]::Abs($ds)
      if($d -lt 1){ continue }
      if(-not $forceTwo -and $d -lt $MinGap){ continue }
      $sg=[Math]::Sign($ds)
      $sb1=($L[$j].x1-$a.x1)*$a.ux+($L[$j].y1-$a.y1)*$a.uy
      $sb2=($L[$j].x2-$a.x1)*$a.ux+($L[$j].y2-$a.y1)*$a.uy
      $o1=[Math]::Max(0,[Math]::Min($sb1,$sb2)); $o2=[Math]::Min($a.len,[Math]::Max($sb1,$sb2))
      $ov=$o2-$o1
      if(-not $forceTwo -and $ov -lt $OverlapMin){ continue }
      $cand.Add([pscustomobject]@{i=$i;j=$j;d=$d;o1=$o1;o2=$o2;nx=$nx*$sg;ny=$ny*$sg})
    }
  }

  # --- "間に別の平行線が無い=隣接する対面" だけ採用 ---
  $pairs = New-Object System.Collections.Generic.List[object]
  foreach($p in $cand){
    if($forceTwo){ $pairs.Add($p); continue }
    $a=$L[$p.i]; $between=$false
    for($k=0;$k -lt $L.Count;$k++){
      if($k -eq $p.i -or $k -eq $p.j){ continue }
      if(-not (AngClose $a.ang $L[$k].ang $AngleTol)){ continue }
      $ok1=($L[$k].x1-$a.x1)*$p.nx+($L[$k].y1-$a.y1)*$p.ny
      $ok2=($L[$k].x2-$a.x1)*$p.nx+($L[$k].y2-$a.y1)*$p.ny
      $ov=($ok1+$ok2)/2
      if($ov -gt 5 -and $ov -lt ($p.d-5)){
        $sk1=($L[$k].x1-$a.x1)*$a.ux+($L[$k].y1-$a.y1)*$a.uy
        $sk2=($L[$k].x2-$a.x1)*$a.ux+($L[$k].y2-$a.y1)*$a.uy
        $io1=[Math]::Max($p.o1,[Math]::Min($sk1,$sk2)); $io2=[Math]::Min($p.o2,[Math]::Max($sk1,$sk2))
        if(($io2-$io1) -gt 0){ $between=$true; break }
      }
    }
    if(-not $between){ $pairs.Add($p) }
  }
  Log("候補ペア=$($cand.Count)  採用ペア=$($pairs.Count)")

  # --- 出力生成 ---
  $out = New-Object System.Collections.Generic.List[string]
  foreach($p in $pairs){
    $a=$L[$p.i]
    $uc=(($p.o1+$p.o2)/2)+$Offset
    $P1x=$a.x1+$a.ux*$uc; $P1y=$a.y1+$a.uy*$uc
    $P2x=$P1x+$p.nx*$p.d; $P2y=$P1y+$p.ny*$p.d
    $val=[Math]::Round($p.d/$RoundStep,[MidpointRounding]::AwayFromZero)*$RoundStep
    $intFmt = if($Comma){"#,##0"}else{"0"}; $decFmt = if($Comma){"#,##0.0"}else{"0.0"}
    $valStr = if($RoundStep -ge 1){ ([Math]::Round($val,[MidpointRounding]::AwayFromZero)).ToString($intFmt,$inv) } else { $val.ToString($decFmt,$inv) }
    $Mx=($P1x+$P2x)/2; $My=($P1y+$P2y)/2
    # 文字方向: hv=水平 / para=寸法線と同じ角度
    if($Orient -eq "hv"){ $tdx=1.0; $tdy=0.0 } else { $tdx=$p.nx; $tdy=$p.ny }
    # 左→右(または下→上)に読める向きへ正規化(上下逆さ防止)
    if($tdx -lt -1e-9 -or ([Math]::Abs($tdx) -lt 1e-9 -and $tdy -lt 0)){ $tdx=-$tdx; $tdy=-$tdy }
    $tw=$valStr.Length*$charW
    # 文字は基準ベクトルの左側(=perp)へ伸びる。その側へ一定隙間で逃がすと寸法線と重ならない
    $perpx=-$tdy; $perpy=$tdx
    $g=$charH*0.3
    $Tsx=$Mx-$tdx*$tw/2+$perpx*$g; $Tsy=$My-$tdy*$tw/2+$perpy*$g
    # cs/ch の第2座標は「左下を(0,0)とした方向ベクトル」= 角度指定(仕様上 cs も角度を持てる)
    $Vx=$tdx*$tw; $Vy=$tdy*$tw
    if($NoFigure){
      # フォールバック: 寸法線(単線) + 角度付き文字(ch)
      $out.Add(((P $P1x)+" "+(P $P1y)+" "+(P $P2x)+" "+(P $P2y)))
      $out.Add($cnLine)
      $out.Add(("ch "+(P $Tsx)+" "+(P $Tsy)+" "+(P $Vx)+" "+(P $Vy)+' "'+$valStr))
    } else {
      # 既定: 寸法図形(msg = 寸法線 + 寸法値, 値に方向ベクトルで角度を付与)
      $out.Add("msg")
      $out.Add(((P $P1x)+" "+(P $P1y)+" "+(P $P2x)+" "+(P $P2y)))
      $out.Add($cnLine)
      $out.Add(("cs "+(P $Tsx)+" "+(P $Tsy)+" "+(P $Vx)+" "+(P $Vy)+' "'+$valStr))
      $out.Add("#")
    }
    # 端部
    if($Term -eq "dot"){
      $out.Add(("pt "+(P $P1x)+" "+(P $P1y)))
      $out.Add(("pt "+(P $P2x)+" "+(P $P2y)))
    } elseif($Term -eq "slash"){
      # 斜線(止め): 寸法線に対し45度の短い線が端点を横切る
      $sq=[Math]::Sqrt(2.0); $sl=$charH*0.35
      $sx=($p.nx-$p.ny)/$sq; $sy=($p.ny+$p.nx)/$sq
      foreach($e in @(@($P1x,$P1y),@($P2x,$P2y))){
        $out.Add(((P ($e[0]-$sx*$sl))+" "+(P ($e[1]-$sy*$sl))+" "+(P ($e[0]+$sx*$sl))+" "+(P ($e[1]+$sy*$sl))))
      }
    } elseif($Term -eq "line"){
      # 線矢印: 塗りつぶさない2本線の開き矢印
      $al=$charH*0.5; $px=-$p.ny; $py=$p.nx; $aw=$al*0.35
      $out.Add(((P $P1x)+" "+(P $P1y)+" "+(P ($P1x+$p.nx*$al+$px*$aw))+" "+(P ($P1y+$p.ny*$al+$py*$aw))))
      $out.Add(((P $P1x)+" "+(P $P1y)+" "+(P ($P1x+$p.nx*$al-$px*$aw))+" "+(P ($P1y+$p.ny*$al-$py*$aw))))
      $out.Add(((P $P2x)+" "+(P $P2y)+" "+(P ($P2x-$p.nx*$al+$px*$aw))+" "+(P ($P2y-$p.ny*$al+$py*$aw))))
      $out.Add(((P $P2x)+" "+(P $P2y)+" "+(P ($P2x-$p.nx*$al-$px*$aw))+" "+(P ($P2y-$p.ny*$al-$py*$aw))))
    } else {
      $al=$charH*0.5; $px=-$p.ny; $py=$p.nx; $aw=$al*0.35
      $b1x=$P1x+$p.nx*$al+$px*$aw; $b1y=$P1y+$p.ny*$al+$py*$aw
      $b2x=$P1x+$p.nx*$al-$px*$aw; $b2y=$P1y+$p.ny*$al-$py*$aw
      $out.Add(("sl "+(P $P1x)+" "+(P $P1y)+" "+(P $b1x)+" "+(P $b1y)+" "+(P $b2x)+" "+(P $b2y)))
      $c1x=$P2x-$p.nx*$al+$px*$aw; $c1y=$P2y-$p.ny*$al+$py*$aw
      $c2x=$P2x-$p.nx*$al-$px*$aw; $c2y=$P2y-$p.ny*$al-$py*$aw
      $out.Add(("sl "+(P $P2x)+" "+(P $P2y)+" "+(P $c1x)+" "+(P $c1y)+" "+(P $c2x)+" "+(P $c2y)))
    }
    Log("採用 d=$(P $p.d) -> $valStr  P1=($(P $P1x),$(P $P1y)) P2=($(P $P2x),$(P $P2y))")
  }

  if($pairs.Count -eq 0){
    $out.Clear()
    if($L.Count -lt 2){ $out.Add("he線が2本以上選択されていません") }
    elseif($forceTwo){ $out.Add("he選択した2本が平行でないか重なりがありません") }
    else{ $out.Add("he内法ペアが見つかりません(最小間隔=$(P $MinGap)mm)") }
  }

  [IO.File]::WriteAllLines($tp,$out,$enc)
  Log("出力行=$($out.Count)")
}
catch {
  try { [IO.File]::WriteAllLines($tp,@(("heエラー: "+$_.Exception.Message)),$enc) } catch {}
  Log("例外: "+$_.Exception.Message)
}
finally {
  try { [IO.File]::WriteAllLines($logPath,$log,(New-Object System.Text.UTF8Encoding($true))) } catch {}
}
