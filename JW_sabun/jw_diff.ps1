# kreatio-works  https://kreatio-works.github.io/
<#
  jw_diff.ps1  -- Jw_cad 図面差分チェッカー 外部変形エンジン (v1)
  役割:
    -Mode save    : 旧図面の選択図形を正規化前の生値で jw_diff_base.txt に保存する
    -Mode compare : 新図面の選択図形を base と照合し、差分マーカーを jwc_temp.txt に書き戻す
  重要な前提(確定事項):
    - #bz なし = 座標は実寸そのまま。倍率は一切掛けない。
    - jwc_temp.txt は Shift-JIS(CP932)。読み書きとも 932 で扱う。
    - base / log は UTF-8(BOM)。
    - ch/cv/cs の第2座標は基準点からの相対方向ベクトル。
  比較(v1): 幾何と文字内容のみ。属性(線色/線種/レイヤ)は比較しない。
  実装メモ: 数万要素を 30 秒以内で処理するため、要素ごとのループ内では
            PowerShell 関数呼び出しを避けてインライン展開している(関数呼び出しは 1 回 ~0.1ms)。
#>
param(
  [Parameter(Mandatory=$true)][ValidateSet("save","compare")]
  [string]$Mode,
  [string]$TempFile = "jwc_temp.txt",
  [string]$BaseFile = "jw_diff_base.txt",
  [double]$Tol      = 0.5,   # 座標一致の許容誤差(mm) = 量子化セル幅
  [double]$MarkR    = 100,   # 円マーカー半径(mm, 実寸)
  [int]   $ColorDel = 4,     # 削除マーカー線色
  [int]   $ColorAdd = 5,     # 追加マーカー線色
  [string]$Layer    = "F",   # マーカー作図先レイヤグループ(0-9,A-F)。空=現在のまま
  [int]   $MaxMark  = 500,   # 作図マーカー上限
  [switch]$NoRedraw,         # 指定時: 削除要素の幾何再作図を省略し円マーカーのみ
  [switch]$NoShift           # 指定時: 図面全体の平行移動の自動検出・補正を行わない
)

$ErrorActionPreference = "Stop"
$VERSION = "jw_diff v1.0"
$inv = [Globalization.CultureInfo]::InvariantCulture
$NS  = [Globalization.NumberStyles]::Float
$enc = [Text.Encoding]::GetEncoding(932)
$utf8bom = New-Object System.Text.UTF8Encoding($true)
function P([double]$v){ return $v.ToString("0.######",$inv) }
function P1([double]$v){ return $v.ToString("0.0",$inv) }

$scriptDir = if($PSScriptRoot){$PSScriptRoot}else{(Get-Location).Path}
$logPath = Join-Path $scriptDir "jw_diff_log.txt"
$log = New-Object System.Collections.Generic.List[string]
function Log($m){ $log.Add([string]$m) }

# base の場所: 絶対パスならそのまま、相対なら ps1 と同じフォルダ
$basePath = if([IO.Path]::IsPathRooted($BaseFile)){ $BaseFile } else { Join-Path $scriptDir $BaseFile }

# ------------------------------------------------------------------
# 要素 = hashtable {Type(ln/ch/cv/cs/cc/ci/pt/sl/raw); Nums(double[]); Text; Extra(文字:cn行 / raw:生行); Ints; Key}
# 以下は差分要素(少数)にしか使わない関数群
# ------------------------------------------------------------------
function RepPoint($e){   # 代表点: 線=中点 / 文字=基準点 / 円=中心 / ソリッド=重心 / raw=数値ペアの平均
  $n=$e.Nums
  switch($e.Type){
    "ln" { return @((($n[0]+$n[2])/2),(($n[1]+$n[3])/2)) }
    "sl" { $sx=0.0;$sy=0.0;$c=0; for($i=0;$i+1 -lt $n.Count;$i+=2){ $sx+=$n[$i]; $sy+=$n[$i+1]; $c++ }; if($c -eq 0){return $null}; return @(($sx/$c),($sy/$c)) }
    "raw" {
      $v=@(); foreach($k in ($e.Extra -split '\s+')){ $d=0.0; if([double]::TryParse($k,$NS,$inv,[ref]$d)){ $v += $d } }
      if($v.Count -lt 2){ return $null }
      $sx=0.0;$sy=0.0;$c=0; for($i=0;$i+1 -lt $v.Count;$i+=2){ $sx+=$v[$i]; $sy+=$v[$i+1]; $c++ }; return @(($sx/$c),($sy/$c))
    }
    default { return @($n[0],$n[1]) }
  }
}
function Describe($e){   # ログ用(生値・小数1桁)
  $n=$e.Nums
  switch($e.Type){
    "ln" { return ("線分  ("+(P1 $n[0])+", "+(P1 $n[1])+")-("+(P1 $n[2])+", "+(P1 $n[3])+")") }
    { $_ -in "ch","cv","cs","cc" } { return ("文字"+$(if($e.Type -ne "ch"){"["+$e.Type+"]"}else{"  "})+"("+(P1 $n[0])+", "+(P1 $n[1])+") `""+$e.Text+"`"") }
    "ci" { $s="円    ("+(P1 $n[0])+", "+(P1 $n[1])+")"; if($n.Count -ge 3){ $s+=" r="+(P1 $n[2]) }; if($n.Count -gt 3){ $s+=" ["+((($n | Select-Object -Skip 3) | %{ P1 $_ }) -join ' ')+"]" }; return $s }
    "pt" { return ("点    ("+(P1 $n[0])+", "+(P1 $n[1])+")") }
    "bl" { return ("ブロック ("+(P1 $n[0])+", "+(P1 $n[1])+") `""+$e.Text+"`"") }
    "sl" { $s="ソリッド "; for($i=0;$i+1 -lt $n.Count;$i+=2){ $s+="("+(P1 $n[$i])+", "+(P1 $n[$i+1])+")" }; return $s }
    default { return ("未知  "+$e.Extra) }
  }
}
function GeomLines($e){  # 削除要素の幾何を再作図する行
  $n=$e.Nums; $o=@()
  switch($e.Type){
    "ln" { $o += ((P $n[0])+" "+(P $n[1])+" "+(P $n[2])+" "+(P $n[3])) }
    { $_ -in "ch","cv","cs","cc" } {
      if($e.Extra -ne ""){ $o += $e.Extra }
      $tt = if($e.Type -eq "cs"){"ch"}else{$e.Type}   # 寸法値は通常文字として再作図
      $o += ($tt+" "+(P $n[0])+" "+(P $n[1])+" "+(P $n[2])+" "+(P $n[3])+' "'+$e.Text)
    }
    "raw" { $o += $e.Extra }
    "bl"  { }   # ブロックは新図に定義が無い可能性があるので再作図せずマーカーのみ
    default { $o += ($e.Type+" "+(($n | %{ P $_ }) -join ' ')) }
  }
  return $o
}

# ------------------------------------------------------------------
$tp = $null
try {
  Log("[jw_diff] "+(Get-Date -Format "yyyy-MM-dd HH:mm")+"  Mode=$Mode  Tol=$(P $Tol)  ($VERSION)")
  $sw = [Diagnostics.Stopwatch]::StartNew()

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
  $geomStart = $hashIdx+1

  # --- ヘッダ: 縮尺 / 文字サイズ / 書込レイヤG / 選択枠 ---
  $hs=@(); $hcw=@(); $hch=@(); $lgIdx=0; $hn=$null
  foreach($s in $settings){
    $t=$s.Trim()
    if($t -match '^hs\s+(.+)$'){ $hs = ($matches[1] -split '\s+') | %{ [double]::Parse($_,$inv) } }
    elseif($t -match '^hcw\s+(.+)$'){ $hcw = ($matches[1] -split '\s+') | %{ [double]::Parse($_,$inv) } }
    elseif($t -match '^hch\s+(.+)$'){ $hch = ($matches[1] -split '\s+') | %{ [double]::Parse($_,$inv) } }
    elseif($t -match '^lg([0-9a-fA-F])(\s|$)'){ $lgIdx = [Convert]::ToInt32($matches[1],16) }
    elseif($t -match '^hn\s+(.+)$'){ $hn = @(); foreach($k in ($matches[1] -split '\s+')){ $d=0.0; if([double]::TryParse($k,$NS,$inv,[ref]$d)){ $hn += $d } } }
  }
  $scale = if($hs.Count -ge 16){ $hs[$lgIdx] } else { 1 }
  if($scale -le 0){ $scale = 1 }
  $charW = if($hcw.Count -ge 10){ $hcw[2]*$scale } else { 3*$scale }   # 文字種3
  $charH = if($hch.Count -ge 10){ $hch[2]*$scale } else { 3*$scale }

  # ================================================================
  # 図形部パース(ホットパス・インライン)
  # ================================================================
  $newEls = New-Object 'System.Collections.Generic.List[object]'
  $curCn = ""; $unknown = 0
  $rxText = [regex]'^(c[hvsc])\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+"(.*)$'
  $rxBlock = [regex]'^BL\s+(\S+)\s+(\S+)\s+"(.*)$'   # ブロック: BL x y "名前(基準点+名前で比較)
  $rxSkip = [regex]'^(lg|ly|lc|lt|pn|hq|hd|hs|hcw|hch|hcd|hcc|hn|hp|hk|hz|msg|msz)'
  for($li=$geomStart; $li -lt $lines.Count; $li++){
    $t = $lines[$li].Trim()
    if($t.Length -eq 0 -or $t -eq "#"){ continue }
    $c0 = $t[0]
    if($c0 -eq 'c'){
      if($t.Length -ge 2 -and $t[1] -eq 'n'){ $curCn = $t; continue }
      $m = $rxText.Match($t)
      if($m.Success){
        $ok=$true; $nums = [double[]]::new(4)
        for($j=0;$j -lt 4;$j++){ $d=0.0; if([double]::TryParse($m.Groups[$j+2].Value,$NS,$inv,[ref]$d)){ $nums[$j]=$d } else { $ok=$false; break } }
        if($ok){ $newEls.Add(@{Type=$m.Groups[1].Value;Nums=$nums;Text=$m.Groups[6].Value;Extra=$curCn;Ints=$null;Key=$null}); continue }
      }
    }
    if($rxSkip.IsMatch($t)){ continue }
    if($c0 -eq 'B'){
      $m = $rxBlock.Match($t)
      if($m.Success){
        $bx=0.0; $by=0.0
        if([double]::TryParse($m.Groups[1].Value,$NS,$inv,[ref]$bx) -and [double]::TryParse($m.Groups[2].Value,$NS,$inv,[ref]$by)){
          $newEls.Add(@{Type="bl";Nums=[double[]]@($bx,$by);Text=$m.Groups[3].Value;Extra="";Ints=$null;Key=$null}); continue
        }
      }
    }
    $tok = $t -split '\s+'
    $head = $tok[0]
    $isNumHead = ($c0 -eq '-' -or $c0 -eq '.' -or [char]::IsDigit($c0))
    if($isNumHead -or $head -eq "ci" -or $head -eq "pt" -or $head -eq "sl"){
      $from = if($isNumHead){0}else{1}
      $cnt = $tok.Count-$from
      if(($isNumHead -and $cnt -eq 4) -or (-not $isNumHead -and $cnt -ge 2)){
        $ok=$true; $nums = [double[]]::new($cnt)
        for($j=0;$j -lt $cnt;$j++){ $d=0.0; if([double]::TryParse($tok[$j+$from],$NS,$inv,[ref]$d)){ $nums[$j]=$d } else { $ok=$false; break } }
        if($ok){ $newEls.Add(@{Type=$(if($isNumHead){"ln"}else{$head});Nums=$nums;Text="";Extra="";Ints=$null;Key=$null}); continue }
      }
    }
    # 未知書式: 生行で比較(数値のみ量子化)
    $unknown++
    $newEls.Add(@{Type="raw";Nums=@();Text="";Extra=$t;Ints=$null;Key=$null})
  }
  Log("図形部行数=$($lines.Count-$geomStart)  要素数=$($newEls.Count)  未知書式=$unknown  ($([int]$sw.ElapsedMilliseconds)ms)")

  # ================================================================
  if($Mode -eq "save"){
    $bl = New-Object 'System.Collections.Generic.List[string]'
    $bl.Add("# jw_diff base`t"+(Get-Date -Format "yyyy-MM-dd HH:mm:ss")+"`t"+$newEls.Count+"`t"+$VERSION)
    foreach($e in $newEls){
      $ns = [string[]]::new($e.Nums.Count)
      for($j=0;$j -lt $ns.Count;$j++){ $ns[$j] = ([double]$e.Nums[$j]).ToString("0.######",$inv) }
      $bl.Add($e.Type + "`t" + ($ns -join ' ') + "`t" + $e.Text + "`t" + $e.Extra)
    }
    [IO.File]::WriteAllLines($basePath,$bl,$utf8bom)
    Log("保存: $basePath  要素数=$($newEls.Count)  ($([int]$sw.ElapsedMilliseconds)ms)")
    [IO.File]::WriteAllLines($tp,@(("he旧図を保存しました: "+$newEls.Count+"要素 ("+[IO.Path]::GetFileName($basePath)+")")),$enc)
    return
  }

  # ================================================================
  # compare
  if(-not (Test-Path $basePath)){
    Log("base が見つかりません: $basePath")
    [IO.File]::WriteAllLines($tp,@("he旧図が未保存です。先に「差分_旧図保存」を実行してください"),$enc)
    return
  }
  $bLines = [IO.File]::ReadAllLines($basePath,$utf8bom)
  $baseMeta = ""; $oldEls = New-Object 'System.Collections.Generic.List[object]'
  foreach($s in $bLines){
    if($s.Length -eq 0){ continue }
    if($s[0] -eq '#'){ $mf = $s -split "`t"; if($mf.Count -ge 3){ $baseMeta = "$($mf[1])保存, $($mf[2])要素" }; continue }
    $f = $s -split "`t",4
    if($f.Count -lt 4){ continue }
    $nums = @()
    if($f[1].Length -gt 0){
      $tok = $f[1] -split ' '; $nums = [double[]]::new($tok.Count); $ok=$true
      for($j=0;$j -lt $tok.Count;$j++){ $d=0.0; if([double]::TryParse($tok[$j],$NS,$inv,[ref]$d)){ $nums[$j]=$d } else { $ok=$false; break } }
      if(-not $ok){ continue }
    }
    $oldEls.Add(@{Type=$f[0];Nums=$nums;Text=$f[2];Extra=$f[3];Ints=$null;Key=$null})
  }
  Log("base: $basePath ($baseMeta)  読込要素=$($oldEls.Count)  ($([int]$sw.ElapsedMilliseconds)ms)")

  # --- 量子化キー + 一次照合 (平行移動補正後に再実行できるよう関数化。中身はインライン) ---
  function Invoke-Match($oldEls,$newEls){
    # --- 量子化キー(インライン)と キー→要素リスト の多重集合 ---
    #   key = Type|q1,q2,...|tail   q = floor(v/Tol+0.5)
    #   ln: 端点を辞書順ソート / sl: 頂点を辞書順ソート / 文字: 基準点 + 単位方向(1/1000量子化) + 文字列 / raw: 数値だけ量子化
    $oldIdx = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $newIdx = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    foreach($set in @(@($oldEls,$oldIdx),@($newEls,$newIdx))){
      $els=$set[0]; $dict=$set[1]
      foreach($e in $els){
        $n=$e.Nums; $ty=$e.Type; $tail=""
        if($ty -eq "ln"){
            # 端点は「量子化後の整数」で辞書順ソート(生値でソートすると x がほぼ同じ縦線で
            # 桁落ちにより順序が入れ替わり、base と新図でキーが変わる)
            $ax=[long][Math]::Floor($n[0]/$Tol+0.5); $ay=[long][Math]::Floor($n[1]/$Tol+0.5)
            $bx=[long][Math]::Floor($n[2]/$Tol+0.5); $by=[long][Math]::Floor($n[3]/$Tol+0.5)
            if(($ax -gt $bx) -or ($ax -eq $bx -and $ay -gt $by)){ $ints=@($bx,$by,$ax,$ay) } else { $ints=@($ax,$ay,$bx,$by) }
        } elseif($ty -eq "sl"){
            # 頂点も量子化後の整数で辞書順ソート(挿入ソート・頂点は3〜4個)
            $np=[Math]::Floor($n.Count/2); $qx=[long[]]::new($np); $qy=[long[]]::new($np)
            for($i=0;$i -lt $np;$i++){ $qx[$i]=[long][Math]::Floor($n[2*$i]/$Tol+0.5); $qy[$i]=[long][Math]::Floor($n[2*$i+1]/$Tol+0.5) }
            for($i=1;$i -lt $np;$i++){ $tx=$qx[$i]; $ty2=$qy[$i]; $j=$i-1
              while($j -ge 0 -and (($qx[$j] -gt $tx) -or ($qx[$j] -eq $tx -and $qy[$j] -gt $ty2))){ $qx[$j+1]=$qx[$j]; $qy[$j+1]=$qy[$j]; $j-- }
              $qx[$j+1]=$tx; $qy[$j+1]=$ty2 }
            $ints=[long[]]::new(2*$np); for($i=0;$i -lt $np;$i++){ $ints[2*$i]=$qx[$i]; $ints[2*$i+1]=$qy[$i] }
        } elseif($ty -eq "ch" -or $ty -eq "cv" -or $ty -eq "cs" -or $ty -eq "cc"){
            $dx=$n[2]; $dy=$n[3]; $len=[Math]::Sqrt($dx*$dx+$dy*$dy)
            if($len -gt 1e-9){ $ux=$dx/$len; $uy=$dy/$len } else { $ux=1.0; $uy=0.0 }
            $ints=@([long][Math]::Floor($n[0]/$Tol+0.5),[long][Math]::Floor($n[1]/$Tol+0.5),[long][Math]::Floor($ux*1000+0.5),[long][Math]::Floor($uy*1000+0.5))
            $tail=$e.Text
        } elseif($ty -eq "bl"){
          $ints=@([long][Math]::Floor($n[0]/$Tol+0.5),[long][Math]::Floor($n[1]/$Tol+0.5)); $tail=$e.Text
        } elseif($ty -eq "raw"){
            $ints=@(); $tt=@()
            foreach($k in ($e.Extra -split '\s+')){ $d=0.0; if([double]::TryParse($k,$NS,$inv,[ref]$d)){ $ints += [long][Math]::Floor($d/$Tol+0.5) } else { $tt += $k } }
            $tail = $tt -join ' '
        } else { $ints = [long[]]::new($n.Count); for($j=0;$j -lt $n.Count;$j++){ $ints[$j]=[long][Math]::Floor($n[$j]/$Tol+0.5) } }
        $k = $ty + "|" + ($ints -join ",") + "|" + $tail
        $e.Ints=$ints; $e.Key=$k
        if($dict.ContainsKey($k)){ $dict[$k].Add($e) } else { $l = [System.Collections.Generic.List[object]]::new(); $l.Add($e); $dict[$k]=$l }
      }
    }
    Log("キー生成 ($([int]$sw.ElapsedMilliseconds)ms)")
  
    $deleted = New-Object 'System.Collections.Generic.List[object]'
    $added   = New-Object 'System.Collections.Generic.List[object]'
    $matched = 0
    foreach($k in $oldIdx.Keys){
      $ol = $oldIdx[$k]
      if($newIdx.ContainsKey($k)){
        $nl = $newIdx[$k]; $m=[Math]::Min($ol.Count,$nl.Count); $matched += $m
        for($i=$m;$i -lt $ol.Count;$i++){ $deleted.Add($ol[$i]) }
        for($i=$m;$i -lt $nl.Count;$i++){ $added.Add($nl[$i]) }
      } else { foreach($e in $ol){ $deleted.Add($e) } }
    }
    foreach($k in $newIdx.Keys){ if(-not $oldIdx.ContainsKey($k)){ foreach($e in $newIdx[$k]){ $added.Add($e) } } }
    Log("一次照合: 一致=$matched 削除候補=$($deleted.Count) 追加候補=$($added.Count)  ($([int]$sw.ElapsedMilliseconds)ms)")
  return @{Matched=$matched;Deleted=$deleted;Added=$added}
  }

  $r = Invoke-Match $oldEls $newEls
  $matched=$r.Matched; $deleted=$r.Deleted; $added=$r.Added

  # --- 平行移動の自動検出・補正 ---
  #   未一致が旧図の半数を超えるとき、未一致同士で「同じ文字列の文字 / 同じ長さ・角度の線 / 同じ半径の円」が
  #   1対1 で対応する組の位置差を投票し、支配的なずれがあれば旧図全体をその分だけ動かして再照合する。
  $shiftDx=0.0; $shiftDy=0.0; $shiftInfo=""
  if(-not $NoShift -and $deleted.Count -gt 0 -and $added.Count -gt 0 -and ($deleted.Count*2) -gt $oldEls.Count){
    $sigO = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $sigN = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    foreach($set in @(@($deleted,$sigO),@($added,$sigN))){
      $els=$set[0]; $dict=$set[1]
      foreach($e in $els){
        $n=$e.Nums; $ty=$e.Type; $sig=$null
        if($ty -eq "ln"){
          $dx=$n[2]-$n[0]; $dy=$n[3]-$n[1]; $len=[Math]::Sqrt($dx*$dx+$dy*$dy)
          if($len -lt 20){ continue }
          $ang=[Math]::Atan2($dy,$dx)*180/[Math]::PI; $ang=(($ang%180)+180)%180
          $sig="L|"+[Math]::Round($len)+"|"+[Math]::Round($ang*2); $px=($n[0]+$n[2])/2; $py=($n[1]+$n[3])/2
        } elseif($ty -eq "ch" -or $ty -eq "cv" -or $ty -eq "cs" -or $ty -eq "cc"){
          $sig="T|"+$e.Text+"|"+$e.Ints[2]+","+$e.Ints[3]; $px=$n[0]; $py=$n[1]
        } elseif($ty -eq "bl"){
          $sig="B|"+$e.Text; $px=$n[0]; $py=$n[1]
        } elseif($ty -eq "ci" -and $n.Count -ge 3){
          $sig="C|"+[Math]::Round($n[2]); $px=$n[0]; $py=$n[1]
        } else { continue }
        $rec=@($px,$py)
        if($dict.ContainsKey($sig)){ $dict[$sig].Add($rec) } else { $l=[System.Collections.Generic.List[object]]::new(); $l.Add($rec); $dict[$sig]=$l }
      }
    }
    # 1対1 の組だけ採用し、1mm セルで投票
    $votes = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $samples=0
    foreach($sig in $sigO.Keys){
      if(-not $sigN.ContainsKey($sig)){ continue }
      $lo=$sigO[$sig]; $ln=$sigN[$sig]
      if($lo.Count -ne 1 -or $ln.Count -ne 1){ continue }
      $ddx=$ln[0][0]-$lo[0][0]; $ddy=$ln[0][1]-$lo[0][1]
      if([Math]::Abs($ddx) -gt 100000 -or [Math]::Abs($ddy) -gt 100000){ continue }
      $samples++
      $vk=([Math]::Round($ddx)).ToString()+","+([Math]::Round($ddy)).ToString()
      if($votes.ContainsKey($vk)){ $votes[$vk].Add(@($ddx,$ddy)) } else { $l=[System.Collections.Generic.List[object]]::new(); $l.Add(@($ddx,$ddy)); $votes[$vk]=$l }
    }
    $best=$null; $bestN=0
    foreach($vk in $votes.Keys){ if($votes[$vk].Count -gt $bestN){ $bestN=$votes[$vk].Count; $best=$vk } }
    Log("平行移動検出: 対応候補=$samples 最多セル=$best 票=$bestN")
    if($null -ne $best -and $bestN -ge 10 -and ($bestN*10) -ge ($samples*3)){
      # 最多セルと隣接セルの平均で精密化
      $bp=$best -split ','; $cx=[int]$bp[0]; $cy=[int]$bp[1]; $sx=0.0; $sy=0.0; $cnt=0
      for($ox=-1;$ox -le 1;$ox++){ for($oy=-1;$oy -le 1;$oy++){ $vk=($cx+$ox).ToString()+","+($cy+$oy).ToString(); if($votes.ContainsKey($vk)){ foreach($v in $votes[$vk]){ $sx+=$v[0]; $sy+=$v[1]; $cnt++ } } } }
      $shiftDx=$sx/$cnt; $shiftDy=$sy/$cnt
      if([Math]::Abs($shiftDx) -ge $Tol/2 -or [Math]::Abs($shiftDy) -ge $Tol/2){
        # 旧図全体を補正(座標のみ。文字の方向ベクトルは不変)
        foreach($e in $oldEls){
          $n=$e.Nums; $ty=$e.Type
          if($ty -eq "ln" -or $ty -eq "sl"){ for($j=0;$j+1 -lt $n.Count;$j+=2){ $n[$j]+=$shiftDx; $n[$j+1]+=$shiftDy } }
          elseif($ty -eq "raw"){
            $tk=$e.Extra -split ' '; $ni=0
            for($j=0;$j -lt $tk.Count;$j++){ $d=0.0; if([double]::TryParse($tk[$j],$NS,$inv,[ref]$d)){ if($ni -lt 2){ $d += $(if($ni -eq 0){$shiftDx}else{$shiftDy}); $tk[$j]=$d.ToString("0.######",$inv) }; $ni++ } }
            $e.Extra=($tk -join ' ')
          }
          elseif($n.Count -ge 2){ $n[0]+=$shiftDx; $n[1]+=$shiftDy }
        }
        $shiftInfo="平行移動補正 dx="+(P1 $shiftDx)+" dy="+(P1 $shiftDy)+" (票 $bestN/$samples)"
        Log("$shiftInfo → 旧図を補正して再照合(以下の旧図座標は補正後)")
        $r = Invoke-Match $oldEls $newEls
        $matched=$r.Matched; $deleted=$r.Deleted; $added=$r.Added
      } else { $shiftDx=0.0; $shiftDy=0.0 }
    }
  }

  # --- 境界すり抜け救済: 未マッチ集合内だけ、同種別・同文字列で全座標セル差が ±1 以内なら一致扱い ---
  #     追加側を「Type|tail|先頭点の粗バケット」で索引し、削除側は近傍9バケットだけ走査する
  $rescued = 0
  if($deleted.Count -gt 0 -and $added.Count -gt 0){
    $B = 64
    $bk = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    foreach($e in $added){
      $ii=$e.Ints; if($ii.Count -lt 2){ continue }
      $tail = if($e.Type -eq "raw"){ $e.Key.Substring($e.Key.LastIndexOf('|')+1) } else { $e.Text }
      $gk = $e.Type+"|"+$tail+"|"+[Math]::Floor($ii[0]/$B)+","+[Math]::Floor($ii[1]/$B)
      if($bk.ContainsKey($gk)){ $bk[$gk].Add($e) } else { $l = [System.Collections.Generic.List[object]]::new(); $l.Add($e); $bk[$gk]=$l }
    }
    $keepDel = New-Object 'System.Collections.Generic.List[object]'
    foreach($e in $deleted){
      $ii=$e.Ints; $found=$null
      if($ii.Count -ge 2){
        $tail = if($e.Type -eq "raw"){ $e.Key.Substring($e.Key.LastIndexOf('|')+1) } else { $e.Text }
        # 探索起点: 先頭点。線分/ソリッドは端点(頂点)ソートが境界で入れ替わる場合に備え
        # 各頂点を起点にし、点の並び順に依らない比較(点集合として ±1 セル)を行う
        $isPts = ($e.Type -eq "ln" -or $e.Type -eq "sl")
        $probes = @()
        if($isPts){ for($pi=0;$pi+1 -lt $ii.Count;$pi+=2){ $probes += ,@($ii[$pi],$ii[$pi+1]) } } else { $probes += ,@($ii[0],$ii[1]) }
        foreach($pr in $probes){
          if($null -ne $found){ break }
          $bx=[Math]::Floor($pr[0]/$B); $by=[Math]::Floor($pr[1]/$B)
          for($ox=-1;$ox -le 1 -and $null -eq $found;$ox++){ for($oy=-1;$oy -le 1 -and $null -eq $found;$oy++){
            $gk = $e.Type+"|"+$tail+"|"+($bx+$ox)+","+($by+$oy)
            if(-not $bk.ContainsKey($gk)){ continue }
            $cands = $bk[$gk]
            for($ci=0;$ci -lt $cands.Count;$ci++){
              $c=$cands[$ci]; $cj=$c.Ints
              if($cj.Count -ne $ii.Count){ continue }
              $near=$true
              if($isPts){
                # e の各点に対し、未使用の c の点で ±1 セル以内のものを貪欲に割り当てる
                $np=[Math]::Floor($ii.Count/2); $used=[bool[]]::new($np)
                for($pa=0;$pa -lt $np -and $near;$pa++){
                  $hit=$false
                  for($pb=0;$pb -lt $np;$pb++){
                    if($used[$pb]){ continue }
                    if([Math]::Abs($ii[2*$pa]-$cj[2*$pb]) -le 1 -and [Math]::Abs($ii[2*$pa+1]-$cj[2*$pb+1]) -le 1){ $used[$pb]=$true; $hit=$true; break }
                  }
                  if(-not $hit){ $near=$false }
                }
              } else {
                for($j=0;$j -lt $ii.Count;$j++){ if([Math]::Abs($ii[$j]-$cj[$j]) -gt 1){ $near=$false; break } }
              }
              if($near){ $found=$c; $cands.RemoveAt($ci); break }
            }
          } }
        }
      }
      if($null -ne $found){ $rescued++; $matched++; Log("境界救済: "+(Describe $e)+"  ~  "+(Describe $found)) }
      else { $keepDel.Add($e) }
    }
    $deleted = $keepDel
    $keepAdd = New-Object 'System.Collections.Generic.List[object]'
    foreach($e in $added){ if($e.Ints.Count -lt 2){ $keepAdd.Add($e) } }
    foreach($gk in $bk.Keys){ foreach($e in $bk[$gk]){ $keepAdd.Add($e) } }
    $added = $keepAdd
    Log("境界救済 ($([int]$sw.ElapsedMilliseconds)ms)")
  }

  Log("新図要素: $($newEls.Count) / 一致: $matched / 削除: $($deleted.Count) / 追加: $($added.Count) / 境界救済: $rescued / 未知書式: $unknown"+$(if($shiftInfo){" / "+$shiftInfo}else{""}))
  Log("--- 削除 ---")
  foreach($e in $deleted){ Log((Describe $e)) }
  Log("--- 追加 ---")
  foreach($e in $added){ Log((Describe $e)) }

  # --- 書き戻し ---
  $out = New-Object 'System.Collections.Generic.List[string]'
  if($deleted.Count -eq 0 -and $added.Count -eq 0){
    $out.Add("he差分なし: 新図"+$newEls.Count+"要素 / 旧図"+$oldEls.Count+"要素 が一致しました")
    [IO.File]::WriteAllLines($tp,$out,$enc)
    return
  }
  if($Layer -ne ""){ $out.Add("lg"+$Layer.ToUpper()) }
  $marks = 0; $truncated = $false
  $rStr = P $MarkR
  if($deleted.Count -gt 0){
    $out.Add("lc$ColorDel")
    foreach($e in $deleted){
      if($marks -ge $MaxMark){ $truncated=$true; break }
      if(-not $NoRedraw){ foreach($l in (GeomLines $e)){ $out.Add($l) } }
      $rp = RepPoint $e
      if($null -ne $rp){ $out.Add(("ci "+(P $rp[0])+" "+(P $rp[1])+" "+$rStr)) }
      $marks++
    }
  }
  if($added.Count -gt 0 -and -not $truncated){
    $out.Add("lc$ColorAdd")
    foreach($e in $added){
      if($marks -ge $MaxMark){ $truncated=$true; break }
      $rp = RepPoint $e
      if($null -ne $rp){ $out.Add(("ci "+(P $rp[0])+" "+(P $rp[1])+" "+$rStr)) }
      $marks++
    }
  }
  if($truncated){
    $msg = "差分 "+($deleted.Count+$added.Count)+"件中 "+$marks+"件のみ作図(全件は jw_diff_log.txt 参照)"
    if($null -ne $hn -and $hn.Count -ge 4){
      $rx=[Math]::Max($hn[0],$hn[2]); $ry=[Math]::Max($hn[1],$hn[3])
      $tw = $msg.Length*$charW
      $out.Add("cn3")
      $out.Add(("ch "+(P ($rx-$tw))+" "+(P ($ry+$charH))+" "+(P $tw)+" 0 `""+$msg))
    }
    Log("マーカー上限 $MaxMark で打ち切り: $msg")
  }
  [IO.File]::WriteAllLines($tp,$out,$enc)
  Log("出力行=$($out.Count)  マーカー=$marks  ($([int]$sw.ElapsedMilliseconds)ms)")
}
catch {
  try { if($tp){ [IO.File]::WriteAllLines($tp,@(("heエラー: "+$_.Exception.Message)),$enc) } } catch {}
  Log("例外: "+$_.Exception.Message+"  "+$_.ScriptStackTrace)
}
finally {
  try { [IO.File]::WriteAllLines($logPath,$log,$utf8bom) } catch {}
}
