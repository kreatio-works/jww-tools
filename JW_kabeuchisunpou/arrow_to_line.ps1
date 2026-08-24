# kreatio-works  https://kreatio-works.github.io/
<#
  arrow_to_line.ps1  -- 塗り矢印(三角ソリッド) → 線矢印(2本線) 変換 外部変形
  役割:
    範囲選択されたデータのうち「3点ソリッド(sl x1 y1 x2 y2 x3 y3)」で矢印らしい
    細長い三角形だけを、頂点→底辺両端の2本線(開き矢印)に置き換える。
    それ以外の図形・属性行はすべてそのまま書き戻す(パススルー)。
  前提:
    - 元図形の削除は「出力 jwc_temp.txt の先頭に hd 行を書く」ことで行う。
      (bat の REM #hd は存在しない指令＝無視され、追加コピーになる。実機検証済みの一次情報)
      hd は選択図形を全部削除するため、変換対象以外も必ず全行書き戻すこと(書き戻し漏れ=消失)。
    - #bz なし = 座標は実寸そのまま。
    - jwc_temp.txt は Shift-JIS(CP932)。
  安全策:
    - 変換対象ゼロ/例外/ブロック検出時は hd を出さず he メッセージのみ → 図面は無変更。
  矢印判定:
    - 3点ソリッドで、最短辺/最長辺 <= RatioMax (正三角形などを除外)
    - MaxSide > 0 のときは最長辺がそれ以下のものだけ
    - 頂点 = 最短辺(底辺)の対頂点。線2本 = 頂点→底辺の両端
#>
param(
  [string]$TempFile = "jwc_temp.txt",
  [double]$RatioMax = 0.95,  # 最短辺/最長辺 がこれ以下なら矢印とみなす(正三角形1.0を除外)
  [double]$MaxSide  = 0      # >0 のとき: 最長辺がこの実寸mm以下の三角形だけ変換(誤変換防止)
)

$ErrorActionPreference = "Stop"
$inv = [Globalization.CultureInfo]::InvariantCulture
$enc = [Text.Encoding]::GetEncoding(932)
function P([double]$v){ return $v.ToString("0.######",$inv) }

$scriptDir = if($PSScriptRoot){$PSScriptRoot}else{(Get-Location).Path}
$logPath = Join-Path $scriptDir "arrow_to_line_log.txt"
$log = New-Object System.Collections.Generic.List[string]
function Log($m){ $log.Add(((Get-Date -Format "HH:mm:ss") + "  " + $m)) }

$tp = $null
try {
  $tp = $TempFile
  if(-not (Test-Path $tp)){ $alt = Join-Path $scriptDir "jwc_temp.txt"; if(Test-Path $alt){ $tp = $alt } }
  if(-not (Test-Path $tp)){ throw "jwc_temp.txt が見つかりません" }
  $tp = (Resolve-Path $tp).Path
  $lines = [IO.File]::ReadAllLines($tp,$enc)
  Log("読込: $tp  行数=$($lines.Count)")

  # 設定部/図形部を最初の '#' で分割
  $hashIdx = -1
  for($i=0;$i -lt $lines.Count;$i++){ if($lines[$i].Trim() -eq "#"){ $hashIdx=$i; break } }
  if($hashIdx -lt 0){ throw "区切り '#' が見つかりません" }
  $geom = if($hashIdx -lt $lines.Count-1){ $lines[($hashIdx+1)..($lines.Count-1)] } else { @() }

  # ブロックが混ざると往復で崩れる恐れ → 中断(元図形は削除されない想定。実機初回は要確認)
  foreach($g in $geom){ if($g.Trim() -match '^(bl|BL)\b'){
    [IO.File]::WriteAllLines($tp,@("heブロックを含む範囲は変換できません。ブロックを避けて選択してください"),$enc)
    Log("中断: ブロック検出"); throw "block-abort"
  } }

  $out = New-Object System.Collections.Generic.List[string]
  $conv=0; $keptSl=0
  foreach($g in $geom){
    $t=$g.Trim()
    # 選択枠・指示点などの制御行は書き戻さない
    if($t -match '^(hn|hp\d|hq)\b' -or $t -match '^hp\d'){ continue }
    if($t -match '^sl\s+(.+)$'){
      $tok = $matches[1].Trim() -split '\s+'
      $ok = ($tok.Count -eq 6)
      $n=@()
      if($ok){ foreach($k in $tok){ $d=0.0; if([double]::TryParse($k,[Globalization.NumberStyles]::Float,$inv,[ref]$d)){ $n+=$d } else { $ok=$false; break } } }
      if($ok){
        # 3辺: s12(対頂点=3), s23(対頂点=1), s13(対頂点=2)
        $vx=@($n[0],$n[2],$n[4]); $vy=@($n[1],$n[3],$n[5])
        function d2($a,$b){ [Math]::Sqrt(($vx[$a]-$vx[$b])*($vx[$a]-$vx[$b])+($vy[$a]-$vy[$b])*($vy[$a]-$vy[$b])) }
        $s=@((d2 0 1),(d2 1 2),(d2 0 2))   # 対頂点 index: 2, 0, 1
        $apexOf=@(2,0,1)
        $minI=0; $maxI=0
        for($k=1;$k -lt 3;$k++){ if($s[$k] -lt $s[$minI]){$minI=$k}; if($s[$k] -gt $s[$maxI]){$maxI=$k} }
        $isArrow = ($s[$maxI] -gt 1e-6) -and (($s[$minI]/$s[$maxI]) -le $RatioMax) -and ($MaxSide -le 0 -or $s[$maxI] -le $MaxSide)
        if($isArrow){
          $ap=$apexOf[$minI]
          $others = @(0,1,2) | Where-Object { $_ -ne $ap }
          foreach($o in $others){
            $out.Add(((P $vx[$ap])+" "+(P $vy[$ap])+" "+(P $vx[$o])+" "+(P $vy[$o])))
          }
          $conv++
          continue
        }
        $keptSl++
      }
    }
    $out.Add($g)   # それ以外は原文のまま書き戻し
  }
  Log("変換した三角=$conv  変換しなかったソリッド=$keptSl  書き戻し行=$($out.Count)")

  if($conv -eq 0){
    # 変換対象なし → hd を出さずメッセージのみ(図面は無変更)
    [IO.File]::WriteAllLines($tp,@("he変換対象の塗り矢印(三角ソリッド)が見つかりませんでした"),$enc)
    Log("変換対象なし: he のみ出力(無変更)")
  } else {
    # 先頭の hd 行 = 選択元データを削除して置換(移動)。これが無いと追加コピーになる
    $out.Insert(0,"hd")
    [IO.File]::WriteAllLines($tp,$out,$enc)
  }
}
catch {
  if($_.Exception.Message -ne "block-abort" -and $tp -and (Test-Path $tp)){
    # hd を出していないので元図形は削除されない。エラー表示のみ
    try { [IO.File]::WriteAllLines($tp,@(("heエラー: "+$_.Exception.Message)),$enc) } catch {}
  }
  Log("例外: "+$_.Exception.Message)
}
finally {
  try { [IO.File]::WriteAllLines($logPath,$log,(New-Object System.Text.UTF8Encoding($true))) } catch {}
}
