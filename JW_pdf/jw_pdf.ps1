# Jw_cad 図面 一括PDF化
#   Jw_cad で1枚ずつ開き、CubePDF に印刷して、図面ごとに1つの PDF を作る。
#   使い方: 「JWW一括PDF化.bat」をダブルクリック、または .jww / フォルダをドロップ。
#   画面のクリックやキー入力は使わず、Jw_cad と CubePDF のボタンへ直接命令を送る。
#   そのため変換中に他の作業をしても乱れにくい（Jw_cad の窓は出たり消えたりする）。
param([Parameter(Position = 0, ValueFromRemainingArguments = $true)][string[]]$Paths, [switch]$NoGui, [string]$Paper = '', [string]$OutDir = '', [string]$Color = '')

$ErrorActionPreference = 'Stop'
$Here = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogFile = Join-Path $Here 'jw_pdf_log.txt'
$SettingFile = Join-Path $Here 'jw_pdf_settings.json'
$PrinterName = 'CubePDF'

Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# ---------- Win32 ----------
Add-Type -ReferencedAssemblies System.Windows.Forms @"
using System; using System.Text; using System.Runtime.InteropServices; using System.Collections.Generic;
public static class JwW {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc p, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr h, EnumProc p, IntPtr l);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool IsWindowEnabled(IntPtr h);
  [DllImport("user32.dll")] public static extern int GetDlgCtrlID(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetDlgItem(IntPtr h, int id);
  [DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr h);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, string l);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, StringBuilder l);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern IntPtr SendMessageTimeout(IntPtr h, uint m, IntPtr w, StringBuilder l, uint f, uint t, out IntPtr r);
  public static string Text(IntPtr h){ var s=new StringBuilder(1024); GetWindowText(h,s,1024); if(s.Length>0) return s.ToString(); IntPtr r; s.Clear(); if(SendMessageTimeout(h,0x0D,(IntPtr)1024,s,2,2000,out r)==IntPtr.Zero) s.Clear(); return s.ToString(); }
  public static string Cls(IntPtr h){ var s=new StringBuilder(256); GetClassName(h,s,256); return s.ToString(); }
  public static IntPtr[] Tops(int pid){ var r=new List<IntPtr>(); EnumWindows((h,l)=>{ uint p; GetWindowThreadProcessId(h,out p); if(p==(uint)pid && IsWindowVisible(h)) r.Add(h); return true; }, IntPtr.Zero); return r.ToArray(); }
  public static IntPtr[] Kids(IntPtr w){ var r=new List<IntPtr>(); EnumChildWindows(w,(h,l)=>{ r.Add(h); return true; }, IntPtr.Zero); return r.ToArray(); }
  public static string ComboText(IntPtr cb, int i){ var s=new StringBuilder(512); SendMessage(cb,0x148,(IntPtr)i,s); return s.ToString(); }
}
"@

$WM_COMMAND = 0x111; $WM_CLOSE = 0x10; $WM_SETTEXT = 0x0C; $BM_CLICK = 0xF5; $BM_GETCHECK = 0xF0
$CB_GETCOUNT = 0x146; $CB_GETCURSEL = 0x147; $CB_SETCURSEL = 0x14E
$ID_FILE_PRINT = 57607

# Jw_cad の用紙コード（図面ファイルの先頭に入っている）→ 用紙名
$ZumenName = @{ 0='A0'; 1='A1'; 2='A2'; 3='A3'; 4='A4'; 8='2A'; 9='3A'; 10='4A'; 11='5A'; 12='10m'; 13='50m'; 14='100m' }
# A3 にそろえる時の印刷倍率（Jw_cad の倍率リストの何番目か）
#   0:200% 1:141% 2:100% 3:71% 4:50% 5:35% 6:25%
$ToA3Scale = @{ 'A0'=5; 'A1'=4; 'A2'=3; 'A3'=2 }

function Write-Log([string]$msg) {
  $line = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '  ' + $msg
  Add-Content -Path $LogFile -Value $line -Encoding UTF8
}

function Pump { [System.Windows.Forms.Application]::DoEvents() }

# 条件が満たされるまで待つ（画面が固まらないよう DoEvents しながら）
function Wait-Until([scriptblock]$cond, [int]$timeoutSec, [int]$stepMs = 150) {
  $end = (Get-Date).AddSeconds($timeoutSec)
  while ((Get-Date) -lt $end) {
    $r = & $cond
    if ($r) { return $r }
    Pump; Start-Sleep -Milliseconds $stepMs
  }
  return $null
}

function Find-JwExe {
  $cands = @()
  try {
    $prog = (Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\.jww' -ErrorAction Stop).'(default)'
    if ($prog) {
      $cmd = (Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\$prog\shell\open\command" -ErrorAction Stop).'(default)'
      if ($cmd -match '^"([^"]+)"' -or $cmd -match '^(\S+)') { $cands += $Matches[1] }
    }
  } catch {}
  $cands += 'C:\jww\Jw_win.exe', 'C:\Program Files\jww\Jw_win.exe', 'C:\Program Files (x86)\jww\Jw_win.exe'
  foreach ($c in $cands) { if ($c -and (Test-Path $c)) { return $c } }
  return $null
}

# 図面ファイルの先頭から用紙サイズを読む（JwwData. / 版 / メモ / 用紙）
function Get-JwwPaper([string]$path) {
  $fs = [IO.File]::OpenRead($path)
  try {
    $br = New-Object IO.BinaryReader($fs)
    $sig = [Text.Encoding]::ASCII.GetString($br.ReadBytes(8))
    if ($sig -ne 'JwwData.') { return $null }
    [void]$br.ReadUInt32()
    $n = [int]$br.ReadByte()
    if ($n -eq 0xFF) { $n = [int]$br.ReadUInt16(); if ($n -eq 0xFFFF) { $n = [int]$br.ReadUInt32() } }
    [void]$br.ReadBytes($n)
    $z = [int]$br.ReadUInt32()
    if ($ZumenName.ContainsKey($z)) { return $ZumenName[$z] } else { return "?$z" }
  } finally { $fs.Close() }
}

function Find-Dialog([int]$procId, [string]$title) {
  foreach ($h in [JwW]::Tops($procId)) {
    if ([JwW]::Cls($h) -eq '#32770' -and [JwW]::Text($h) -eq $title) { return $h }
  }
  return $null
}

# 想定外のメッセージ箱（読込エラーなど）を探す
function Find-OtherDialog([int]$procId) {
  foreach ($h in [JwW]::Tops($procId)) {
    if ([JwW]::Cls($h) -eq '#32770' -and [JwW]::Text($h) -ne 'プリンターの設定') {
      $msg = ([JwW]::Kids($h) | Where-Object { [JwW]::Cls($_) -eq 'Static' } | ForEach-Object { [JwW]::Text($_) }) -join ' '
      if ($msg -match '印刷中') { continue }   # 印刷の進み具合の表示は無視
      if (-not ($msg.Trim()) -and -not ([JwW]::Text($h))) { continue }
      return @{ Handle = $h; Text = ([JwW]::Text($h) + ' ' + $msg).Trim() }
    }
  }
  return $null
}

# ---------- Jw_cad の設定を元に戻す ----------
# Jw_cad は終了時に設定（最新のファイル一覧・カラー印刷・用紙など）をレジストリへ書く。
# 変換のために開いた Jw_cad がふだんの設定を変えないよう、起動前の値を控えて、終了後に変わった所だけ戻す。
$JwRegPath = 'Software\Jw_cad'
function Get-RegSnapshot {
  $snap = @{}
  $root = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($JwRegPath)
  if (-not $root) { return $snap }
  $stack = New-Object System.Collections.Stack
  $stack.Push(@($root, $JwRegPath))
  while ($stack.Count -gt 0) {
    $pair = $stack.Pop(); $key = $pair[0]; $path = $pair[1]
    foreach ($n in $key.GetValueNames()) {
      $snap["$path|$n"] = @($key.GetValueKind($n), $key.GetValue($n, $null, 'DoNotExpandEnvironmentNames'))
    }
    foreach ($s in $key.GetSubKeyNames()) { $sk = $key.OpenSubKey($s); if ($sk) { $stack.Push(@($sk, "$path\$s")) } }
    $key.Close()
  }
  return $snap
}
function Get-RegValueText($v) { if ($v -is [array]) { return ($v -join ',') } else { return [string]$v } }
function Restore-RegSnapshot($before) {
  if (-not $before -or $before.Count -eq 0) { return }
  $after = Get-RegSnapshot
  $setVal = {
    param($id, $kind, $val)
    $path, $name = $id.Split('|', 2)
    $k = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($path)
    try { $k.SetValue($name, $val, $kind) } finally { $k.Close() }
  }
  foreach ($id in @($after.Keys)) {
    if (-not $before.ContainsKey($id)) {
      $path, $name = $id.Split('|', 2)
      $k = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($path, $true)
      if ($k) { try { $k.DeleteValue($name, $false) } finally { $k.Close() } }
    } elseif ((Get-RegValueText $after[$id][1]) -ne (Get-RegValueText $before[$id][1]) -or $after[$id][0] -ne $before[$id][0]) {
      & $setVal $id $before[$id][0] $before[$id][1]
    }
  }
  foreach ($id in @($before.Keys)) {
    if (-not $after.ContainsKey($id)) { & $setVal $id $before[$id][0] $before[$id][1] }
  }
}

function Select-Combo([IntPtr]$dlg, [int]$id, [string]$text) {
  $cb = [JwW]::GetDlgItem($dlg, $id)
  $n = [int][JwW]::SendMessage($cb, $CB_GETCOUNT, [IntPtr]0, [IntPtr]0)
  for ($i = 0; $i -lt $n; $i++) {
    if ([JwW]::ComboText($cb, $i) -eq $text) {
      [void][JwW]::SendMessage($cb, $CB_SETCURSEL, [IntPtr]$i, [IntPtr]0)
      $wp = [IntPtr](($id -band 0xFFFF) -bor (1 -shl 16))   # CBN_SELCHANGE
      [void][JwW]::SendMessage([JwW]::GetParent($cb), $WM_COMMAND, $wp, $cb)
      return $true
    }
  }
  return $false
}

function Stop-Jw($proc) {
  if (-not $proc -or $proc.HasExited) { return }
  foreach ($h in [JwW]::Tops($proc.Id)) { [void][JwW]::PostMessage($h, $WM_CLOSE, [IntPtr]0, [IntPtr]0) }
  $ok = Wait-Until { $proc.HasExited } 8
  if (-not $ok) {
    # 保存確認などが出ていたら「いいえ」（図面は絶対に保存しない）
    foreach ($h in [JwW]::Tops($proc.Id)) {
      if ([JwW]::Cls($h) -eq '#32770') {
        $no = [JwW]::GetDlgItem($h, 7)
        if ($no -ne [IntPtr]::Zero) { [void][JwW]::PostMessage($no, $BM_CLICK, [IntPtr]0, [IntPtr]0) }
      }
    }
    $ok = Wait-Until { $proc.HasExited } 5
  }
  if (-not $proc.HasExited) { try { $proc.Kill() } catch {} }
}

# ---------- 1図面を PDF にする ----------
# 戻り値: 成功なら $null、失敗なら理由の文字列
function Convert-One([string]$jww, [string]$outPdf, $opt, [scriptblock]$status) {
  $jwExe = $opt.JwExe
  $name = [IO.Path]::GetFileName($jww)
  $paper = Get-JwwPaper $jww
  if (-not $paper) { return 'Jw_cad の図面ファイル(.jww)ではありません' }
  if ($paper.StartsWith('?') -or @('A0','A1','A2','A3','A4') -notcontains $paper) {
    return "用紙サイズ $paper の図面には対応していません（A0〜A4のみ）"
  }
  $printPaper = $paper; $scaleIdx = 2
  if ($opt.Paper -eq 'A3' -and $ToA3Scale.ContainsKey($paper)) { $printPaper = 'A3'; $scaleIdx = $ToA3Scale[$paper] }

  $tmpDir = Join-Path $env:TEMP 'jw_pdf_work'
  New-Item -ItemType Directory -Force $tmpDir | Out-Null
  Get-ChildItem $tmpDir -Filter 'out_*.pdf' -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt (Get-Date).AddHours(-1) } | Remove-Item -Force -ErrorAction SilentlyContinue
  $tmpPdf = Join-Path $tmpDir ('out_' + [Guid]::NewGuid().ToString('N').Substring(0, 8) + '.pdf')

  & $status "開いています（$paper → $printPaper）"
  $regSnap = $null
  try { $regSnap = Get-RegSnapshot } catch { Write-Log "  設定の控えに失敗: $($_.Exception.Message)" }
  $proc = Start-Process -FilePath $jwExe -ArgumentList ('"' + $jww + '"') -WorkingDirectory (Split-Path $jwExe) -PassThru -WindowStyle Minimized
  try {
    # 1) Jw_cad の窓が開くのを待つ
    $main = Wait-Until {
      if ($proc.HasExited) { return 'exit' }
      foreach ($h in [JwW]::Tops($proc.Id)) { if ([JwW]::Cls($h) -like 'Afx:*' -and [JwW]::Text($h) -like "*jw_win*") { return $h } }
      $null
    } 40
    if (-not $main -or $main -is [string]) { return 'Jw_cad が起動しませんでした' }
    Start-Sleep -Milliseconds 800
    $od = Find-OtherDialog $proc.Id
    if ($od) { return "Jw_cad のメッセージ: $($od.Text)" }

    # 2) 印刷コマンド → プリンターの設定
    [void][JwW]::PostMessage($main, $WM_COMMAND, [IntPtr]$ID_FILE_PRINT, [IntPtr]0)
    $dlg = Wait-Until { Find-Dialog $proc.Id 'プリンターの設定' } 20
    if (-not $dlg) {
      $od = Find-OtherDialog $proc.Id
      if ($od) { return "Jw_cad のメッセージ: $($od.Text)" }
      return '印刷の設定画面が出ませんでした'
    }
    Start-Sleep -Milliseconds 300
    if (-not (Select-Combo $dlg 1136 $PrinterName)) {
      [void][JwW]::PostMessage($dlg, $WM_COMMAND, [IntPtr]2, [IntPtr]0)
      return "プリンター「$PrinterName」が見つかりません（CubePDF を入れてください）"
    }
    Start-Sleep -Milliseconds 500
    if (-not (Select-Combo $dlg 1137 $printPaper)) {
      [void][JwW]::PostMessage($dlg, $WM_COMMAND, [IntPtr]2, [IntPtr]0)
      return "用紙 $printPaper を選べませんでした"
    }
    $orientId = if ($opt.Orient -eq 'tate') { 1056 } else { 1057 }
    [void][JwW]::SendMessage([JwW]::GetDlgItem($dlg, $orientId), $BM_CLICK, [IntPtr]0, [IntPtr]0)
    Start-Sleep -Milliseconds 200
    [void][JwW]::PostMessage($dlg, $WM_COMMAND, [IntPtr]1, [JwW]::GetDlgItem($dlg, 1))

    # 3) 印刷モードの操作バー（「印刷 (L)」ボタン）を待つ
    $bar = Wait-Until {
      foreach ($k in [JwW]::Kids($main)) {
        if ([JwW]::Cls($k) -eq 'Button' -and [JwW]::GetDlgCtrlID($k) -eq 1065 -and [JwW]::IsWindowVisible($k)) { return [JwW]::GetParent($k) }
      }
      $null
    } 20
    if (-not $bar) { return '印刷の操作バーが出ませんでした' }
    Start-Sleep -Milliseconds 400

    # 印刷倍率
    $cb = [JwW]::GetDlgItem($bar, 1411)
    if ($cb -ne [IntPtr]::Zero) {
      $cur = [int][JwW]::SendMessage($cb, $CB_GETCURSEL, [IntPtr]0, [IntPtr]0)
      if ($cur -ne $scaleIdx) {
        [void][JwW]::SendMessage($cb, $CB_SETCURSEL, [IntPtr]$scaleIdx, [IntPtr]0)
        $wp = [IntPtr]((1411 -band 0xFFFF) -bor (1 -shl 16))
        [void][JwW]::SendMessage($bar, $WM_COMMAND, $wp, $cb)
        Start-Sleep -Milliseconds 400
      }
    }
    # カラー印刷（keep のときは Jw_cad の今の設定のまま）
    if ($opt.Color -ne 'keep') {
      $cbtn = [JwW]::GetDlgItem($bar, 1343)
      if ($cbtn -ne [IntPtr]::Zero) {
        $isColor = ([int][JwW]::SendMessage($cbtn, $BM_GETCHECK, [IntPtr]0, [IntPtr]0)) -eq 1
        if ($isColor -ne ($opt.Color -eq 'color')) {
          [void][JwW]::SendMessage($cbtn, $BM_CLICK, [IntPtr]0, [IntPtr]0)
          Start-Sleep -Milliseconds 300
        }
      }
    }

    # 4) 印刷
    & $status '印刷しています'
    $t1 = Get-Date
    $before = @(Get-Process CubePdf -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    [void][JwW]::PostMessage([JwW]::GetDlgItem($bar, 1065), $BM_CLICK, [IntPtr]0, [IntPtr]0)

    # 5) CubePDF の画面を待ち、保存先を入れて「変換」
    $cube = Wait-Until {
      foreach ($p in (Get-Process CubePdf -ErrorAction SilentlyContinue)) {
        if ($before -contains $p.Id) { continue }
        $p.Refresh()
        if ($p.MainWindowHandle -ne [IntPtr]::Zero -and $p.MainWindowTitle.StartsWith("$name - CubePDF")) { return $p }
      }
      if (((Get-Date) - $t1).TotalSeconds -gt 15) { $od = Find-OtherDialog $proc.Id; if ($od) { return @{ Err = $od.Text } } }
      $null
    } 60
    if (-not $cube) { return 'CubePDF の画面が出ませんでした' }
    if ($cube -is [hashtable]) { return "Jw_cad のメッセージ: $($cube.Err)" }

    $ctrls = Wait-Until {
      $edit = $null; $btn = $null
      foreach ($k in [JwW]::Kids($cube.MainWindowHandle)) {
        $c = [JwW]::Cls($k)
        if (-not $edit -and $c -like 'WindowsForms10.EDIT*' -and [JwW]::Text($k) -match '\.pdf$') { $edit = $k }
        if (-not $btn -and $c -like 'WindowsForms10.BUTTON*' -and [JwW]::Text($k) -eq '変換') { $btn = $k }
      }
      if ($edit -and $btn) { return @($edit, $btn) }
      $null
    } 15
    if (-not $ctrls) {
      Write-Log ('  CubePDF 部品: ' + ((([JwW]::Kids($cube.MainWindowHandle)) | ForEach-Object { [JwW]::Cls($_) + '=' + [JwW]::Text($_) }) -join ' | '))
      try { $cube.Kill() } catch {}; return 'CubePDF の保存先欄が見つかりませんでした' }
    & $status 'PDF を保存しています'
    # 変換後に PDF を開く設定になっていると、ビューアがファイルを掴むので今回だけ「なし」に
    foreach ($k in [JwW]::Kids($cube.MainWindowHandle)) {
      if ([JwW]::Cls($k) -notlike 'WindowsForms10.COMBOBOX*') { continue }
      $n = [int][JwW]::SendMessage($k, $CB_GETCOUNT, [IntPtr]0, [IntPtr]0)
      $items = @(); for ($i = 0; $i -lt $n; $i++) { $items += [JwW]::ComboText($k, $i) }
      if ($env:JW_PDF_DEBUG) { Write-Log ("  combo sel=" + [int][JwW]::SendMessage($k, $CB_GETCURSEL, [IntPtr]0, [IntPtr]0) + " : " + ($items -join ',')) }
      $iNone = [array]::IndexOf($items, '何もしない'); if ($iNone -lt 0) { $iNone = [array]::IndexOf($items, 'なし') }
      if ($iNone -ge 0 -and $items -contains '開く') {
        if ([int][JwW]::SendMessage($k, $CB_GETCURSEL, [IntPtr]0, [IntPtr]0) -ne $iNone) {
          [void][JwW]::SendMessage($k, $CB_SETCURSEL, [IntPtr]$iNone, [IntPtr]0)
          $wp = [IntPtr](([JwW]::GetDlgCtrlID($k) -band 0xFFFF) -bor (1 -shl 16))
          [void][JwW]::SendMessage([JwW]::GetParent($k), $WM_COMMAND, $wp, $k)
        }
      }
    }
    [void][JwW]::SendMessage($ctrls[0], $WM_SETTEXT, [IntPtr]0, $tmpPdf)
    Start-Sleep -Milliseconds 200
    [void][JwW]::PostMessage($ctrls[1], $BM_CLICK, [IntPtr]0, [IntPtr]0)

    $done = Wait-Until { $cube.HasExited } 180 300
    if (-not $done) { try { $cube.Kill() } catch {}; return 'CubePDF の変換が終わりませんでした' }
    if (-not (Test-Path $tmpPdf) -or (Get-Item $tmpPdf).Length -eq 0) { return 'PDF ができませんでした' }

    # 6) 置き場所へ移す（同名 PDF は上書き）
    $outDir = Split-Path -Parent $outPdf
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Force $outDir | Out-Null }
    # CubePDF が閉じた直後はまだ書き込み中のことがあるので、移せるまで少し待つ
    $script:moveErr = ''
    $moved = Wait-Until {
      try { Move-Item -LiteralPath $tmpPdf -Destination $outPdf -Force -ErrorAction Stop; $true }
      catch { $script:moveErr = $_.Exception.Message; $false }
    } 30 500
    if (-not $moved) { return "PDF を保存できませんでした（同じ名前の PDF を開いたままではありませんか？）: $outPdf … $($script:moveErr)" }
    return $null
  }
  finally {
    Stop-Jw $proc
    if ($proc.HasExited) {
      try { Restore-RegSnapshot $regSnap } catch { Write-Log "  設定の戻しに失敗: $($_.Exception.Message)" }
    }
    if (Test-Path $tmpPdf) { Remove-Item $tmpPdf -Force -ErrorAction SilentlyContinue }
  }
}

function Get-OutPath([string]$jww, $opt) {
  $base = [IO.Path]::GetFileNameWithoutExtension($jww) + '.pdf'
  if ($opt.OutMode -eq 'folder' -and $opt.OutDir) { return Join-Path $opt.OutDir $base }
  return Join-Path (Split-Path -Parent $jww) $base
}

function Expand-Inputs([string[]]$items, [bool]$recurse) {
  $list = New-Object System.Collections.Generic.List[string]
  foreach ($it in $items) {
    if (-not $it) { continue }
    $it = $it.Trim('"')
    if (Test-Path -LiteralPath $it -PathType Container) {
      $gci = @{ LiteralPath = $it; Filter = '*.jww'; File = $true }
      if ($recurse) { $gci.Recurse = $true }
      Get-ChildItem @gci | Sort-Object FullName | ForEach-Object { $list.Add($_.FullName) }
    } elseif ((Test-Path -LiteralPath $it -PathType Leaf) -and $it -match '\.jww$') {
      $list.Add((Resolve-Path -LiteralPath $it).ProviderPath)
    }
  }
  return $list
}

# ---------- 画面なしで実行（試験用）: jw_pdf.ps1 -NoGui a.jww b.jww ----------
if ($NoGui) {
  $o = [ordered]@{ OutMode = 'same'; OutDir = ''; Paper = 'same'; Orient = 'yoko'; Color = 'keep'; JwExe = (Find-JwExe) }
  if ($Paper) { $o.Paper = $Paper }
  if ($Color) { $o.Color = $Color }
  if ($OutDir) { $o.OutMode = 'folder'; $o.OutDir = $OutDir }
  $ng = 0
  foreach ($f in (Expand-Inputs $Paths $false)) {
    $out = Get-OutPath $f $o
    $err = Convert-One $f $out $o { param($m) }
    if ($err) { $ng++; Write-Host "NG  $f … $err"; Write-Log "NG  $f  … $err" } else { Write-Host "OK  $f → $out"; Write-Log "OK  $f  → $out" }
  }
  exit $ng
}
# ---------- 画面 ----------
[System.Windows.Forms.Application]::EnableVisualStyles()

$opt = [ordered]@{ OutMode = 'same'; OutDir = ''; Paper = 'same'; Orient = 'yoko'; Color = 'keep'; Recurse = $false; JwExe = '' }
if (Test-Path $SettingFile) {
  try { $j = Get-Content $SettingFile -Raw -Encoding UTF8 | ConvertFrom-Json; foreach ($k in @($opt.Keys)) { if ($null -ne $j.$k) { $opt[$k] = $j.$k } } } catch {}
}
if (-not $opt.JwExe -or -not (Test-Path $opt.JwExe)) { $opt.JwExe = Find-JwExe }

$font = New-Object System.Drawing.Font('Yu Gothic UI', 10)
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Jw_cad 図面 一括PDF化'
$form.Size = New-Object System.Drawing.Size(820, 640)
$form.MinimumSize = New-Object System.Drawing.Size(700, 520)
$form.StartPosition = 'CenterScreen'
$form.Font = $font
$form.AllowDrop = $true

$lbl = New-Object System.Windows.Forms.Label
$lbl.Text = 'PDF にする図面（.jww ファイルやフォルダをここへドラッグしても追加できます）'
$lbl.Location = '12,10'; $lbl.AutoSize = $true
$form.Controls.Add($lbl)

$lv = New-Object System.Windows.Forms.ListView
$lv.View = 'Details'; $lv.FullRowSelect = $true; $lv.GridLines = $true; $lv.HideSelection = $false
$lv.Location = '12,34'; $lv.Size = '780,300'; $lv.Anchor = 'Top,Left,Right,Bottom'
[void]$lv.Columns.Add('図面', 280); [void]$lv.Columns.Add('用紙', 60); [void]$lv.Columns.Add('状態', 120); [void]$lv.Columns.Add('場所', 300)
$lv.AllowDrop = $true
$form.Controls.Add($lv)

function Add-Files([string[]]$items) {
  $have = @{}; foreach ($it in $lv.Items) { $have[$it.Tag] = $true }
  foreach ($f in (Expand-Inputs $items $chkSub.Checked)) {
    if ($have[$f]) { continue }
    $paper = try { Get-JwwPaper $f } catch { '?' }
    $it = New-Object System.Windows.Forms.ListViewItem([IO.Path]::GetFileName($f))
    [void]$it.SubItems.Add([string]$paper); [void]$it.SubItems.Add('待ち'); [void]$it.SubItems.Add((Split-Path -Parent $f))
    $it.Tag = $f
    [void]$lv.Items.Add($it); $have[$f] = $true
  }
  Update-Count
}

$y = 344
$btnAdd = New-Object System.Windows.Forms.Button; $btnAdd.Text = 'ファイルを追加…'; $btnAdd.Location = "12,$y"; $btnAdd.Size = '140,32'; $btnAdd.Anchor = 'Bottom,Left'
$btnDir = New-Object System.Windows.Forms.Button; $btnDir.Text = 'フォルダを追加…'; $btnDir.Location = "158,$y"; $btnDir.Size = '140,32'; $btnDir.Anchor = 'Bottom,Left'
$chkSub = New-Object System.Windows.Forms.CheckBox; $chkSub.Text = 'サブフォルダも'; $chkSub.Location = "306,$($y+6)"; $chkSub.AutoSize = $true; $chkSub.Anchor = 'Bottom,Left'; $chkSub.Checked = [bool]$opt.Recurse
$btnDel = New-Object System.Windows.Forms.Button; $btnDel.Text = '選んだ行を外す'; $btnDel.Location = "530,$y"; $btnDel.Size = '130,32'; $btnDel.Anchor = 'Bottom,Right'
$btnClr = New-Object System.Windows.Forms.Button; $btnClr.Text = '全部外す'; $btnClr.Location = "666,$y"; $btnClr.Size = '126,32'; $btnClr.Anchor = 'Bottom,Right'
$form.Controls.AddRange(@($btnAdd, $btnDir, $chkSub, $btnDel, $btnClr))

# 設定
$grp = New-Object System.Windows.Forms.GroupBox; $grp.Text = '設定'; $grp.Location = "12,$($y+40)"; $grp.Size = '780,140'; $grp.Anchor = 'Bottom,Left,Right'
$form.Controls.Add($grp)
function New-Radio($text, $x, $yy, $parent) { $r = New-Object System.Windows.Forms.RadioButton; $r.Text = $text; $r.Location = "$x,$yy"; $r.AutoSize = $true; $parent.Controls.Add($r); return $r }
function New-Panel($x, $yy, $w) { $p = New-Object System.Windows.Forms.Panel; $p.Location = "$x,$yy"; $p.Size = "$w,26"; $grp.Controls.Add($p); return $p }
function New-Label($text, $x, $yy) { $l = New-Object System.Windows.Forms.Label; $l.Text = $text; $l.Location = "$x,$yy"; $l.AutoSize = $true; $grp.Controls.Add($l) }

New-Label '保存先' 12 26
$pOut = New-Panel 90 22 680
$rSame = New-Radio '図面と同じフォルダ' 0 2 $pOut
$rDir = New-Radio '指定フォルダ:' 170 2 $pOut
$txtDir = New-Object System.Windows.Forms.TextBox; $txtDir.Location = '290,0'; $txtDir.Size = '300,24'; $txtDir.Text = $opt.OutDir; $pOut.Controls.Add($txtDir)
$btnBrowse = New-Object System.Windows.Forms.Button; $btnBrowse.Text = '参照…'; $btnBrowse.Location = '596,-1'; $btnBrowse.Size = '70,27'; $pOut.Controls.Add($btnBrowse)
if ($opt.OutMode -eq 'folder') { $rDir.Checked = $true } else { $rSame.Checked = $true }

New-Label '用紙' 12 56
$pPaper = New-Panel 90 52 680
$rPSame = New-Radio '図面と同じ大きさ（A2 の図面は A2 の PDF）' 0 2 $pPaper
$rPA3 = New-Radio 'A3 にそろえる（大きい図面は縮小）' 330 2 $pPaper
if ($opt.Paper -eq 'A3') { $rPA3.Checked = $true } else { $rPSame.Checked = $true }

New-Label '向き・色' 12 86
$pOr = New-Panel 90 82 200
$rYoko = New-Radio '横' 0 2 $pOr; $rTate = New-Radio '縦' 60 2 $pOr
if ($opt.Orient -eq 'tate') { $rTate.Checked = $true } else { $rYoko.Checked = $true }
$pCol = New-Panel 290 82 480
$rCKeep = New-Radio 'Jw_cad の設定のまま' 0 2 $pCol; $rCMono = New-Radio '白黒' 180 2 $pCol; $rCColor = New-Radio 'カラー' 250 2 $pCol
switch ($opt.Color) { 'mono' { $rCMono.Checked = $true } 'color' { $rCColor.Checked = $true } default { $rCKeep.Checked = $true } }

New-Label 'Jw_cad' 12 114
$lblJw = New-Object System.Windows.Forms.Label; $lblJw.Location = '92,114'; $lblJw.AutoSize = $true; $grp.Controls.Add($lblJw)
$lblJw.Text = if ($opt.JwExe) { $opt.JwExe } else { '見つかりません（Jw_win.exe の場所を設定してください）' }
$btnJw = New-Object System.Windows.Forms.Button; $btnJw.Text = '変更…'; $btnJw.Location = '686,108'; $btnJw.Size = '80,27'; $grp.Controls.Add($btnJw)

$lblStat = New-Object System.Windows.Forms.Label; $lblStat.Location = "12,$($y+192)"; $lblStat.Size = '560,40'; $lblStat.Anchor = 'Bottom,Left,Right'
$btnRun = New-Object System.Windows.Forms.Button; $btnRun.Text = 'PDF にする'; $btnRun.Location = "582,$($y+188)"; $btnRun.Size = '210,40'; $btnRun.Anchor = 'Bottom,Right'
$btnRun.Font = New-Object System.Drawing.Font('Yu Gothic UI', 11, [System.Drawing.FontStyle]::Bold)
$form.Controls.AddRange(@($lblStat, $btnRun))

function Update-Count { $lblStat.Text = "$($lv.Items.Count) 図面" }

function Read-Options {
  $opt.OutMode = if ($rDir.Checked) { 'folder' } else { 'same' }
  $opt.OutDir = $txtDir.Text.Trim()
  $opt.Paper = if ($rPA3.Checked) { 'A3' } else { 'same' }
  $opt.Orient = if ($rTate.Checked) { 'tate' } else { 'yoko' }
  $opt.Color = if ($rCMono.Checked) { 'mono' } elseif ($rCColor.Checked) { 'color' } else { 'keep' }
  $opt.Recurse = $chkSub.Checked
  try { ($opt | ConvertTo-Json) | Set-Content $SettingFile -Encoding UTF8 } catch {}
}

$btnAdd.Add_Click({
  $d = New-Object System.Windows.Forms.OpenFileDialog
  $d.Filter = 'Jw_cad 図面 (*.jww)|*.jww'; $d.Multiselect = $true; $d.Title = 'PDF にする図面を選ぶ（複数選べます）'
  if ($d.ShowDialog($form) -eq 'OK') { Add-Files $d.FileNames }
})
$btnDir.Add_Click({
  $d = New-Object System.Windows.Forms.FolderBrowserDialog; $d.Description = 'このフォルダの .jww をすべて追加します'
  if ($d.ShowDialog($form) -eq 'OK') { Add-Files @($d.SelectedPath) }
})
$btnDel.Add_Click({ foreach ($it in @($lv.SelectedItems)) { $lv.Items.Remove($it) }; Update-Count })
$btnClr.Add_Click({ $lv.Items.Clear(); Update-Count })
$btnBrowse.Add_Click({
  $d = New-Object System.Windows.Forms.FolderBrowserDialog; $d.Description = 'PDF の保存先'
  if ($d.ShowDialog($form) -eq 'OK') { $txtDir.Text = $d.SelectedPath; $rDir.Checked = $true }
})
$btnJw.Add_Click({
  $d = New-Object System.Windows.Forms.OpenFileDialog; $d.Filter = 'Jw_win.exe|Jw_win.exe|exe|*.exe'
  if ($d.ShowDialog($form) -eq 'OK') { $opt.JwExe = $d.FileName; $lblJw.Text = $d.FileName }
})
$dragEnter = { param($s, $e) if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { $e.Effect = 'Copy' } }
$dragDrop = { param($s, $e) Add-Files ([string[]]$e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)) }
$form.Add_DragEnter($dragEnter); $form.Add_DragDrop($dragDrop); $lv.Add_DragEnter($dragEnter); $lv.Add_DragDrop($dragDrop)

$script:running = $false
$btnRun.Add_Click({
  if ($script:running) { return }
  Read-Options
  if ($lv.Items.Count -eq 0) { [void][System.Windows.Forms.MessageBox]::Show($form, '図面を追加してください。', 'Jw_cad 図面 一括PDF化'); return }
  if (-not $opt.JwExe -or -not (Test-Path $opt.JwExe)) { [void][System.Windows.Forms.MessageBox]::Show($form, 'Jw_win.exe が見つかりません。「変更…」で場所を選んでください。', 'Jw_cad 図面 一括PDF化'); return }
  if ($opt.OutMode -eq 'folder' -and -not $opt.OutDir) { [void][System.Windows.Forms.MessageBox]::Show($form, '保存先フォルダを選んでください。', 'Jw_cad 図面 一括PDF化'); return }
  if (-not (Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue)) { [void][System.Windows.Forms.MessageBox]::Show($form, "プリンター「$PrinterName」がありません。CubePDF を入れてください。", 'Jw_cad 図面 一括PDF化'); return }

  $script:running = $true
  foreach ($c in @($btnRun, $btnAdd, $btnDir, $btnDel, $btnClr, $grp)) { $c.Enabled = $false }
  $ok = 0; $ng = 0; $total = $lv.Items.Count; $i = 0
  Write-Log "===== 開始 $total 図面 / 用紙=$($opt.Paper) 向き=$($opt.Orient) 色=$($opt.Color)"
  foreach ($it in $lv.Items) {
    $i++
    if ($it.SubItems[2].Text -eq '完了') { $ok++; continue }
    $jww = [string]$it.Tag; $out = Get-OutPath $jww $opt
    $it.EnsureVisible()
    $st = { param($m) $it.SubItems[2].Text = $m; $lblStat.Text = "$i / $total  $($it.Text)  …$m"; Pump }
    $err = $null
    try { $err = Convert-One $jww $out $opt $st } catch { $err = "エラー: $($_.Exception.Message)" }
    if ($err) {
      $ng++; $it.SubItems[2].Text = '失敗'; $it.ForeColor = [System.Drawing.Color]::Firebrick; $it.ToolTipText = $err
      Write-Log "NG  $jww  … $err"
    } else {
      $ok++; $it.SubItems[2].Text = '完了'; $it.ForeColor = [System.Drawing.Color]::DarkGreen
      Write-Log "OK  $jww  → $out"
    }
    Pump
  }
  Write-Log "===== 終了 成功 $ok / 失敗 $ng"
  foreach ($c in @($btnRun, $btnAdd, $btnDir, $btnDel, $btnClr, $grp)) { $c.Enabled = $true }
  $script:running = $false
  $lblStat.Text = "終わりました。成功 $ok / 失敗 $ng" + $(if ($ng) { '（失敗の理由は行にマウスを乗せるか jw_pdf_log.txt で）' } else { '' })
  $msg = "成功 $ok 図面" + $(if ($ng) { " / 失敗 $ng 図面" } else { '' })
  [void][System.Windows.Forms.MessageBox]::Show($form, $msg, 'Jw_cad 図面 一括PDF化')
})
$lv.ShowItemToolTips = $true

$form.Add_FormClosing({ param($s, $e) if ($script:running) { $e.Cancel = $true } else { Read-Options } })
Update-Count
if ($Paths) { $form.Add_Shown({ Add-Files $Paths }) }
[void]$form.ShowDialog()
