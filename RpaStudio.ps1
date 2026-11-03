#requires -Version 5.1
$script:RuntimeSourcePath=$PSCommandPath
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Security
. (Join-Path $PSScriptRoot 'ElementAutomation.ps1')
. (Join-Path $PSScriptRoot 'WebAutomation.ps1')
. (Join-Path $PSScriptRoot 'ExcelAutomation.ps1')
. (Join-Path $PSScriptRoot 'DateAutomation.ps1')
. (Join-Path $PSScriptRoot 'WindowAutomation.ps1')
. (Join-Path $PSScriptRoot 'ImageAutomation.ps1')
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class RpaNative {
 [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
 [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint x, uint y, uint d, UIntPtr e);
 [DllImport("user32.dll")] public static extern short GetAsyncKeyState(int key);
 [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
 [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
 public static uint ForegroundProcess() { uint process; GetWindowThreadProcessId(GetForegroundWindow(),out process); return process; }
 [StructLayout(LayoutKind.Sequential)] struct Keyboard { public ushort vk, scan; public uint flags, time; public UIntPtr extra; }
 [StructLayout(LayoutKind.Sequential)] struct Mouse { public int x,y; public uint data,flags,time; public UIntPtr extra; }
 [StructLayout(LayoutKind.Explicit)] struct InputData { [FieldOffset(0)] public Keyboard keyboard; [FieldOffset(0)] public Mouse mouse; }
 [StructLayout(LayoutKind.Sequential)] struct Input { public uint type; public InputData data; }
 [DllImport("user32.dll", SetLastError=true)] static extern uint SendInput(uint count, Input[] inputs, int size);
 public static void TypeCharacter(char character) {
  Input[] keys = new Input[2];
  keys[0].type=keys[1].type=1;
  keys[0].data.keyboard.scan=keys[1].data.keyboard.scan=character;
  keys[0].data.keyboard.flags=4; keys[1].data.keyboard.flags=6;
  if(SendInput(2,keys,Marshal.SizeOf(typeof(Input)))!=2) throw new InvalidOperationException("文字を送信できませんでした。対象アプリの権限を確認してください。");
 }
}
'@
[System.Windows.Forms.Application]::EnableVisualStyles()
$script:running = $false
$script:cancel = $false
$script:path = $null
$script:AutoRunScenario = $null
$form = New-Object Windows.Forms.Form
$form.Text = 'PowerShell RPA Studio'
$form.Size = New-Object Drawing.Size(1060,740)
$form.MinimumSize = New-Object Drawing.Size(900,650)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object Drawing.Font('Yu Gothic UI',10)

$toolbar = New-Object Windows.Forms.FlowLayoutPanel
$toolbar.Dock = 'Top'; $toolbar.Height = 90
$toolbar.AutoSize = $true; $toolbar.AutoSizeMode = 'GrowAndShrink'
$form.Controls.Add($toolbar)
function New-Button($label, $handler) {
 $b = New-Object Windows.Forms.Button
 $b.Text = $label; $b.AutoSize = $true; $b.Height = 34
 $b.Add_Click($handler); $toolbar.Controls.Add($b)
 return $b
}
$bottom = New-Object Windows.Forms.Panel
$bottom.Dock = 'Bottom'; $bottom.Height = 215
$form.Controls.Add($bottom)
$hint = New-Object Windows.Forms.Label
$hint.Text = '操作を追加・編集 → 保存 → 実行。実行開始まで3秒。Esc長押し、または停止ボタンで停止します。'
$hint.Dock = 'Top'; $hint.Height = 30; $bottom.Controls.Add($hint)
$log = New-Object Windows.Forms.TextBox
$log.Multiline = $true; $log.ReadOnly = $true; $log.ScrollBars = 'Vertical'
$log.Dock = 'Bottom'; $log.Height = 180; $bottom.Controls.Add($log)
$grid = New-Object Windows.Forms.DataGridView
$grid.Dock = 'Fill'; $grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false; $grid.RowHeadersVisible = $true
$grid.SelectionMode = 'FullRowSelect'; $grid.MultiSelect = $true
$grid.AutoSizeColumnsMode = 'Fill'
$form.Controls.Add($grid); $grid.BringToFront()
$status = New-Object Windows.Forms.StatusStrip
$positionLabel = New-Object Windows.Forms.ToolStripStatusLabel
$executionLabel = New-Object Windows.Forms.ToolStripStatusLabel
$executionLabel.Text = '待機中'
[void]$status.Items.Add($executionLabel)
[void]$status.Items.Add($positionLabel)
$form.Controls.Add($status)
$positionTimer = New-Object Windows.Forms.Timer
$positionTimer.Interval = 100
$positionTimer.Add_Tick({
 $p = [Windows.Forms.Cursor]::Position
 $positionLabel.Text = "現在のマウス位置  X: $($p.X)   Y: $($p.Y)    |    Ctrl / Shift + 行見出しクリックで複数選択"
})
$positionTimer.Start()
$excelActions=@('Excelブックを開く','Excelシート選択','Excelセル入力','Excelブック保存','Excelブックを閉じる')
$actions = @('待機','マウス移動','左クリック','ダブルクリック','右クリック','文字入力','キー送信','URLを開く','アプリ起動','要素クリック','要素入力','要素フォーカス','ID入力','パスワード入力','Web用URLを開く','ウェブ要素クリック','ウェブ要素入力','日付入力','ウィンドウをアクティブ化','RPA終了','画像クリック','画像ダブルクリック','画像待機') + $excelActions
$actionColumn = New-Object Windows.Forms.DataGridViewComboBoxColumn
$actionColumn.Name = 'Action'; $actionColumn.HeaderText = '操作'
$actionColumn.Items.AddRange([object[]]$actions); $actionColumn.FillWeight = 32
[void]$grid.Columns.Add($actionColumn)
foreach ($col in @(@('Value','値（文字・キー・URL・アプリ）',85),@('X','X',18),@('Y','Y',18),@('Delay','操作後の待機(ms)',30),@('Target','対象要素（JSON）',65))) {
 $c = New-Object Windows.Forms.DataGridViewTextBoxColumn
 $c.Name = $col[0]; $c.HeaderText = $col[1]; $c.FillWeight = $col[2]
 [void]$grid.Columns.Add($c)
}
$stateColumn = New-Object Windows.Forms.DataGridViewTextBoxColumn
$stateColumn.Name='State'; $stateColumn.HeaderText='実行状態'; $stateColumn.ReadOnly=$true
$stateColumn.FillWeight=40
[void]$grid.Columns.Add($stateColumn)
$secretColumn = New-Object Windows.Forms.DataGridViewTextBoxColumn
$secretColumn.Name='ProtectedValue'; $secretColumn.Visible=$false
[void]$grid.Columns.Add($secretColumn)
$dateColumn=New-Object Windows.Forms.DataGridViewTextBoxColumn
$dateColumn.Name='DateOptions';$dateColumn.Visible=$false
[void]$grid.Columns.Add($dateColumn)
$repeatColumn=New-Object Windows.Forms.DataGridViewTextBoxColumn
$repeatColumn.Name='Repeat';$repeatColumn.HeaderText='キー回数';$repeatColumn.FillWeight=22
$repeatColumn.DefaultCellStyle.NullValue='1'
[void]$grid.Columns.Add($repeatColumn)
$repeatColumn.DisplayIndex=2
$grid.Add_CellFormatting({param($sender,$e)
 if($e.RowIndex -ge 0 -and $grid.Columns[$e.ColumnIndex].Name -eq 'Target' -and $grid.Rows[$e.RowIndex].Cells['Action'].Value -in @('画像クリック','画像ダブルクリック','画像待機')){
  try{
   $target=[string]$e.Value|ConvertFrom-Json -ErrorAction Stop
   if($target.Kind -eq 'Image'){$e.Value="画像 $($target.Width)×$($target.Height)px / 横 $([int]$target.OffsetX) / 縦 $([int]$target.OffsetY)";$e.FormattingApplied=$true}
  }catch{}
 }
})
$grid.Add_CellBeginEdit({ param($sender,$e)
 if ($grid.Columns[$e.ColumnIndex].Name -eq 'Value' -and $grid.Rows[$e.RowIndex].Cells['Action'].Value -in @('ID入力','パスワード入力')) {
  $e.Cancel=$true
  Write-Log 'ID・パスワードは「入力情報設定」から登録してください。'
 }
})
$grid.Add_CellValueChanged({ param($sender,$e)
 if($e.RowIndex -ge 0 -and $e.ColumnIndex -eq 0) {
  $row=$grid.Rows[$e.RowIndex]
  if($row.Cells['Action'].Value -in @('ID入力','パスワード入力')) {
   $row.Cells['Value'].Value=if($row.Cells['ProtectedValue'].Value){'[暗号化済み]'}else{'[未設定]'}
  } else { $row.Cells['ProtectedValue'].Value='' }
 }
})
function Protect-InputValue([string]$value) {
 $bytes=[Text.Encoding]::UTF8.GetBytes($value)
 try {
  $encrypted=[Security.Cryptography.ProtectedData]::Protect($bytes,$null,[Security.Cryptography.DataProtectionScope]::CurrentUser)
  'DPAPI:v1:' + [Convert]::ToBase64String($encrypted)
 } finally { [Array]::Clear($bytes,0,$bytes.Length) }
}
function Unprotect-InputValue([string]$value) {
 $bytes=$null
 try {
  if(-not $value.StartsWith('DPAPI:v1:')){throw 'Unsupported format'}
  $bytes=[Security.Cryptography.ProtectedData]::Unprotect([Convert]::FromBase64String($value.Substring(9)),$null,[Security.Cryptography.DataProtectionScope]::CurrentUser)
  [Text.Encoding]::UTF8.GetString($bytes)
 } catch { throw '入力情報を復号できません。同じPC・Windowsユーザーで実行するか、「入力情報設定」で登録し直してください。' }
 finally { if($null -ne $bytes){[Array]::Clear($bytes,0,$bytes.Length)} }
}
function Add-StoredStep($step) {
 $index=$grid.Rows.Add($step.Action,$step.Value,$step.X,$step.Y,$step.Delay,$step.Target)
 $grid.Rows[$index].Cells['ProtectedValue'].Value=[string]$step.ProtectedValue
 $grid.Rows[$index].Cells['DateOptions'].Value=[string]$step.DateOptions
 $grid.Rows[$index].Cells['Repeat'].Value=if([string]::IsNullOrWhiteSpace([string]$step.Repeat)){'1'}else{[string]$step.Repeat}
 if($step.Action -in @('ID入力','パスワード入力')) { $grid.Rows[$index].Cells['Value'].Value='[暗号化済み]' }
}
function Invoke-ProtectedInput($step) {
 $plain=$null
 try {
  $plain=Unprotect-InputValue $step.ProtectedValue
  if($step.Target -and ($step.Target|ConvertFrom-Json).Kind -eq 'Web'){
   Invoke-WebStep $step -InputMode -inputValue $plain
   return
  }
  $element=$null; $direct=$false
  if($step.Target) {
   $element=Resolve-TargetElement $step.Target
   $pattern=$null
   if($element.TryGetCurrentPattern([Windows.Automation.ValuePattern]::Pattern,[ref]$pattern)) {
    if($pattern.Current.IsReadOnly){throw 'Read only'}
    ([Windows.Automation.ValuePattern]$pattern).SetValue($plain)
    $direct=$true
   } else { $element.SetFocus(); Wait-Rpa 100 }
  }
  if(-not $direct) {
   $foreground=[RpaNative]::GetForegroundWindow()
   if($foreground -eq [IntPtr]::Zero -or [RpaNative]::ForegroundProcess() -eq $PID){throw 'No external input target'}
   if(-not $element){$element=[Windows.Automation.AutomationElement]::FocusedElement}
   foreach($character in $plain.ToCharArray()) {
    Wait-Rpa 0
    if([RpaNative]::GetForegroundWindow() -ne $foreground){throw 'Focus changed'}
    if($element -and -not [Windows.Automation.Automation]::Compare($element,[Windows.Automation.AutomationElement]::FocusedElement)){throw 'Target focus changed'}
    [RpaNative]::TypeCharacter($character)
   }
  }
 } catch {
  if($_.Exception -is [OperationCanceledException] -or $script:cancel){throw [OperationCanceledException]::new('実行を停止しました。')}
  throw 'ID・パスワード入力に失敗しました。登録情報の復号、対象欄の指定、入力対応、対象アプリの権限を確認してください。'
 } finally { $plain=$null }
}
function Set-StepState([int]$rowNumber, [string]$state) {
 $row=$grid.Rows[$rowNumber-1]
 $row.Cells['State'].Value="$state（行 $rowNumber）"
 $color=switch($state) {
  '実行中' { [Drawing.Color]::LightYellow }
  'エラー' { [Drawing.Color]::MistyRose }
  '中断' { [Drawing.Color]::Gainsboro }
  default { [Drawing.Color]::White }
 }
 $row.DefaultCellStyle.BackColor=$color
 $row.DefaultCellStyle.SelectionBackColor=$color
 $row.DefaultCellStyle.SelectionForeColor=[Drawing.Color]::Black
 if($state -in @('実行中','エラー','中断')) {
  $grid.FirstDisplayedScrollingRowIndex=$rowNumber-1
 }
 $grid.Refresh()
}
function Write-Log($message) { $log.AppendText(('[' + (Get-Date -Format HH:mm:ss) + '] ' + $message + "`r`n")) }
function Get-Steps([switch]$SelectedOnly) {
 [void]$grid.EndEdit()
 foreach ($r in $grid.Rows) {
  if ($SelectedOnly -and -not $r.Selected) { continue }
  $value=if($r.Cells['Action'].Value -in @('ID入力','パスワード入力')){'[暗号化済み]'}else{[string]$r.Cells['Value'].Value}
  $repeat=if([string]::IsNullOrWhiteSpace([string]$r.Cells['Repeat'].Value)){'1'}else{[string]$r.Cells['Repeat'].Value}
  [pscustomobject]@{ RowNumber=$r.Index+1; Action=[string]$r.Cells['Action'].Value; Value=$value; X=[string]$r.Cells['X'].Value; Y=[string]$r.Cells['Y'].Value; Delay=[string]$r.Cells['Delay'].Value; Target=[string]$r.Cells['Target'].Value; ProtectedValue=[string]$r.Cells['ProtectedValue'].Value; DateOptions=[string]$r.Cells['DateOptions'].Value; Repeat=$repeat }
 }
}
function Check-Steps($steps, [switch]$DecryptSecrets) {
 if ($steps.Count -eq 0) { throw '操作を追加してください。' }
 foreach ($s in $steps) {
  if ($s.Action -notin $actions) { throw "未対応の操作: $($s.Action)" }
  if($s.Action -in @('画像クリック','画像ダブルクリック','画像待機')){
   $imageTarget=Read-ImageTarget $s.Target
   $imageTarget.Bitmap.Dispose()
  }
  if($s.Action -eq 'キー送信'){$null=Get-KeyRepeatCount $s}
  if($s.Action -eq 'ウィンドウをアクティブ化'){
   if([string]::IsNullOrWhiteSpace($s.Value)){throw '「ウィンドウ指定」で対象を選ぶか、値にウィンドウ名を入力してください。'}
   if($s.Target -and ($s.Target|ConvertFrom-Json -ErrorAction Stop).Kind -ne 'Window'){throw '「ウィンドウ指定」で対象を指定し直してください。'}
  }
  $dateTarget=$null;$excelTarget=$null
  if($s.Action -eq '日付入力'){
   $null=Get-DateInputText $s
   if($s.Target){$dateTarget=$s.Target|ConvertFrom-Json -ErrorAction Stop}
  }
  if($s.Action -eq 'Excelブックを開く'){
   if(-not [IO.Path]::IsPathRooted([string]$s.Value) -or [IO.Path]::GetExtension([string]$s.Value) -notin @('.xlsx','.xlsm','.xlsb','.xls')){throw 'ExcelブックにはExcelファイルの絶対パスを指定してください。'}
  }
  if($s.Action -eq 'Excelシート選択' -and [string]::IsNullOrWhiteSpace($s.Value)){throw 'Excelのシート名を値に入力してください。'}
  if($s.Action -eq 'Excelセル入力' -or ($dateTarget -and $dateTarget.Kind -eq 'Excel')){
   $excelTarget=$s.Target|ConvertFrom-Json -ErrorAction Stop
   if($excelTarget.Kind -ne 'Excel' -or $excelTarget.Cell -notmatch '^\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}$'){throw '「Excel操作設定」で入力先セル（A1など）を設定してください。'}
  }
  if ($s.Action -in @('ID入力','パスワード入力')) {
   if(-not $s.ProtectedValue -or -not $s.ProtectedValue.StartsWith('DPAPI:v1:')) { throw "行 $($s.RowNumber): 「入力情報設定」でID・パスワードを登録してください。" }
   if($DecryptSecrets) {
    $secretCheck=Unprotect-InputValue $s.ProtectedValue
    $secretCheck=$null
   }
  }
  $webTarget=$null
  if($s.Action -in @('ウェブ要素クリック','ウェブ要素入力') -or ($s.Action -in @('ID入力','パスワード入力','日付入力') -and $s.Target -and ($s.Target|ConvertFrom-Json).Kind -eq 'Web')){
   $webTarget=$s.Target|ConvertFrom-Json -ErrorAction Stop
   $pageUri=$null
   if($webTarget.Kind -ne 'Web' -or -not $webTarget.Selector -or -not [uri]::TryCreate($webTarget.PageUrl,[UriKind]::Absolute,[ref]$pageUri) -or $pageUri.Scheme -notin @('http','https')){throw '「ウェブ要素取得」で対象を設定してください。'}
  }
  if ($s.Action -in @('要素クリック','要素入力','要素フォーカス') -or ($s.Action -in @('ID入力','パスワード入力','日付入力') -and $s.Target -and -not $webTarget -and -not $excelTarget)) {
   $target = $s.Target | ConvertFrom-Json -ErrorAction Stop
   if (-not $target.Process -or -not $target.Window -or -not $target.Type -or (-not $target.Id -and -not $target.Name)) { throw '対象要素が未設定です。「要素取得」で取得してください。' }
  }
  $n = 0
  if (-not [int]::TryParse($s.Delay,[ref]$n) -or $n -lt 0 -or $n -gt 3600000) { throw '操作後の待機は0～3600000msの整数で指定してください。' }
  if ($s.Action -in @('マウス移動','左クリック','ダブルクリック','右クリック')) {
   foreach ($v in @($s.X,$s.Y)) { if (-not [int]::TryParse($v,[ref]$n)) { throw '座標は整数で指定してください。' } }
  }
  if ($s.Action -eq '待機' -and (-not [int]::TryParse($s.Value,[ref]$n) -or $n -lt 0 -or $n -gt 3600000)) { throw '待機の値は0～3600000msで指定してください。' }
  if ($s.Action -in @('URLを開く','Web用URLを開く')) {
   $uri = $null
   if (-not [uri]::TryCreate($s.Value,[UriKind]::Absolute,[ref]$uri) -or $uri.Scheme -notin @('http','https')) { throw 'URLはhttp://またはhttps://で指定してください。' }
  }
  if ($s.Action -in @('アプリ起動','キー送信') -and [string]::IsNullOrWhiteSpace($s.Value)) { throw '値を入力してください。' }
 }
}
function Wait-Rpa([int]$ms) {
 $timer = [Diagnostics.Stopwatch]::StartNew()
 do {
  [Windows.Forms.Application]::DoEvents()
  if ($script:cancel -or ([RpaNative]::GetAsyncKeyState(27) -band 0x8000)) { throw [OperationCanceledException]::new('実行を停止しました。') }
  if ($timer.ElapsedMilliseconds -ge $ms) { break }
  Start-Sleep -Milliseconds 20
 } while ($true)
}
function Mouse-Click([uint32]$down,[uint32]$up) {
 [RpaNative]::mouse_event($down,0,0,0,[UIntPtr]::Zero)
 [RpaNative]::mouse_event($up,0,0,0,[UIntPtr]::Zero)
}
function Get-KeyRepeatCount($step) {
 if([string]::IsNullOrWhiteSpace([string]$step.Repeat)){return 1}
 $count=0
 if(-not [int]::TryParse([string]$step.Repeat,[ref]$count) -or $count -lt 1 -or $count -gt 10000){throw 'キー回数は1～10000の整数で指定してください。'}
 return $count
}
function Send-RpaKeys([string]$keys){[Windows.Forms.SendKeys]::SendWait([regex]::Replace($keys,'(?i)\{SPACE\}',' '))}
function Invoke-KeyStep($step) {
 $count=Get-KeyRepeatCount $step
 for($iteration=0;$iteration -lt $count;$iteration++){
  Wait-Rpa 0
  Send-RpaKeys $step.Value
  if($iteration -lt $count-1){Wait-Rpa 50}
 }
}
function Copy-SelectedSteps {
 if($grid.SelectedRows.Count -eq 0){Write-Log 'コピーする行を選択してください。';return}
 $steps=@(Get-Steps -SelectedOnly)
 $insertIndex=1+($grid.SelectedRows|Measure-Object -Property Index -Maximum).Maximum
 for($offset=0;$offset -lt $steps.Count;$offset++){
  $step=$steps[$offset];$index=[int]$insertIndex+$offset
  $grid.Rows.Insert($index,[object[]]@($step.Action,$step.Value,$step.X,$step.Y,$step.Delay,$step.Target))
  $grid.Rows[$index].Cells['ProtectedValue'].Value=$step.ProtectedValue
  $grid.Rows[$index].Cells['DateOptions'].Value=$step.DateOptions
  $grid.Rows[$index].Cells['Repeat'].Value=$step.Repeat
 }
 $grid.ClearSelection();$grid.CurrentCell=$grid.Rows[[int]$insertIndex].Cells['Action']
 for($offset=0;$offset -lt $steps.Count;$offset++){$grid.Rows[[int]$insertIndex+$offset].Selected=$true}
 Write-Log "$($steps.Count)行を選択範囲の下にコピーしました。"
}
function Export-ScenarioFile($destination, $steps) {
 Check-Steps $steps
 $encoding = [Text.UTF8Encoding]::new($true)
 $scenario = @{Version=1; Steps=@($steps)} | ConvertTo-Json -Depth 8
 $payload = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($scenario))
 if($script:AutoRunScenario){
  # The running temporary script already contains every module. Reuse it so
  # recipients can export again without the original installation directory.
  $source=[IO.File]::ReadAllText($script:RuntimeSourcePath)
 }else{
  $source = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'RpaStudio.ps1'))
  foreach($moduleName in @('ElementAutomation.ps1','WebAutomation.ps1','ExcelAutomation.ps1','DateAutomation.ps1','WindowAutomation.ps1','ImageAutomation.ps1')){
   $module=[IO.File]::ReadAllText((Join-Path $PSScriptRoot $moduleName))
   $source=$source.Replace((". (Join-Path `$PSScriptRoot '"+$moduleName+"')"),$module)
  }
 }
 $declaration='(?m)^\$script:AutoRunScenario = (?:\$null|''[A-Za-z0-9+/=]*'')\r?$'
 if([regex]::Matches($source,$declaration).Count -ne 1){throw '実行プログラムの読み込みに失敗しました。最新版のRPA Studioから実行ファイルを作成し直してください。'}
 $source = [regex]::Replace($source,$declaration, [System.Text.RegularExpressions.MatchEvaluator]{ param($match) '$script:AutoRunScenario = ''' + $payload + '''' })
 $bytes = $encoding.GetPreamble() + $encoding.GetBytes($source)
 $encoded = [Convert]::ToBase64String($bytes)
 $lines = New-Object 'System.Collections.Generic.List[string]'
 $lines.Add('Option Explicit')
 $lines.Add('Dim fso, shell, tempPath, node, stream, data, result')
 $lines.Add('Set fso = CreateObject("Scripting.FileSystemObject")')
 $lines.Add('Set shell = CreateObject("WScript.Shell")')
 $lines.Add('tempPath = fso.BuildPath(fso.GetSpecialFolder(2), fso.GetTempName & ".ps1")')
 $lines.Add('data = ""')
 for ($i=0; $i -lt $encoded.Length; $i+=120) {
  $lines.Add('data = data & "' + $encoded.Substring($i,[Math]::Min(120,$encoded.Length-$i)) + '"')
 }
 $lines.Add('Set node = CreateObject("Msxml2.DOMDocument.6.0").createElement("base64")')
 $lines.Add('node.dataType = "bin.base64"')
 $lines.Add('node.text = data')
 $lines.Add('Set stream = CreateObject("ADODB.Stream")')
 $lines.Add('stream.Type = 1')
 $lines.Add('stream.Open')
 $lines.Add('stream.Write node.nodeTypedValue')
 $lines.Add('stream.SaveToFile tempPath, 2')
 $lines.Add('stream.Close')
 $lines.Add('result = shell.Run("powershell.exe -NoProfile -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File " & Chr(34) & tempPath & Chr(34), 0, True)')
 $lines.Add('fso.DeleteFile tempPath, True')
 $lines.Add('If result <> 0 Then MsgBox "Scenario launch failed. Code: " & result, 16, "RPA Studio"')
 [IO.File]::WriteAllLines($destination,$lines.ToArray(),[Text.Encoding]::ASCII)
}
$new = New-Button '新規' {
 if ($grid.Rows.Count -gt 0 -and [Windows.Forms.MessageBox]::Show('未保存の変更を破棄して新規作成しますか？','新規','YesNo') -ne 'Yes') { return }
 $grid.Rows.Clear(); $script:path = $null
}
$open = New-Button '開く' {
 $d = New-Object Windows.Forms.OpenFileDialog; $d.Filter = 'RPAシナリオ (*.json)|*.json'
 if ($d.ShowDialog() -eq 'OK') {
  try {
   $doc = Get-Content -LiteralPath $d.FileName -Raw -Encoding UTF8 | ConvertFrom-Json
   if ($doc.Version -ne 1) { throw '対応していないシナリオ形式です。' }
   $steps = @($doc.Steps); Check-Steps $steps
   if ($grid.Rows.Count -gt 0 -and [Windows.Forms.MessageBox]::Show('現在の編集内容を置き換えますか？','開く','YesNo') -ne 'Yes') { return }
   $grid.Rows.Clear()
   foreach ($s in $steps) { Add-StoredStep $s }
   $script:path = $d.FileName; Write-Log "読込: $script:path"
  } catch { [void][Windows.Forms.MessageBox]::Show($_.Exception.Message,'読込エラー') }
 }
 $d.Dispose()
}
$save = New-Button '名前を付けて保存' {
 try {
  $steps = @(Get-Steps); Check-Steps $steps
  $d = New-Object Windows.Forms.SaveFileDialog; $d.Filter = 'RPAシナリオ (*.json)|*.json'; $d.DefaultExt = 'json'
  if ($script:path) { $d.FileName = $script:path }
  if ($d.ShowDialog() -eq 'OK') {
   @{Version=1; Steps=$steps} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $d.FileName -Encoding UTF8
   $script:path = $d.FileName; Write-Log "保存: $script:path"
  }
  $d.Dispose()
 } catch { [void][Windows.Forms.MessageBox]::Show($_.Exception.Message,'保存エラー') }
}
function Save-CurrentScenario {
 if (-not $script:path) { $save.PerformClick(); return }
 try {
  $steps = @(Get-Steps); Check-Steps $steps
  @{Version=1; Steps=$steps} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $script:path -Encoding UTF8 -ErrorAction Stop
  Write-Log "上書き保存: $script:path"
 } catch { [void][Windows.Forms.MessageBox]::Show($_.Exception.Message,'保存エラー') }
}
$export = New-Button '実行ファイル保存' {
 $d = New-Object Windows.Forms.SaveFileDialog
 $d.Filter = 'シナリオ実行ファイル (*.vbs)|*.vbs'; $d.DefaultExt='vbs'; $d.FileName='シナリオ実行.vbs'
 try {
  $steps=@(Get-Steps); Check-Steps $steps
  if ($d.ShowDialog() -eq 'OK') {
   Export-ScenarioFile $d.FileName $steps
   Write-Log "実行ファイル保存: $($d.FileName)（保存時点の全行を実行）"
  }
 } catch { [void][Windows.Forms.MessageBox]::Show($_.Exception.Message,'書き出しエラー') }
 finally { $d.Dispose() }
}
$add = New-Button '操作追加' {
 [void]$grid.EndEdit()
 $insertIndex = $grid.Rows.Count
 if ($grid.SelectedRows.Count) {
  $insertIndex = ($grid.SelectedRows | Measure-Object -Property Index -Minimum).Minimum
 }
 $grid.Rows.Insert([int]$insertIndex, [object[]]@('待機','1000','0','0','300'))
 $grid.ClearSelection()
 $grid.CurrentCell = $grid.Rows[$insertIndex].Cells['Action']
 $grid.Rows[$insertIndex].Selected = $true
}
$addBelow = New-Button '操作追加（下）' {
 [void]$grid.EndEdit()
 $insertIndex = $grid.Rows.Count
 if ($grid.SelectedRows.Count) {
  $insertIndex = 1 + ($grid.SelectedRows | Measure-Object -Property Index -Maximum).Maximum
 }
 $grid.Rows.Insert([int]$insertIndex, [object[]]@('待機','1000','0','0','300'))
 $grid.ClearSelection()
 $grid.CurrentCell = $grid.Rows[$insertIndex].Cells['Action']
 $grid.Rows[$insertIndex].Selected = $true
}
$copyRows=New-Button '行コピー' {Copy-SelectedSteps}
$windowSetup=New-Button 'ウィンドウ指定' {
 if($grid.SelectedRows.Count -ne 1){Write-Log '対象を登録する行を1行選択してください。';return}
 [void]$grid.EndEdit();$row=$grid.SelectedRows[0]
 $windows=@([RpaWindows]::List()|Where-Object{$_.ProcessId -ne $PID}|Sort-Object ProcessName,Title)
 if($windows.Count -eq 0){Write-Log '対象ウィンドウがありません。対象アプリを開いてください。';return}
 $dialog=New-Object Windows.Forms.Form
 $dialog.Text='アクティブ化するウィンドウを選択';$dialog.Size=New-Object Drawing.Size(760,360);$dialog.StartPosition='CenterParent'
 $list=New-Object Windows.Forms.ListBox;$list.Dock='Fill';$list.HorizontalScrollbar=$true
 foreach($window in $windows){[void]$list.Items.Add(($window.ProcessName+' | '+$window.Title))};$list.SelectedIndex=0
 $ok=New-Object Windows.Forms.Button;$ok.Text='指定';$ok.Dock='Bottom';$ok.Height=35;$ok.DialogResult='OK'
 $dialog.Controls.Add($list);$dialog.Controls.Add($ok);$dialog.AcceptButton=$ok
 try{
  if($dialog.ShowDialog($form) -eq 'OK'){
   $window=$windows[$list.SelectedIndex]
   $row.Cells['Action'].Value='ウィンドウをアクティブ化';$row.Cells['Value'].Value=$window.Title
   $row.Cells['Target'].Value=@{Kind='Window';Process=$window.ProcessName;Class=$window.ClassName}|ConvertTo-Json -Compress
   Write-Log "行 $($row.Index+1): 対象ウィンドウを指定しました。"
  }
 }finally{$dialog.Dispose()}
}
$imagePick=New-Button '画像取得' {
 if($grid.SelectedRows.Count -ne 1){Write-Log '画像を登録する行を1行選択してください。';return}
 [void]$grid.EndEdit();$row=$grid.SelectedRows[0]
 $previousState=$form.WindowState
 try{
  $toolbar.Enabled=$false;$form.WindowState='Minimized'
  $script:cancel=$false;Wait-Rpa 500
  $target=Capture-ImageTarget
  if($target){
   $row.Cells['Target'].Value=$target|ConvertTo-Json -Compress
   if($row.Cells['Action'].Value -notin @('画像クリック','画像ダブルクリック','画像待機')){$row.Cells['Action'].Value='画像クリック';$row.Cells['Value'].Value=''}
   Write-Log "見本画像を保存しました（$($target.Width)×$($target.Height)px）。実行時は画像の中心をクリックします。"
  }
 }catch{Write-Log "画像取得停止: $($_.Exception.Message)"}
 finally{$toolbar.Enabled=$true;$form.WindowState=$previousState}
}
$imageSetup=New-Button '画像設定' {
 if($grid.SelectedRows.Count -ne 1){Write-Log '見本画像を登録した行を1行選択してください。';return}
 [void]$grid.EndEdit();$row=$grid.SelectedRows[0]
 $decoded=$null;$dialog=$null;$preview=$null
 try{
  if($row.Cells['Action'].Value -notin @('画像クリック','画像ダブルクリック','画像待機')){throw '画像操作の行を選択してください。'}
  $json=[string]$row.Cells['Target'].Value
  $decoded=Read-ImageTarget $json;$target=$json|ConvertFrom-Json
  $dialog=New-Object Windows.Forms.Form
  $dialog.Text='画像設定';$dialog.ClientSize=New-Object Drawing.Size(560,385)
  $dialog.StartPosition='CenterParent';$dialog.FormBorderStyle='FixedDialog';$dialog.MaximizeBox=$false;$dialog.MinimizeBox=$false
  $preview=New-Object Windows.Forms.PictureBox;$preview.SetBounds(12,12,536,160);$preview.SizeMode='Zoom'
  $preview.Image=$decoded.Bitmap;$dialog.Controls.Add($preview)
  $hintLabel=New-Object Windows.Forms.Label
  $hintLabel.Text='基準は見本画像の中心です。右・下が＋、左・上が－です。'
  $hintLabel.SetBounds(12,180,536,25);$dialog.Controls.Add($hintLabel)
  $numbers=@{}
  foreach($item in @(@('OffsetX','横のずらし量（px）',210,-30000,30000,$decoded.OffsetX),@('OffsetY','縦のずらし量（px）',250,-30000,30000,$decoded.OffsetY),@('Tolerance','色の許容差（0～100）',290,0,100,$decoded.Tolerance))){
   $label=New-Object Windows.Forms.Label;$label.Text=$item[1];$label.SetBounds(12,$item[2],300,27);$dialog.Controls.Add($label)
   $number=New-Object Windows.Forms.NumericUpDown;$number.Name=$item[0];$number.Minimum=$item[3];$number.Maximum=$item[4];$number.Value=$item[5]
   $number.SetBounds(330,$item[2],218,27);$dialog.Controls.Add($number);$numbers[$item[0]]=$number
  }
  $ok=New-Object Windows.Forms.Button;$ok.Text='設定';$ok.DialogResult='OK';$ok.SetBounds(365,340,85,30)
  $cancel=New-Object Windows.Forms.Button;$cancel.Text='キャンセル';$cancel.DialogResult='Cancel';$cancel.SetBounds(460,340,88,30)
  $dialog.Controls.Add($ok);$dialog.Controls.Add($cancel);$dialog.AcceptButton=$ok;$dialog.CancelButton=$cancel
  if($dialog.ShowDialog($form) -eq 'OK'){
   foreach($name in @('OffsetX','OffsetY','Tolerance')){$target|Add-Member -NotePropertyName $name -NotePropertyValue ([int]$numbers[$name].Value) -Force}
   $row.Cells['Target'].Value=$target|ConvertTo-Json -Depth 8 -Compress
   Write-Log "画像設定: 中心から横 $($target.OffsetX)px / 縦 $($target.OffsetY)px"
  }
 }catch{[void][Windows.Forms.MessageBox]::Show($_.Exception.Message,'画像設定エラー')}
 finally{
  if($preview){$preview.Image=$null};if($dialog){$dialog.Dispose()};if($decoded){$decoded.Bitmap.Dispose()}
 }
}
$dateSetup = New-Button '日付入力設定' {
 if($grid.SelectedRows.Count -ne 1){Write-Log '日付入力を設定する行を1行選択してください。';return}
 [void]$grid.EndEdit();$row=$grid.SelectedRows[0]
 $dialog=New-Object Windows.Forms.Form
 $dialog.Text='日付入力設定';$dialog.ClientSize=New-Object Drawing.Size(570,310)
 $dialog.StartPosition='CenterParent';$dialog.FormBorderStyle='FixedDialog';$dialog.MaximizeBox=$false;$dialog.MinimizeBox=$false
 foreach($item in @(@('基準日',15),@('指定日',70),@('加減算する日数（昨日=-1、明日=1）',125),@('書式（選択または直接入力）',180))){
  $label=New-Object Windows.Forms.Label;$label.Text=$item[0];$label.SetBounds(12,$item[1],545,22);$dialog.Controls.Add($label)
 }
 $basisBox=New-Object Windows.Forms.ComboBox;$basisBox.DropDownStyle='DropDownList';$basisBox.SetBounds(12,38,545,26)
 $basisBox.Items.AddRange([object[]]@('実行日（日本時間）','指定日'));$basisBox.SelectedIndex=0;$dialog.Controls.Add($basisBox)
 $fixedPicker=New-Object Windows.Forms.DateTimePicker;$fixedPicker.Format='Custom';$fixedPicker.CustomFormat='yyyy/MM/dd';$fixedPicker.Value=Get-RpaToday
 $fixedPicker.SetBounds(12,93,250,26);$dialog.Controls.Add($fixedPicker)
 $offsetBox=New-Object Windows.Forms.NumericUpDown;$offsetBox.Minimum=-36500;$offsetBox.Maximum=36500;$offsetBox.SetBounds(12,148,250,26);$dialog.Controls.Add($offsetBox)
 $formatBox=New-Object Windows.Forms.ComboBox;$formatBox.DropDownStyle='DropDown';$formatBox.SetBounds(12,203,545,26)
 $formatBox.Items.AddRange([object[]]@('yyyy','yy','MM','M','dd','d','yyyy/MM/dd','yyyy/M/d','yyyy-MM-dd','yyyy.MM.dd','yyyyMMdd','yyyy/MM','yyyy-MM','yyyy.MM','MM/dd','M/d','MM.dd','M.d','MMdd','yyyy年M月d日','yyyy年MM月dd日','M月d日','MM月dd日','yyyy年','M月','MM月','d日','dd日','yyyy/MM/dd(ddd)'))
 $formatBox.Text='yyyy/MM/dd';$dialog.Controls.Add($formatBox)
 if($row.Cells['Action'].Value -eq '日付入力'){
  $formatBox.Text=[string]$row.Cells['Value'].Value
  try{
   $options=[string]$row.Cells['DateOptions'].Value|ConvertFrom-Json -ErrorAction Stop
   if($options.Basis -eq 'Fixed'){$basisBox.SelectedIndex=1;$fixedPicker.Value=[datetime]::ParseExact([string]$options.FixedDate,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture)}
   $offsetBox.Value=[decimal]$options.OffsetDays
  }catch{}
 }
 $preview=New-Object Windows.Forms.Label;$preview.SetBounds(12,240,545,30);$dialog.Controls.Add($preview)
 $updatePreview={
  $fixedPicker.Enabled=$basisBox.SelectedIndex -eq 1
  try{
   $settings=@{Basis=if($basisBox.SelectedIndex -eq 1){'Fixed'}else{'Today'};FixedDate=$fixedPicker.Value.ToString('yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture);OffsetDays=[int]$offsetBox.Value}|ConvertTo-Json -Compress
   $sample=Get-DateInputText ([pscustomobject]@{Value=$formatBox.Text;DateOptions=$settings})
   $preview.Text='入力例: '+$sample;$preview.ForeColor=[Drawing.Color]::DarkSlateBlue
  }catch{$preview.Text='入力例: 設定を確認してください';$preview.ForeColor=[Drawing.Color]::Firebrick}
 }
 $basisBox.Add_SelectedIndexChanged($updatePreview);$fixedPicker.Add_ValueChanged($updatePreview)
 $offsetBox.Add_ValueChanged($updatePreview);$formatBox.Add_TextChanged($updatePreview)
 & $updatePreview
 $ok=New-Object Windows.Forms.Button;$ok.Text='設定';$ok.SetBounds(375,273,85,28)
 $cancel=New-Object Windows.Forms.Button;$cancel.Text='キャンセル';$cancel.DialogResult='Cancel';$cancel.SetBounds(470,273,87,28)
 $dialog.Controls.Add($ok);$dialog.Controls.Add($cancel);$dialog.AcceptButton=$ok;$dialog.CancelButton=$cancel
 $ok.Add_Click({
  try{
   $settings=@{Basis=if($basisBox.SelectedIndex -eq 1){'Fixed'}else{'Today'};FixedDate=$fixedPicker.Value.ToString('yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture);OffsetDays=[int]$offsetBox.Value}|ConvertTo-Json -Compress
   $null=Get-DateInputText ([pscustomobject]@{Value=$formatBox.Text;DateOptions=$settings})
   $dialog.DialogResult='OK'
  }catch{[void][Windows.Forms.MessageBox]::Show($_.Exception.Message,'日付設定エラー')}
 })
 try{
  if($dialog.ShowDialog($form) -eq 'OK'){
   $settings=@{Basis=if($basisBox.SelectedIndex -eq 1){'Fixed'}else{'Today'};FixedDate=$fixedPicker.Value.ToString('yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture);OffsetDays=[int]$offsetBox.Value}|ConvertTo-Json -Compress
   if($row.Cells['Action'].Value -notin @('日付入力','要素入力','ウェブ要素入力','Excelセル入力','ID入力','パスワード入力','要素クリック','ウェブ要素クリック','要素フォーカス')){$row.Cells['Target'].Value=''}
   $row.Cells['Action'].Value='日付入力';$row.Cells['Value'].Value=$formatBox.Text;$row.Cells['DateOptions'].Value=$settings
   Write-Log "行 $($row.Index+1): 日付入力を設定しました。"
  }
 }finally{$dialog.Dispose()}
}
$excelSetup = New-Button 'Excel操作設定' {
 if($grid.SelectedRows.Count -ne 1){Write-Log 'Excel操作を設定する行を1行選択してください。';return}
 [void]$grid.EndEdit();$row=$grid.SelectedRows[0]
 $isDateRow=$row.Cells['Action'].Value -eq '日付入力'
 $dialog=New-Object Windows.Forms.Form
 $dialog.Text='Excel操作設定';$dialog.ClientSize=New-Object Drawing.Size(570,270)
 $dialog.StartPosition='CenterParent';$dialog.FormBorderStyle='FixedDialog';$dialog.MaximizeBox=$false;$dialog.MinimizeBox=$false
 foreach($item in @(@('操作',15),@('値（ファイルパス・シート名・入力文字）',70),@('シート名（セル入力用・空欄は選択中）',125),@('セル（セル入力用）',180))){
  $label=New-Object Windows.Forms.Label;$label.Text=$item[0];$label.SetBounds(12,$item[1],545,22);$dialog.Controls.Add($label)
 }
 $kind=New-Object Windows.Forms.ComboBox;$kind.DropDownStyle='DropDownList';$kind.SetBounds(12,38,545,26)
 $kind.Items.AddRange([object[]]$excelActions);$kind.SelectedIndex=0
 if($row.Cells['Action'].Value -in $excelActions){$kind.SelectedItem=$row.Cells['Action'].Value}
 if($isDateRow){$kind.SelectedItem='Excelセル入力'}
 $dialog.Controls.Add($kind)
 $valueBox=New-Object Windows.Forms.TextBox;$valueBox.SetBounds(12,93,450,26);$dialog.Controls.Add($valueBox)
 if($row.Cells['Action'].Value -in $excelActions -or $isDateRow){$valueBox.Text=[string]$row.Cells['Value'].Value}
 $browse=New-Object Windows.Forms.Button;$browse.Text='参照';$browse.SetBounds(470,92,87,28);$dialog.Controls.Add($browse)
 $sheetBox=New-Object Windows.Forms.TextBox;$sheetBox.SetBounds(12,148,545,26);$dialog.Controls.Add($sheetBox)
 $cellBox=New-Object Windows.Forms.TextBox;$cellBox.Text='A1';$cellBox.SetBounds(12,203,260,26);$dialog.Controls.Add($cellBox)
 try{
  $target=[string]$row.Cells['Target'].Value|ConvertFrom-Json -ErrorAction Stop
  if($target.Kind -eq 'Excel'){$sheetBox.Text=[string]$target.Sheet;$cellBox.Text=[string]$target.Cell}
 }catch{}
 $browse.Add_Click({
  $fileDialog=New-Object Windows.Forms.OpenFileDialog;$fileDialog.Filter='Excelブック|*.xlsx;*.xlsm;*.xlsb;*.xls'
  try{if($fileDialog.ShowDialog($dialog) -eq 'OK'){$valueBox.Text=$fileDialog.FileName}}finally{$fileDialog.Dispose()}
 })
 $kind.Add_SelectedIndexChanged({
  $isCell=$kind.SelectedItem -eq 'Excelセル入力';$sheetBox.Enabled=$isCell;$cellBox.Enabled=$isCell
  $browse.Enabled=$kind.SelectedItem -eq 'Excelブックを開く'
  $valueBox.Enabled=$kind.SelectedItem -notin @('Excelブック保存','Excelブックを閉じる')
 })
 $sheetBox.Enabled=$kind.SelectedItem -eq 'Excelセル入力';$cellBox.Enabled=$sheetBox.Enabled
 $browse.Enabled=$kind.SelectedItem -eq 'Excelブックを開く'
 $valueBox.Enabled=$kind.SelectedItem -notin @('Excelブック保存','Excelブックを閉じる')
 $ok=New-Object Windows.Forms.Button;$ok.Text='設定';$ok.SetBounds(375,236,85,28)
 $cancel=New-Object Windows.Forms.Button;$cancel.Text='キャンセル';$cancel.DialogResult='Cancel';$cancel.SetBounds(470,236,87,28)
 $dialog.Controls.Add($ok);$dialog.Controls.Add($cancel);$dialog.AcceptButton=$ok;$dialog.CancelButton=$cancel
 $ok.Add_Click({
  if($kind.SelectedItem -eq 'Excelセル入力' -and $cellBox.Text -notmatch '^\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}$'){
   [void][Windows.Forms.MessageBox]::Show('セルはA1やB2の形式で指定してください。','入力エラー');return
  }
  if($kind.SelectedItem -in @('Excelブックを開く','Excelシート選択') -and [string]::IsNullOrWhiteSpace($valueBox.Text)){
   [void][Windows.Forms.MessageBox]::Show('値を入力してください。','入力エラー');return
  }
  $dialog.DialogResult='OK'
 })
 try{
  if($dialog.ShowDialog($form) -eq 'OK'){
   $row.Cells['Action'].Value=if($isDateRow -and $kind.SelectedItem -eq 'Excelセル入力'){'日付入力'}else{[string]$kind.SelectedItem}
   $row.Cells['Value'].Value=if($valueBox.Enabled){$valueBox.Text}else{''}
   $row.Cells['Target'].Value=if($kind.SelectedItem -eq 'Excelセル入力'){@{Kind='Excel';Sheet=$sheetBox.Text;Cell=$cellBox.Text}|ConvertTo-Json -Compress}else{''}
   Write-Log "行 $($row.Index+1): Excel操作を設定しました。"
  }
 }finally{$dialog.Dispose()}
}
$configureInput = New-Button '入力情報設定' {
 if($grid.SelectedRows.Count -ne 1){Write-Log 'ID入力またはパスワード入力の行を1行選択してください。'; return}
 [void]$grid.EndEdit()
 $row=$grid.SelectedRows[0]
 if($row.Cells['Action'].Value -notin @('ID入力','パスワード入力')){Write-Log '操作を「ID入力」または「パスワード入力」に変更してください。'; return}
 $dialog=New-Object Windows.Forms.Form
 $dialog.Text=[string]$row.Cells['Action'].Value + 'の設定'
 $dialog.ClientSize=New-Object Drawing.Size(460,155)
 $dialog.StartPosition='CenterParent'; $dialog.FormBorderStyle='FixedDialog'; $dialog.MaximizeBox=$false; $dialog.MinimizeBox=$false
 $label=New-Object Windows.Forms.Label
 $label.Text='入力内容を登録します（既存の内容は表示しません）。'
 $label.SetBounds(12,12,435,25); $dialog.Controls.Add($label)
 $secretBox=New-Object Windows.Forms.TextBox
 $secretBox.UseSystemPasswordChar=$true; $secretBox.SetBounds(12,43,435,26)
 $dialog.Controls.Add($secretBox)
 $note=New-Object Windows.Forms.Label
 $note.Text='同じPC・Windowsユーザー用に暗号化して保存します。'
 $note.SetBounds(12,78,435,25); $dialog.Controls.Add($note)
 $ok=New-Object Windows.Forms.Button; $ok.Text='登録'; $ok.DialogResult='OK'; $ok.SetBounds(265,114,85,28)
 $cancel=New-Object Windows.Forms.Button; $cancel.Text='キャンセル'; $cancel.DialogResult='Cancel'; $cancel.SetBounds(360,114,85,28)
 $dialog.Controls.Add($ok); $dialog.Controls.Add($cancel); $dialog.AcceptButton=$ok; $dialog.CancelButton=$cancel
 try {
  if($dialog.ShowDialog($form) -eq 'OK') {
   if($secretBox.Text.Length -eq 0){throw '入力内容が空です。登録し直してください。'}
   $row.Cells['ProtectedValue'].Value=Protect-InputValue $secretBox.Text
   $row.Cells['Value'].Value='[暗号化済み]'
   Write-Log "行 $($row.Index+1): 入力情報を暗号化して登録しました。"
  }
 } catch { [void][Windows.Forms.MessageBox]::Show('入力情報を登録できませんでした。空欄でないこと、Windowsの暗号化機能が利用できることを確認してください。','登録エラー') }
 finally { $secretBox.Clear(); $dialog.Dispose() }
}
$remove = New-Button '削除' {
 foreach ($i in @($grid.SelectedRows | ForEach-Object { $_.Index } | Sort-Object -Descending)) { $grid.Rows.RemoveAt($i) }
}
function Move-Step($offset) {
 if ($grid.SelectedRows.Count -ne 1) { Write-Log '並べ替える行を1行選択してください。'; return }
 $i = $grid.SelectedRows[0].Index; $j = $i + $offset
 if ($j -lt 0 -or $j -ge $grid.Rows.Count) { return }
 [void]$grid.EndEdit()
 $first=New-Object object[] $grid.Columns.Count
 $second=New-Object object[] $grid.Columns.Count
 for ($k=0; $k -lt $grid.Columns.Count; $k++) {
  $first[$k]=$grid.Rows[$i].Cells[$k].Value; $second[$k]=$grid.Rows[$j].Cells[$k].Value
 }
 for ($k=0; $k -lt $grid.Columns.Count; $k++) {
  $grid.Rows[$i].Cells[$k].Value=$second[$k]; $grid.Rows[$j].Cells[$k].Value=$first[$k]
 }
 $grid.ClearSelection(); $grid.Rows[$j].Selected = $true
}
$up = New-Button '↑' { Move-Step -1 }
$down = New-Button '↓' { Move-Step 1 }
$capture = New-Button '座標取得(3秒後)' {
 if ($grid.SelectedRows.Count -ne 1) { Write-Log '座標を設定する行を1行選択してください。'; return }
 $script:cancel = $false
 $toolbar.Enabled = $false
 try {
  Write-Log '3秒以内に対象へマウスを移動してください。'; Wait-Rpa 3000
  $p = [Windows.Forms.Cursor]::Position
  if ($grid.SelectedRows.Count) { $r=$grid.SelectedRows[0]; $r.Cells['X'].Value=[string]$p.X; $r.Cells['Y'].Value=[string]$p.Y }
  Write-Log "座標: X=$($p.X), Y=$($p.Y)"
 } catch { Write-Log $_.Exception.Message } finally { $toolbar.Enabled = $true }
}
$pickElement = New-Button '要素取得(3秒後)' {
 if ($grid.SelectedRows.Count -ne 1) { Write-Log '対象を設定する行を1行選択してください。'; return }
 [void]$grid.EndEdit()
 $row = $grid.SelectedRows[0]
 $script:cancel=$false; $toolbar.Enabled=$false; $grid.Enabled=$false
 try {
  Write-Log '3秒以内に対象のボタンや入力欄へマウスを合わせてください。'
  Wait-Rpa 3000
  $p=[Windows.Forms.Cursor]::Position
  $element=[Windows.Automation.AutomationElement]::FromPoint([Windows.Point]::new($p.X,$p.Y))
  $target=Get-ElementTarget $element
  if ($element.Current.ProcessId -eq $PID) { throw 'RPA Studio自身は指定できません。対象アプリへマウスを移動してください。' }
  $row.Cells['Target'].Value=($target | ConvertTo-Json -Compress)
  $pattern=$null
  if ($row.Cells['Action'].Value -notin @('要素クリック','要素入力','要素フォーカス','ID入力','パスワード入力','日付入力')) {
   if ($element.TryGetCurrentPattern([Windows.Automation.ValuePattern]::Pattern,[ref]$pattern)) { $row.Cells['Action'].Value='要素入力'; $row.Cells['Value'].Value='' }
   elseif ($element.TryGetCurrentPattern([Windows.Automation.InvokePattern]::Pattern,[ref]$pattern)) { $row.Cells['Action'].Value='要素クリック'; $row.Cells['Value'].Value='' }
   else { $row.Cells['Action'].Value='要素フォーカス' }
  }
  Write-Log "対象取得: $($target.Name) / $($target.Type) / $($target.Window)"
 } catch { Write-Log "取得失敗: $($_.Exception.Message)" }
 finally { $toolbar.Enabled=$true; $grid.Enabled=$true }
}
$webBrowser = New-Button 'Webブラウザ起動' {
 $script:cancel=$false
 $toolbar.Enabled=$false
 try{
  $url='about:blank'
  if($grid.SelectedRows.Count -eq 1 -and $grid.SelectedRows[0].Cells['Action'].Value -in @('URLを開く','Web用URLを開く')){
   $url=[string]$grid.SelectedRows[0].Cells['Value'].Value
  }
  Start-RpaWebBrowser $url
  Write-Log '専用Edgeを起動しました。対象サイトを開いて「ウェブ要素取得」を押してください。'
 }catch{Write-Log "Webブラウザ起動失敗: $($_.Exception.Message)"}
 finally{$toolbar.Enabled=$true}
}
$webPick = New-Button 'ウェブ要素取得' {
 if($grid.SelectedRows.Count -ne 1){Write-Log '対象を登録する行を1行選択してください。';return}
 [void]$grid.EndEdit();$row=$grid.SelectedRows[0]
 $script:cancel=$false
 try{
  $page=Select-WebPage
  if(-not $page){return}
  $script:running=$true;$grid.Enabled=$false
  foreach($button in $toolbar.Controls){$button.Enabled=$false};$stop.Enabled=$true
  Write-Log '専用Edgeの対象要素をクリックしてください（60秒以内）。Escまたは実行停止で中止できます。'
  $target=Capture-WebTarget $page
  $row.Cells['Target'].Value=$target|ConvertTo-Json -Depth 8 -Compress
  if($row.Cells['Action'].Value -notin @('ID入力','パスワード入力','ウェブ要素入力','日付入力')){$row.Cells['Action'].Value='ウェブ要素クリック';$row.Cells['Value'].Value=''}
  Write-Log "Web要素取得: tag=$($target.Tag) / id=$($target.Id) / class=$($target.Classes -join ' ') / selector=$($target.Selector)"
  Write-Log 'WebシナリオのURL操作には「Web用URLを開く」を使用してください。'
 }catch{Write-Log "Web要素取得停止: $($_.Exception.Message)"}
 finally{
  $script:running=$false;$grid.Enabled=$true
  foreach($button in $toolbar.Controls){$button.Enabled=$true};$stop.Enabled=$false
 }
}
$webClick = New-Button 'ウェブ要素クリック' {
 if($grid.SelectedRows.Count -ne 1){Write-Log 'Web要素を取得済みの行を1行選択してください。';return}
 [void]$grid.EndEdit();$row=$grid.SelectedRows[0]
 try{
  $target=[string]$row.Cells['Target'].Value|ConvertFrom-Json -ErrorAction Stop
  if($target.Kind -ne 'Web' -or -not $target.Selector){throw '「ウェブ要素取得」で対象を取得してください。'}
  $row.Cells['Action'].Value='ウェブ要素クリック';$row.Cells['Value'].Value=''
  Write-Log '選択行を「ウェブ要素クリック」に設定しました。「選択行を実行」でクリックできます。'
 }catch{Write-Log $_.Exception.Message}
}
function Invoke-Scenario([switch]$SelectedOnly) {
 try {
  if ($SelectedOnly -and $grid.SelectedRows.Count -eq 0) { throw '実行する行を選択してください。' }
  $steps = @(Get-Steps -SelectedOnly:$SelectedOnly); Check-Steps $steps -DecryptSecrets
  $useWebBrowser=@($steps | Where-Object {
   $_.Action -in @('Web用URLを開く','ウェブ要素クリック','ウェブ要素入力') -or
   ($_.Action -in @('ID入力','パスワード入力','日付入力') -and $_.Target -and ($_.Target|ConvertFrom-Json).Kind -eq 'Web')
  }).Count -gt 0
 } catch { [void][Windows.Forms.MessageBox]::Show($_.Exception.Message,'入力エラー'); return }
 $script:running = $true; $script:cancel = $false
 $exitRequested=$false
 $script:ScenarioDateBase=Get-RpaToday
 $activeStep=$null
 foreach($row in $grid.Rows) { $row.Cells['State'].Value=''; $row.DefaultCellStyle=New-Object Windows.Forms.DataGridViewCellStyle }
 $executionLabel.Text='開始待ち（3秒）'
 $grid.ReadOnly = $true
 foreach ($b in $toolbar.Controls) { $b.Enabled = $false }; $stop.Enabled = $true
 try {
  Write-Log '3秒後に実行します。操作対象を手前に表示してください。'; Wait-Rpa 3000
  $index = 0
  foreach ($s in $steps) {
   Wait-Rpa 0
   $activeStep=$s; $index++
   Set-StepState $s.RowNumber '実行中'
   $executionLabel.Text="実行中: 行 $($s.RowNumber) / $($s.Action)（$index / $($steps.Count)）"
   Write-Log "行 $($s.RowNumber) ($index / $($steps.Count)) : $($s.Action)"
   switch ($s.Action) {
    { $_ -in @('画像クリック','画像ダブルクリック','画像待機') } {Invoke-ImageStep $s}
    'ウィンドウをアクティブ化' {Invoke-WindowActivation $s}
    'RPA終了' { }
    '日付入力' { Invoke-DateInput $s }
    { $_ -in $excelActions } { Invoke-ExcelStep $s }
    'Web用URLを開く' { Start-RpaWebBrowser $s.Value }
    'ウェブ要素クリック' { Invoke-WebStep $s }
    'ウェブ要素入力' { Invoke-WebStep $s -InputMode -inputValue $s.Value }
    { $_ -in @('ID入力','パスワード入力') } { Invoke-ProtectedInput $s }
    { $_ -in @('要素クリック','要素入力','要素フォーカス') } { Invoke-ElementStep $s }
    '待機' { Wait-Rpa ([int]$s.Value) }
    { $_ -in @('マウス移動','左クリック','ダブルクリック','右クリック') } {
     [void][RpaNative]::SetCursorPos([int]$s.X,[int]$s.Y)
     Wait-Rpa 80
     if ($s.Action -eq '左クリック') { Mouse-Click 2 4 }
     if ($s.Action -eq '右クリック') { Mouse-Click 8 16 }
     if ($s.Action -eq 'ダブルクリック') { Mouse-Click 2 4; Wait-Rpa 80; Mouse-Click 2 4 }
    }
    '文字入力' {
     if ($s.Value.Length) {
      [Windows.Forms.Clipboard]::SetText($s.Value)
      [Windows.Forms.SendKeys]::SendWait('^v')
     }
    }
    'キー送信' { Invoke-KeyStep $s }
    'URLを開く' {
     if($useWebBrowser){
      Write-Log 'Web操作を含むため、このURLは専用Edgeで開きます。'
      Start-RpaWebBrowser $s.Value
     } else { Start-Process -FilePath $s.Value -ErrorAction Stop }
    }
    'アプリ起動' { Start-Process -FilePath $s.Value -ErrorAction Stop }
   }
   Wait-Rpa ([int]$s.Delay)
   Set-StepState $s.RowNumber '完了'
   $activeStep=$null
   if($s.Action -eq 'RPA終了'){$exitRequested=$true;break}
  }
  $executionLabel.Text="完了（$index 行）"
  Write-Log '実行が完了しました。'
 } catch {
  $cancelled=$script:cancel -or $_.Exception -is [OperationCanceledException]
  $state=if($cancelled){'中断'}else{'エラー'}
  if($null -ne $activeStep) {
   Set-StepState $activeStep.RowNumber $state
   $executionLabel.Text="$state : 行 $($activeStep.RowNumber) / $($activeStep.Action)"
   Write-Log "$state : 行 $($activeStep.RowNumber) / $($activeStep.Action) : $($_.Exception.Message)"
  } else {
   $executionLabel.Text=$state
   Write-Log "$state : $($_.Exception.Message)"
  }
 }
 finally {
  $script:running=$false; $grid.Enabled=$true; $grid.ReadOnly=$false
  $script:ScenarioDateBase=$null
  foreach ($b in $toolbar.Controls) { $b.Enabled=$true }; $stop.Enabled=$false
 }
 if($exitRequested){$form.Close()}
}
$run = New-Button '全体実行' { Invoke-Scenario }
$runSelected = New-Button '選択行を実行' { Invoke-Scenario -SelectedOnly }
$stop = New-Button '実行停止' { $script:cancel = $true }; $stop.Enabled = $false
$form.KeyPreview = $true
$form.Add_KeyDown({
 param($sender, $e)
 if ($e.Control -and -not $e.Alt -and -not $e.Shift -and $e.KeyCode -eq [Windows.Forms.Keys]::S) {
  $e.Handled = $true
  $e.SuppressKeyPress = $true
  if ($save.Enabled -and $toolbar.Enabled -and -not $script:running) { Save-CurrentScenario }
 }
})
$form.Add_FormClosing({ param($sender,$e)
 if ($script:running) { $script:cancel = $true; $e.Cancel = $true }
})
Write-Log 'キー例: Space={SPACE} / Enter={ENTER} / Tab={TAB} / Ctrl+A=^a / Ctrl+L=^l'
Write-Log '文字入力はクリップボードを書き換えます。URLは既定のブラウザで開きます。'
if ($script:AutoRunScenario) {
 $form.Text='PowerShell RPA シナリオ実行'
 $form.Add_Shown({
  try {
   $doc=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($script:AutoRunScenario)) | ConvertFrom-Json
   $steps=@($doc.Steps); Check-Steps $steps
   foreach($s in $steps) { Add-StoredStep $s }
   Invoke-Scenario
  } catch { Write-Log "起動失敗: $($_.Exception.Message)" }
 })
}
[void]$form.ShowDialog()
$positionTimer.Stop()
$positionTimer.Dispose()
$form.Dispose()
