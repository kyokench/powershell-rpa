Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName WindowsBase

function Get-ElementTarget($element) {
 $ErrorActionPreference = 'Stop'
 if ($null -eq $element) { throw '対象要素を取得できませんでした。' }
 $walker = [Windows.Automation.TreeWalker]::ControlViewWalker
 $window = $element
 while ($true) {
  $parent = $walker.GetParent($window)
  if ($null -eq $parent) { break }
  if ([Windows.Automation.Automation]::Compare($parent,[Windows.Automation.AutomationElement]::RootElement)) { break }
  $window = $parent
 }
 $info = $element.Current
 if (-not $info.AutomationId -and -not $info.Name) { throw '名前とAutomationIdを取得できません。別の要素を指定してください。' }
 $process = Get-Process -Id $window.Current.ProcessId -ErrorAction Stop
 [pscustomobject]@{
  Process=$process.ProcessName; Window=$window.Current.Name
  Id=$info.AutomationId; Name=$info.Name; Type=$info.ControlType.ProgrammaticName
 }
}
function Find-TargetElement($target) {
 $root = [Windows.Automation.AutomationElement]::RootElement
 $windows = $root.FindAll([Windows.Automation.TreeScope]::Children,[Windows.Automation.Condition]::TrueCondition)
 $matches = New-Object 'System.Collections.Generic.List[object]'
 foreach ($window in $windows) {
  try { $process = Get-Process -Id $window.Current.ProcessId -ErrorAction Stop } catch { continue }
  # Window is retained as capture metadata; titles can change after navigation.
  if ($process.ProcessName -ne $target.Process) { continue }
  $conditions = New-Object 'System.Collections.Generic.List[Windows.Automation.Condition]'
  if ($target.Id) { $conditions.Add([Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::AutomationIdProperty,[string]$target.Id)) }
  if ($target.Name) { $conditions.Add([Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::NameProperty,[string]$target.Name)) }
  if ($conditions.Count -eq 1) { $condition=$conditions[0] }
  else { $condition=[Windows.Automation.AndCondition]::new($conditions.ToArray()) }
  $elements=$window.FindAll([Windows.Automation.TreeScope]::Subtree,$condition)
  foreach ($element in $elements) {
   if ($element.Current.ControlType.ProgrammaticName -eq $target.Type) { $matches.Add($element) }
  }
 }
 if ($matches.Count -gt 1) { throw "同じアプリ内に対象が $($matches.Count) 件見つかりました。対象以外のウィンドウを閉じるか、識別情報を見直してください。" }
 if ($matches.Count -eq 1) { return $matches[0] }
 return $null
}
function Resolve-TargetElement([string]$targetJson) {
 $target = $targetJson | ConvertFrom-Json -ErrorAction Stop
 $clock = [Diagnostics.Stopwatch]::StartNew()
 do {
  Wait-Rpa 0
  $element = Find-TargetElement $target
  if ($null -ne $element) { break }
  Wait-Rpa 200
 } while ($clock.ElapsedMilliseconds -lt 10000)
 if ($null -eq $element) { throw '対象が見つかりません（10秒）。対象アプリ・ページ・要素の名前やAutomationIdを確認してください。ウィンドウ名は検索条件に使用していません。' }
 if (-not $element.Current.IsEnabled) { throw '対象は無効状態です。' }
 return $element
}
function Invoke-ElementStep($step) {
 $element = Resolve-TargetElement $step.Target
 $pattern = $null
 switch ($step.Action) {
  '要素クリック' {
   if (-not $element.TryGetCurrentPattern([Windows.Automation.InvokePattern]::Pattern,[ref]$pattern)) { throw 'この要素はボタン実行に対応していません。ボタン本体を取得してください。' }
   ([Windows.Automation.InvokePattern]$pattern).Invoke()
  }
  '要素入力' {
   if (-not $element.TryGetCurrentPattern([Windows.Automation.ValuePattern]::Pattern,[ref]$pattern)) { throw 'この入力欄は直接入力に対応していません。' }
   if ($pattern.Current.IsReadOnly) { throw '対象は読み取り専用です。' }
   ([Windows.Automation.ValuePattern]$pattern).SetValue([string]$step.Value)
  }
  '要素フォーカス' { $element.SetFocus() }
 }
}
