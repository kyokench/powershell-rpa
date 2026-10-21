#requires -Version 5.1
$ErrorActionPreference='Stop'
$test=@'
$ErrorActionPreference='Stop'
function Get-RpaToday { [datetime]::new(2026,10,22) }
$day=[datetime]::new(2026,10,22)
$cases=@(
 @('yyyy','2026'),@('yy','26'),@('MM','10'),@('M','10'),@('dd','22'),@('d','22'),
 @('yyyy/MM/dd','2026/10/22'),@('yyyy.MM.dd','2026.10.22'),@('yyyy-MM-dd','2026-10-22'),@('yyyyMMdd','20261022'),
 @('MM/dd','10/22'),@('MM.dd','10.22'),@('MMdd','1022'),@('yyyy年M月d日','2026年10月22日'),@('yyyy/MM/dd(ddd)','2026/10/22(木)')
)
foreach($case in $cases){
 $actual=Get-DateInputText ([pscustomobject]@{Value=$case[0];DateOptions=''}) -Today $day
 if($actual -cne $case[1]){throw ('Format mismatch: '+$case[0])}
}
foreach($case in @(@('M','2'),@('MM','02'),@('d','3'),@('dd','03'),@('M/d','2/3'),@('M.d','2.3'))){
 if((Get-DateInputText ([pscustomobject]@{Value=$case[0];DateOptions=''}) -Today ([datetime]::new(2026,2,3))) -cne $case[1]){throw 'Zero padding failed'}
}
$step=[pscustomobject]@{Value='yyyy/MM/dd';DateOptions='{"Basis":"Fixed","FixedDate":"2024-02-28","OffsetDays":1}'}
if((Get-DateInputText $step) -ne '2024/02/29'){throw 'Leap day failed'}
$step.DateOptions='{"Basis":"Fixed","FixedDate":"2026-01-01","OffsetDays":-1}'
if((Get-DateInputText $step) -ne '2025/12/31'){throw 'Year boundary failed'}
$step.DateOptions='{"Basis":"Today","OffsetDays":1}'
if((Get-DateInputText $step -Today $day) -ne '2026/10/23'){throw 'Relative day failed'}
$step.DateOptions='{"Basis":"Fixed","FixedDate":"2026-02-30","OffsetDays":0}'
$refused=$false;try{$null=Get-DateInputText $step}catch{$refused=$true}
if(-not $refused){throw 'Invalid date accepted'}
$step.DateOptions='';$step.Value='Q'
$refused=$false;try{$null=Get-DateInputText $step}catch{$refused=$true}
if(-not $refused){throw 'Invalid format accepted'}
Write-Output 'PASS: date formats, single month/day, zero padding, leap day, relative dates, invalid settings'
$form.Show()
$script:captured=New-Object 'System.Collections.Generic.List[string]'
function Invoke-WebStep($step,[string]$inputValue,[switch]$InputMode){$script:captured.Add('Web:'+$inputValue)}
function Invoke-ExcelStep($step){$script:captured.Add('Excel:'+$step.Value)}
function Invoke-ElementStep($step){$script:captured.Add('App:'+$step.Value)}
function Wait-Rpa([int]$ms){}
$targets=@(
 '{"Kind":"Web","PageUrl":"https://example.com/","Selector":"#date"}',
 '{"Kind":"Excel","Sheet":"Sheet1","Cell":"A1"}',
 '{"Process":"Test","Window":"Test","Type":"ControlType.Edit","Name":"Date","Id":"date"}'
)
foreach($target in $targets){
 Add-StoredStep ([pscustomobject]@{Action='日付入力';Value='yyyy.MM.dd';X='0';Y='0';Delay='0';Target=$target;DateOptions='{"Basis":"Today","OffsetDays":0}'})
}
Check-Steps @(Get-Steps)
Invoke-Scenario
if($executionLabel.Text -notmatch '^完了' -or ($script:captured -join ',') -ne 'Web:2026.10.22,Excel:2026.10.22,App:2026.10.22'){throw 'Date input dispatch failed'}
$grid.ClearSelection();$grid.Rows[2].Selected=$true;Move-Step -1
$saved=@{Version=1;Steps=@(Get-Steps)}|ConvertTo-Json -Depth 8|ConvertFrom-Json
$grid.Rows.Clear();foreach($s in $saved.Steps){Add-StoredStep $s}
if($grid.Rows[0].Cells['DateOptions'].Value -notmatch 'Today'){throw 'Date options lost on reload'}
$grid.ClearSelection();$grid.Rows[0].Selected=$true
$dialogTimer=New-Object Windows.Forms.Timer;$dialogTimer.Interval=100
$dialogTimer.Add_Tick({
 $dialog=@([Windows.Forms.Application]::OpenForms|Where-Object{$_.Text -eq '日付入力設定'})[0]
 if(-not $dialog){return}
 $dialogTimer.Stop()
 $basis=@($dialog.Controls|Where-Object{$_ -is [Windows.Forms.ComboBox] -and $_.Top -eq 38})[0]
 $format=@($dialog.Controls|Where-Object{$_ -is [Windows.Forms.ComboBox] -and $_.Top -eq 203})[0]
 $picker=@($dialog.Controls|Where-Object{$_ -is [Windows.Forms.DateTimePicker]})[0]
 $basis.SelectedIndex=1;$picker.Value=[datetime]::new(2026,2,3);$format.Text='M'
 $preview=@($dialog.Controls|Where-Object{$_ -is [Windows.Forms.Label] -and $_.Top -eq 240})[0]
 if($preview.Text -ne '入力例: 2'){throw 'Date preview failed'}
 (@($dialog.Controls|Where-Object{$_ -is [Windows.Forms.Button] -and $_.Text -eq '設定'})[0]).PerformClick()
})
$dialogTimer.Start();$dateSetup.PerformClick();$dialogTimer.Dispose()
if($grid.Rows[0].Cells['Value'].Value -ne 'M' -or $grid.Rows[0].Cells['DateOptions'].Value -notmatch 'Fixed' -or $grid.Rows[0].Cells['Target'].Value -notmatch 'Web'){throw 'Date dialog setting failed'}
Write-Output 'PASS: settings dialog, live preview and target preservation'
$export=Join-Path $PSScriptRoot '_date-test-export.vbs'
try{
 Export-ScenarioFile $export @(Get-Steps)
 $vbs=[IO.File]::ReadAllText($export)
 $encoded=([regex]::Matches($vbs,'data = data & "([A-Za-z0-9+/=]+)"')|ForEach-Object{$_.Groups[1].Value}) -join ''
 $source=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($encoded))
 $errors=$null;[void][Management.Automation.Language.Parser]::ParseInput($source,[ref]$null,[ref]$errors)
 if($errors.Count){throw ($errors|Out-String)}
 if($source -match '(?m)^\. \(Join-Path \$PSScriptRoot'){throw 'Date module not embedded'}
 $line=@($source -split "`n"|Where-Object{$_.StartsWith('$script:AutoRunScenario = ')})[0]
 $doc=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($line.Split([char]39)[1]))|ConvertFrom-Json
 if(-not $doc.Steps[0].DateOptions){throw 'Date settings not exported'}
}finally{Remove-Item -LiteralPath $export -Force -ErrorAction SilentlyContinue}
$form.Hide()
Write-Output 'PASS: Web/Excel/app dispatch, scenario execution, reorder and reload, standalone export'
'@
$path=Join-Path $PSScriptRoot '_date-verify.ps1'
try{
 $source=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'RpaStudio.ps1'))
 [IO.File]::WriteAllText($path,$source.Replace('[void]$form.ShowDialog()',$test),[Text.UTF8Encoding]::new($true))
 & powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File $path
 if($LASTEXITCODE -ne 0){throw 'Date verification failed'}
}finally{Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue}
