#requires -Version 5.1
$ErrorActionPreference='Stop'
$test=@'
$ErrorActionPreference='Stop'
foreach($case in @(
 @([datetime]::new(2026,12,1),'2026年11月','A30'),
 @([datetime]::new(2026,11,1),'2026年10月','A31'),
 @([datetime]::new(2027,1,1),'2026年12月','A31'),
 @([datetime]::new(2028,3,1),'2028年2月','A29'),
 @([datetime]::new(2027,3,1),'2027年2月','A28')
)){
 if((Expand-RpaDateTokens '{年月-1}' $case[0]) -cne $case[1]){throw 'Previous month failed'}
 if((Resolve-ExcelCellAddress 'A{日-1}' $case[0]) -cne $case[2]){throw 'Previous month-end row failed'}
}
if((Resolve-ExcelCellAddress 'A{日+1}' ([datetime]::new(2026,11,30))) -ne 'A1'){throw 'Next day failed'}
$bad=@([pscustomobject]@{Action='Else';Value=''},[pscustomobject]@{Action='EndIf';Value=''})
$rejected=$false;try{Assert-ScenarioFlow $bad}catch{$rejected=$true}
if(-not $rejected){throw 'Unbalanced flow accepted'}
$form.Show()
function Wait-Rpa([int]$ms){}
function Get-RpaToday {$script:testDay}
$script:written=New-Object 'System.Collections.Generic.List[string]'
function Invoke-ExcelStep($step){
 $t=$step.Target|ConvertFrom-Json
 $script:written.Add((Expand-RpaDateTokens $t.Sheet $script:ScenarioDateBase)+'!'+(Resolve-ExcelCellAddress $t.Cell $script:ScenarioDateBase))
}
$rows=@(
 @('日付条件If','日=1',''),
 @('Excelセル入力','月初','{"Kind":"Excel","Sheet":"{年月-1}","Cell":"A{日-1}"}'),
 @('Else','',''),
 @('Excelセル入力','通常','{"Kind":"Excel","Sheet":"{年月}","Cell":"A{日}"}'),
 @('日付条件If','月=12',''),
 @('待機','0',''),
 @('EndIf','',''),
 @('EndIf','','')
)
foreach($row in $rows){Add-StoredStep ([pscustomobject]@{Action=$row[0];Value=$row[1];Target=$row[2];Delay='0';X='0';Y='0'})}
Check-Steps @(Get-Steps)
$script:testDay=[datetime]::new(2026,12,1);Invoke-Scenario
if(($script:written -join ',') -cne '2026年11月!A30'){throw 'First-day branch failed'}
if($grid.Rows[3].Cells['State'].Value -notmatch 'スキップ'){throw 'Skip state missing'}
$script:written.Clear();$script:testDay=[datetime]::new(2026,12,6);Invoke-Scenario
if(($script:written -join ',') -cne '2026年12月!A6'){throw 'Else branch failed'}
if(-not $form.Visible -or $executionLabel.Text -notmatch '完了'){throw 'Execution failed'}
Write-Output 'PASS: month/year boundaries, leap-year month ends, day offsets, If/Else/nesting and skip display'
$form.Hide()
'@
$path=Join-Path $PSScriptRoot '_date-condition-test.ps1'
try{
 $source=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'RpaStudio.ps1'))
 [IO.File]::WriteAllText($path,$source.Replace('[void]$form.ShowDialog()',$test),[Text.UTF8Encoding]::new($true))
 & powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File $path
 if($LASTEXITCODE -ne 0){throw 'Date condition verification failed'}
}finally{Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue}
