#requires -Version 5.1
$ErrorActionPreference='Stop'
$testPath=Join-Path $PSScriptRoot '_excel-verify.ps1'
$workbookPath=Join-Path $PSScriptRoot ('_excel-test-'+[guid]::NewGuid().ToString('N')+'.xlsx')
$test=@'
$ErrorActionPreference='Stop'
$script:ExcelVisible=$false
$fixture=$null;$books=$null;$book=$null
try{
 $fixture=New-Object -ComObject Excel.Application
 $fixture.Visible=$false;$fixture.AutomationSecurity=3
 $books=$fixture.Workbooks;$book=$books.Add()
 $sheets=$book.Worksheets;$sheet=$sheets.Item(1)
 try{$sheet.Name='RpaTest'}finally{Release-ExcelObject $sheet;Release-ExcelObject $sheets}
 $book.SaveAs($workbookPath,51);$book.Close($false)
}finally{
 Release-ExcelObject $book;Release-ExcelObject $books
 if($fixture){$fixture.Quit();Release-ExcelObject $fixture}
}
$form.Show()
$steps=@(
 [pscustomobject]@{Action='Excelブックを開く';Value=$workbookPath;X='0';Y='0';Delay='0';Target=''},
 [pscustomobject]@{Action='Excelシート選択';Value='RpaTest';X='0';Y='0';Delay='0';Target=''},
 [pscustomobject]@{Action='Excelセル入力';Value='00123日本語';X='0';Y='0';Delay='0';Target='{"Kind":"Excel","Sheet":"","Cell":"A1"}'},
 [pscustomobject]@{Action='Excelブック保存';Value='';X='0';Y='0';Delay='0';Target=''},
 [pscustomobject]@{Action='Excelブックを閉じる';Value='';X='0';Y='0';Delay='0';Target=''}
)
foreach($step in $steps){Add-StoredStep $step}
Check-Steps @(Get-Steps)
Invoke-Scenario
if($executionLabel.Text -notmatch '^完了'){throw 'Excel scenario failed'}
if($script:ExcelApp -or $script:ExcelBook){throw 'Excel instance not released'}
Invoke-ExcelStep $steps[0]
$sheets=$script:ExcelBook.Worksheets;$sheet=$sheets.Item('RpaTest');$range=$sheet.Range('A1')
try{
 if($range.Value2 -cne '00123日本語'){throw 'Saved content mismatch'}
}finally{Release-ExcelObject $range;Release-ExcelObject $sheet;Release-ExcelObject $sheets}
$steps[2].Value='=1+1';Invoke-ExcelStep $steps[2]
$refused=$false
try{Invoke-ExcelStep $steps[4]}catch{if($_.Exception.Message -like '*未保存*'){$refused=$true}else{throw}}
if(-not $refused){throw 'Unsaved changes were discarded'}
Invoke-ExcelStep $steps[3]
$sheets=$script:ExcelBook.Worksheets;$sheet=$sheets.Item('RpaTest');$range=$sheet.Range('A1')
try{if($range.HasFormula -or $range.Value2 -ne '=1+1'){throw 'Text input became formula'}}finally{Release-ExcelObject $range;Release-ExcelObject $sheet;Release-ExcelObject $sheets}
Invoke-ExcelStep $steps[4]
$export=Join-Path $PSScriptRoot '_excel-test-export.vbs'
try{
 Export-ScenarioFile $export @(Get-Steps)
 $vbs=[IO.File]::ReadAllText($export)
 $encoded=([regex]::Matches($vbs,'data = data & "([A-Za-z0-9+/=]+)"')|ForEach-Object{$_.Groups[1].Value}) -join ''
 $source=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($encoded))
 $errors=$null;[void][Management.Automation.Language.Parser]::ParseInput($source,[ref]$null,[ref]$errors)
 if($errors.Count){throw ($errors|Out-String)}
 if($source -match '(?m)^\. \(Join-Path \$PSScriptRoot'){throw 'Module not embedded'}
}finally{Remove-Item -LiteralPath $export -Force -ErrorAction SilentlyContinue}
$form.Hide()
Write-Output 'PASS: Excel scenario, sheet selection, text and leading zeros, save/reopen, unsaved-close guard, export'
'@
$encodedPath=[Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($workbookPath))
$setup='$workbookPath=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('''+$encodedPath+'''))' + "`r`n"
$source=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'RpaStudio.ps1'))
$body=$setup+"try {`r`n"+$test+@'
} finally {
 if($script:ExcelBook){try{$script:ExcelBook.Close($false)}catch{};Release-ExcelObject $script:ExcelBook;$script:ExcelBook=$null}
 if($script:ExcelApp){try{$script:ExcelApp.Quit()}catch{};Release-ExcelObject $script:ExcelApp;$script:ExcelApp=$null}
}
'@
try{
 [IO.File]::WriteAllText($testPath,$source.Replace('[void]$form.ShowDialog()',$body),[Text.UTF8Encoding]::new($true))
 & powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File $testPath
 if($LASTEXITCODE -ne 0){throw 'Excel verification failed'}
}finally{Remove-Item -LiteralPath $testPath,$workbookPath -Force -ErrorAction SilentlyContinue}
