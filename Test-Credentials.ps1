#requires -Version 5.1
$ErrorActionPreference='Stop'
$encoding=[Text.UTF8Encoding]::new($true)
$fixturePath=Join-Path $PSScriptRoot '_credential-fixture.ps1'
$testPath=Join-Path $PSScriptRoot '_credential-verify.ps1'
$exportPath=Join-Path $PSScriptRoot '_credential-export.vbs'
$fixture=@'
$ErrorActionPreference='Stop'
Add-Type -AssemblyName PresentationFramework
[xml]$xml='<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="RpaCredentialTest" Width="420" Height="240"><StackPanel><TextBox AutomationProperties.AutomationId="CredentialTestId" AutomationProperties.Name="CredentialTestId"/><PasswordBox Name="Secret" AutomationProperties.AutomationId="CredentialTestPassword" AutomationProperties.Name="CredentialTestPassword"/><Button Name="Check" AutomationProperties.AutomationId="CredentialTestCheck" Content="Check"/><TextBlock Name="Result" AutomationProperties.AutomationId="CredentialTestResult" Text="Pending"/></StackPanel></Window>'
$w=[Windows.Markup.XamlReader]::Load([Xml.XmlNodeReader]::new($xml))
$w.FindName('Check').Add_Click({
 if($w.FindName('Secret').Password -ceq 'TestOnly+^{}日本語123!'){$w.FindName('Result').Text='OK'}else{$w.FindName('Result').Text='FAIL'}
})
[void]$w.ShowDialog()
'@
$test=@'
$ErrorActionPreference='Stop'
$fake='TestOnly+^{}日本語123!'
$cipher=Protect-InputValue $fake
if($cipher -eq $fake -or (Unprotect-InputValue $cipher) -cne $fake){throw 'DPAPI round trip failed'}
if((Protect-InputValue $fake) -eq $cipher){throw 'Ciphertext should be randomized'}
$refused=$false
try{$null=Unprotect-InputValue 'DPAPI:v1:invalid'}catch{$refused=$true}
if(-not $refused){throw 'Invalid ciphertext accepted'}
$inputStep=[pscustomobject]@{Action='ID入力';Value=$fake;X='0';Y='0';Delay='0';Target='';ProtectedValue=$cipher}
Add-StoredStep $inputStep
[void]$grid.Rows.Add('待機','0','0','0','0','')
$grid.ClearSelection();$grid.Rows[0].Selected=$true
Move-Step 1
if($grid.Rows[1].Cells['ProtectedValue'].Value -ne $cipher){throw 'Cipher lost during reorder'}
$steps=@(Get-Steps);Check-Steps $steps -DecryptSecrets
$json=@{Version=1;Steps=$steps}|ConvertTo-Json -Depth 8
if($json.Contains($fake)){throw 'Plaintext saved'}
$doc=$json|ConvertFrom-Json
$grid.Rows.Clear();foreach($step in $doc.Steps){Add-StoredStep $step}
if((Unprotect-InputValue $grid.Rows[1].Cells['ProtectedValue'].Value) -cne $fake){throw 'Reload failed'}
Export-ScenarioFile (Join-Path $PSScriptRoot '_credential-export.vbs') @(Get-Steps)
$vbs=[IO.File]::ReadAllText((Join-Path $PSScriptRoot '_credential-export.vbs'))
$encoded=([regex]::Matches($vbs,'data = data & "([A-Za-z0-9+/=]+)"')|ForEach-Object{$_.Groups[1].Value}) -join ''
$exported=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($encoded))
$errors=$null
[void][Management.Automation.Language.Parser]::ParseInput($exported,[ref]$null,[ref]$errors)
if($errors.Count){throw ($errors|Out-String)}
if($exported.Contains($fake)){throw 'Plaintext exported'}
$autoLine=@($exported -split "`n"|Where-Object{$_.StartsWith('$script:AutoRunScenario = ')})[0]
$embedded=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($autoLine.Split([char]39)[1]))
if($embedded.Contains($fake)){throw 'Plaintext embedded in export'}
Write-Output 'PASS: DPAPI, invalid ciphertext rejection, reorder, JSON reload, encrypted export'
$window=$null
for($i=0;$i -lt 50;$i++){
 $window=[Windows.Automation.AutomationElement]::RootElement.FindFirst([Windows.Automation.TreeScope]::Children,[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::NameProperty,'RpaCredentialTest'))
 if($window){break};Start-Sleep -Milliseconds 100
}
if(-not $window){throw 'Fixture missing'}
function Find-CredentialTestElement($id){$window.FindFirst([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::AutomationIdProperty,$id))}
$idElement=Find-CredentialTestElement 'CredentialTestId'
$inputStep.Target=Get-ElementTarget $idElement|ConvertTo-Json -Compress
Invoke-ProtectedInput $inputStep
if($idElement.GetCurrentPattern([Windows.Automation.ValuePattern]::Pattern).Current.Value -cne $fake){throw 'ID entry failed'}
$inputStep.Action='パスワード入力'
$inputStep.Target=Get-ElementTarget (Find-CredentialTestElement 'CredentialTestPassword')|ConvertTo-Json -Compress
Invoke-ProtectedInput $inputStep
Start-Sleep -Milliseconds 200
(Find-CredentialTestElement 'CredentialTestCheck').GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()
Start-Sleep -Milliseconds 200
if((Find-CredentialTestElement 'CredentialTestResult').Current.Name -ne 'OK'){throw 'Password entry failed'}
if($log.Text.Contains($fake)){throw 'Plaintext logged'}
Write-Output 'PASS: ID element input, masked password input with Unicode and special characters, no plaintext logs'
'@
try{
 [IO.File]::WriteAllText($fixturePath,$fixture,$encoding)
 $source=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'RpaStudio.ps1'))
 [IO.File]::WriteAllText($testPath,$source.Replace('[void]$form.ShowDialog()',$test),$encoding)
 $process=Start-Process powershell.exe -ArgumentList ('-NoProfile -STA -ExecutionPolicy Bypass -File "'+$fixturePath+'"') -WindowStyle Hidden -PassThru
 & powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File $testPath
 if($LASTEXITCODE -ne 0){throw 'Credential verification failed'}
}finally{
 if($process -and -not $process.HasExited){Stop-Process -Id $process.Id}
 Remove-Item -LiteralPath $fixturePath,$testPath,$exportPath -Force -ErrorAction SilentlyContinue
}
