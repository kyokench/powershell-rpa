#requires -Version 5.1
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Windows.Forms
. (Join-Path $PSScriptRoot 'WebAutomation.ps1')
$script:cancel=$false
function Wait-Rpa([int]$ms){
 if($script:cancel){throw [OperationCanceledException]::new('Test cancellation')}
 if($ms -gt 0){Start-Sleep -Milliseconds ([Math]::Min($ms,20))}
}
Add-Type @'
using System;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
public sealed class WebFixtureServer : IDisposable {
 readonly TcpListener listener; readonly Thread worker; volatile bool running=true; readonly byte[] body;
 public int Port { get; private set; }
 public WebFixtureServer(string html) {
  body=Encoding.UTF8.GetBytes(html);listener=new TcpListener(IPAddress.Loopback,0);listener.Start();Port=((IPEndPoint)listener.LocalEndpoint).Port;
  worker=new Thread(Serve);worker.IsBackground=true;worker.Start();
 }
 void Serve(){while(running){try{using(var client=listener.AcceptTcpClient()){
  client.ReceiveTimeout=2000;var stream=client.GetStream();var data=new byte[8192];stream.Read(data,0,data.Length);
  var header=Encoding.ASCII.GetBytes("HTTP/1.1 200 OK\r\nContent-Type: text/html; charset=utf-8\r\nContent-Length: "+body.Length+"\r\nConnection: close\r\n\r\n");
  stream.Write(header,0,header.Length);stream.Write(body,0,body.Length);
 }}catch{if(!running)return;}}}
 public void Dispose(){running=false;listener.Stop();worker.Join(3000);}
}
'@
$html=@'
<!doctype html><html><head><title>Web RPA fixture</title></head><body>
<button id="btn:special" class="action primary" onclick="document.body.dataset.clicked=String(Number(document.body.dataset.clicked||0)+1);document.body.dataset.trusted=String(event.isTrusted)">Click me</button>
<button class="classOnly unique">Class only</button><button class="duplicate">One</button><button class="duplicate">Two</button>
<input id="input" name="login"><input id="password" type="password"><button id="disabled" disabled>Disabled</button>
</body></html>
'@
$server=New-Object WebFixtureServer $html
$temporaryRoot=[IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$profile=[IO.Path]::GetFullPath((Join-Path $temporaryRoot ('RpaWebTest-'+[guid]::NewGuid().ToString('N'))))
$script:WebProfilePath=$profile
$edge=Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'
$url="http://127.0.0.1:$($server.Port)/fixture"
try{
 $process=Start-Process $edge -ArgumentList ('--headless=new --disable-gpu --no-first-run --no-default-browser-check --remote-debugging-address=127.0.0.1 --remote-debugging-port=0 --user-data-dir="'+$profile+'" '+$url) -WindowStyle Hidden -PassThru
 $clock=[Diagnostics.Stopwatch]::StartNew()
 do{Start-Sleep -Milliseconds 100;$base=Get-WebEndpoint}while(-not $base -and $clock.ElapsedMilliseconds -lt 15000)
 if(-not $base){throw 'Headless Edge did not start'}
 $page=@(Get-WebPages|Where-Object{$_.url -eq $url})[0]
 if(-not $page){throw 'Fixture page missing'}
 for($i=0;$i -lt 30;$i++){
  $ready=Invoke-WebExpression $page '!!document.getElementById("btn:special")'
  if($ready){break};Start-Sleep -Milliseconds 100
 }
 if(-not $ready){throw 'Fixture not loaded'}
 [void](Invoke-WebExpression $page (Get-WebPickerScript '__rpaTest'))
 $rect=Invoke-WebExpression $page '(()=>{const r=document.getElementById("btn:special").getBoundingClientRect();return {x:r.left+r.width/2,y:r.top+r.height/2}})()'
 [void](Invoke-WebProtocol $page 'Input.dispatchMouseEvent' @{type='mousePressed';x=$rect.x;y=$rect.y;button='left';clickCount=1})
 [void](Invoke-WebProtocol $page 'Input.dispatchMouseEvent' @{type='mouseReleased';x=$rect.x;y=$rect.y;button='left';clickCount=1})
 $target=Invoke-WebExpression $page 'window.__rpaTest.result'
 if($target.Id -ne 'btn:special' -or $target.Tag -ne 'button' -or $target.Classes -notcontains 'primary' -or -not $target.Selector){throw 'Picker metadata missing'}
 if((Invoke-WebExpression $page 'document.body.dataset.clicked||"0"') -ne '0'){throw 'Picker triggered a normal click'}
 $step=[pscustomobject]@{Target=($target|ConvertTo-Json -Depth 8 -Compress)}
 Invoke-WebStep $step
 if((Invoke-WebExpression $page 'document.body.dataset.clicked') -ne '1' -or (Invoke-WebExpression $page 'document.body.dataset.trusted') -ne 'true'){throw 'Trusted click failed'}
 Write-Output 'PASS: CDP connection, ID/class/tag capture, picker suppresses click, trusted selector click'
 [void](Invoke-WebExpression $page (Get-WebPickerScript '__rpaTest'))
 [void](Invoke-WebExpression $page 'document.querySelector(".classOnly").click()')
 $classTarget=Invoke-WebExpression $page 'window.__rpaTest.result'
 if($classTarget.Selector -notmatch 'classOnly'){throw 'Class selector missing'}
 $target.Selector='#input';$step.Target=$target|ConvertTo-Json -Compress -Depth 8
 $fake='TestOnly+^{}日本語"__SELECTOR____VALUE__'
 Invoke-WebStep $step -InputMode -inputValue $fake
 if((Invoke-WebExpression $page 'document.getElementById("input").value') -cne $fake){throw 'DOM input failed'}
 $target.Selector='#password';$step.Target=$target|ConvertTo-Json -Compress -Depth 8
 Invoke-WebStep $step -InputMode -inputValue $fake
 if((Invoke-WebExpression $page 'document.getElementById("password").value') -cne $fake){throw 'Password DOM input failed'}
 $target.Selector='.duplicate';$step.Target=$target|ConvertTo-Json -Compress -Depth 8
 $refused=$false
 try{Invoke-WebStep $step}catch{if($_.Exception.Message -like '*複数*'){$refused=$true}else{throw}}
 if(-not $refused){throw 'Duplicate elements accepted'}
 $script:cancel=$true;$cancelled=$false
 try{Invoke-WebStep $step}catch{if($_.Exception -is [OperationCanceledException]){$cancelled=$true}else{throw}}
 $script:cancel=$false
 if(-not $cancelled){throw 'Cancellation ignored'}
 Write-Output 'PASS: class selector, Unicode/special-character input, password input, duplicate rejection, cancellation'
 $integration=@'
$script:WebProfilePath=$profile
$form.Show()
Add-StoredStep ([pscustomobject]@{Action='URLを開く';Value=$url;X='0';Y='0';Delay='0';Target='';ProtectedValue=''})
$target.Selector='#input'
$webInput=[pscustomobject]@{Action='ウェブ要素入力';Value='Scenario test';X='0';Y='0';Delay='0';Target=($target|ConvertTo-Json -Compress -Depth 8);ProtectedValue=''}
Add-StoredStep $webInput
$target.Selector='#password'
$protected=[pscustomobject]@{Action='パスワード入力';Value='[暗号化済み]';X='0';Y='0';Delay='0';Target=($target|ConvertTo-Json -Compress -Depth 8);ProtectedValue=(Protect-InputValue $fake)}
Add-StoredStep $protected
Check-Steps @(Get-Steps) -DecryptSecrets
Invoke-Scenario
if($executionLabel.Text -notmatch '^完了'){throw 'Web scenario execution failed'}
if($log.Text -notmatch 'このURLは専用Edge'){throw 'Legacy URL was not routed to dedicated Edge'}
if((Invoke-WebExpression $page 'document.getElementById("password").value') -cne $fake){throw 'Protected Web password input failed'}
if($log.Text.Contains($fake)){throw 'Secret logged'}
$output=Join-Path $PSScriptRoot '_web-test-export.vbs'
try{
 Export-ScenarioFile $output @(Get-Steps)
 $vbs=[IO.File]::ReadAllText($output)
 $encoded=([regex]::Matches($vbs,'data = data & "([A-Za-z0-9+/=]+)"')|ForEach-Object{$_.Groups[1].Value}) -join ''
 $exported=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($encoded))
 $errors=$null;[void][Management.Automation.Language.Parser]::ParseInput($exported,[ref]$null,[ref]$errors)
 if($errors.Count){throw ($errors|Out-String)}
 if($exported -match '(?m)^\. \(Join-Path \$PSScriptRoot'){throw 'Module not embedded'}
 if($exported.Contains($fake)){throw 'Secret exported in plaintext'}
}finally{Remove-Item -LiteralPath $output -Force -ErrorAction SilentlyContinue}
$form.Hide()
Write-Output 'PASS: legacy URL uses dedicated Edge, GUI execution, protected Web password, logs, export syntax'
'@
 $source=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'RpaStudio.ps1'))
 $context=@{Profile=$profile;Fake=$fake;Target=$target;Url=$url}|ConvertTo-Json -Compress -Depth 8
 $payload=[Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($context))
 $setup='$context=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('''+$payload+'''))|ConvertFrom-Json; $profile=$context.Profile; $fake=$context.Fake; $target=$context.Target; $url=$context.Url; $script:WebProfilePath=$profile; $page=@(Get-WebPages|Where-Object{$_.url -eq $target.PageUrl})[0]' + "`r`n"
 $integrationPath=Join-Path $PSScriptRoot '_web-integration.ps1'
 try{
  [IO.File]::WriteAllText($integrationPath,$source.Replace('[void]$form.ShowDialog()',($setup+$integration)),[Text.UTF8Encoding]::new($true))
  & powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File $integrationPath
  if($LASTEXITCODE -ne 0){throw 'GUI integration test failed'}
 }finally{Remove-Item -LiteralPath $integrationPath -Force -ErrorAction SilentlyContinue}
 Start-RpaWebBrowser ("http://127.0.0.1:$($server.Port)/second")
 $navigated=@(Get-WebPages|Where-Object{$_.id -eq $script:WebActivePageId})[0]
 if($navigated.url -notmatch '/second$'){throw 'Dedicated navigation failed'}
 $null=Invoke-RestMethod -Uri ((Get-WebEndpoint)+'/json/new?'+[uri]::EscapeDataString($url+'?alternate=1')) -Method Put -TimeoutSec 3
 $refused=$false
 try{$null=Find-WebTargetPage $target}catch{if($_.Exception.Message -like '*同じURL*'){$refused=$true}else{throw}}
 if(-not $refused){throw 'Duplicate page URLs accepted'}
 Write-Output 'PASS: dedicated browser URL navigation and duplicate page rejection'
}finally{
 $script:cancel=$false
 if($page){try{[void](Invoke-WebProtocol $page 'Browser.close' @{})}catch{}}
 if($process -and -not $process.HasExited){[void]$process.WaitForExit(5000);if(-not $process.HasExited){Stop-Process -Id $process.Id}}
 $server.Dispose()
 if($profile.StartsWith($temporaryRoot,[StringComparison]::OrdinalIgnoreCase) -and (Split-Path $profile -Leaf).StartsWith('RpaWebTest-')){
  Remove-Item -LiteralPath $profile -Recurse -Force -ErrorAction SilentlyContinue
 }
}
