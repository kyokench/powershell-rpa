# Web DOM operations through a dedicated, local-only Edge DevTools connection.
$script:WebProfilePath=Join-Path $env:LOCALAPPDATA 'PowerShellRpaStudio\WebProfile'
$script:WebActivePageId=$null

function Get-WebEndpoint {
 $portFile=Join-Path $script:WebProfilePath 'DevToolsActivePort'
 if(-not (Test-Path -LiteralPath $portFile)){return $null}
 try {
  $lines=[IO.File]::ReadAllLines($portFile)
  $port=0
  if($lines.Count -lt 2 -or -not [int]::TryParse($lines[0],[ref]$port) -or $port -lt 1 -or $port -gt 65535){return $null}
  $base="http://127.0.0.1:$port"
  $version=Invoke-RestMethod -Uri "$base/json/version" -TimeoutSec 2 -ErrorAction Stop
  $socket=[uri]$version.webSocketDebuggerUrl
  if($socket.Host -notin @('127.0.0.1','localhost') -or $socket.Port -ne $port -or $socket.AbsolutePath -ne $lines[1]){return $null}
  return $base
 } catch {return $null}
}
function Get-WebPages {
 $base=Get-WebEndpoint
 if(-not $base){throw '専用Webブラウザに接続できません。「Webブラウザ起動」を押してください。'}
 $pages=Invoke-RestMethod -Uri "$base/json/list" -TimeoutSec 2 -ErrorAction Stop
 foreach($page in $pages){if($page.type -eq 'page' -and $page.webSocketDebuggerUrl){$page}}
}
function Wait-WebTask($task) {
 while(-not $task.IsCompleted){Wait-Rpa 15}
 $task.GetAwaiter().GetResult()
}
function Invoke-WebProtocol($page,[string]$method,$parameters) {
 Wait-Rpa 0
 $base=Get-WebEndpoint
 if(-not $base){throw '専用Webブラウザへの接続が切れました。'}
 $uri=[uri]$page.webSocketDebuggerUrl
 if($uri.Scheme -ne 'ws' -or $uri.Host -notin @('127.0.0.1','localhost') -or $uri.Port -ne ([uri]$base).Port){throw '専用ブラウザ以外には接続できません。'}
 $socket=New-Object Net.WebSockets.ClientWebSocket
 $timeout=New-Object Threading.CancellationTokenSource
 $timeout.CancelAfter(8000)
 try {
  Wait-WebTask ($socket.ConnectAsync($uri,$timeout.Token))
  $json=@{id=1;method=$method;params=$parameters}|ConvertTo-Json -Depth 20 -Compress
  $bytes=[Text.Encoding]::UTF8.GetBytes($json)
  Wait-WebTask ($socket.SendAsync([ArraySegment[byte]]::new($bytes),[Net.WebSockets.WebSocketMessageType]::Text,$true,$timeout.Token))
  $buffer=New-Object byte[] 65536
  while($true){
   $stream=New-Object IO.MemoryStream
   try {
    do {
     $received=Wait-WebTask ($socket.ReceiveAsync([ArraySegment[byte]]::new($buffer),$timeout.Token))
     if($received.MessageType -eq [Net.WebSockets.WebSocketMessageType]::Close){throw 'ブラウザへの接続が閉じられました。'}
     $stream.Write($buffer,0,$received.Count)
     if($stream.Length -gt 4194304){throw 'ブラウザの応答が大きすぎます。'}
    } while(-not $received.EndOfMessage)
    $reply=[Text.Encoding]::UTF8.GetString($stream.ToArray())|ConvertFrom-Json
   } finally {$stream.Dispose()}
   if($reply.id -eq 1){
    if($reply.error){throw 'ブラウザ操作に失敗しました。対象ページが閉じられていないか確認してください。'}
    return $reply.result
   }
  }
 } finally {$socket.Abort();$socket.Dispose();$timeout.Dispose()}
}
function Invoke-WebExpression($page,[string]$expression) {
 $result=Invoke-WebProtocol $page 'Runtime.evaluate' @{expression=$expression;returnByValue=$true;awaitPromise=$false}
 if($result.exceptionDetails){throw 'ページのHTML操作に失敗しました。ページを再読み込みして取得し直してください。'}
 return $result.result.value
}
function Start-RpaWebBrowser([string]$url='about:blank') {
 if($url -ne 'about:blank'){
  $uri=$null
  if(-not [uri]::TryCreate($url,[UriKind]::Absolute,[ref]$uri) -or $uri.Scheme -notin @('http','https')){throw 'Web用URLはhttpまたはhttpsで指定してください。'}
 }
 $base=Get-WebEndpoint
 if(-not $base){
  $candidates=@((Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),(Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe'),(Join-Path $env:LOCALAPPDATA 'Microsoft\Edge\Application\msedge.exe'))
  $edge=$candidates|Where-Object{Test-Path -LiteralPath $_}|Select-Object -First 1
  if(-not $edge){throw 'Microsoft Edgeが見つかりません。'}
  [void][IO.Directory]::CreateDirectory($script:WebProfilePath)
  $arguments='--remote-debugging-address=127.0.0.1 --remote-debugging-port=0 --no-first-run --no-default-browser-check --user-data-dir="'+$script:WebProfilePath+'" about:blank'
  Start-Process -FilePath $edge -ArgumentList $arguments -ErrorAction Stop
  $clock=[Diagnostics.Stopwatch]::StartNew()
  do {Wait-Rpa 200;$base=Get-WebEndpoint} while(-not $base -and $clock.ElapsedMilliseconds -lt 15000)
  if(-not $base){throw 'Webブラウザを起動できませんでした。Edgeの起動制限を確認してください。'}
 }
 $pages=@(Get-WebPages)
 $page=$pages|Where-Object{$_.id -eq $script:WebActivePageId}|Select-Object -First 1
 if($url -ne 'about:blank'){
  $existing=@($pages|Where-Object{$_.url -ceq $url})
  if($existing.Count -gt 1){throw '同じURLのページが複数あります。対象以外のタブを閉じてください。'}
  if($existing.Count -eq 1){$page=$existing[0]}
 }
 if(-not $page -and $pages.Count -eq 1){$page=$pages[0]}
 if(-not $page){$page=Invoke-RestMethod -Uri ($base+'/json/new?'+[uri]::EscapeDataString('about:blank')) -Method Put -TimeoutSec 3}
 $script:WebActivePageId=$page.id
 if($url -ne 'about:blank'){
  $reply=Invoke-WebProtocol $page 'Page.navigate' @{url=$url}
  if($reply.errorText){throw '指定したURLへ移動できませんでした。'}
 }
 [void](Invoke-WebProtocol $page 'Page.bringToFront' @{})
}
function Select-WebPage {
 $pages=@(Get-WebPages|Where-Object{$_.url -match '^https?://'})
 if($pages.Count -eq 0){throw '専用Edgeで対象のWebサイトを開いてください。'}
 if($pages.Count -eq 1){return $pages[0]}
 $dialog=New-Object Windows.Forms.Form
 $dialog.Text='取得するWebページを選択';$dialog.Size=New-Object Drawing.Size(740,330)
 $dialog.StartPosition='CenterParent'
 $list=New-Object Windows.Forms.ListBox;$list.Dock='Fill'
 foreach($page in $pages){[void]$list.Items.Add(($page.title+'  |  '+$page.url))};$list.SelectedIndex=0
 $ok=New-Object Windows.Forms.Button;$ok.Text='選択';$ok.Dock='Bottom';$ok.Height=35;$ok.DialogResult='OK'
 $dialog.Controls.Add($list);$dialog.Controls.Add($ok);$dialog.AcceptButton=$ok
 try{if($dialog.ShowDialog($form) -eq 'OK'){return $pages[$list.SelectedIndex]};return $null}finally{$dialog.Dispose()}
}
function Get-WebPickerScript([string]$key) {
 $script=@'
(() => {
 const key=__KEY__;
 if(window[key]) window[key].cleanup();
 const state={result:null,cancelled:false};window[key]=state;
 const banner=document.createElement('div');banner.textContent='RPA: 対象をクリックして取得 / Escで中止（このクリックでは実行しません）';
 banner.style.cssText='position:fixed;top:0;left:0;right:0;z-index:2147483647;background:#eef4ff;color:#132747;padding:10px;font:14px sans-serif;pointer-events:none;border-bottom:1px solid #aac';
 document.documentElement.appendChild(banner);
 let marked=null,oldOutline='';
 const unmark=()=>{if(marked){marked.style.outline=oldOutline;marked=null;}};
 const over=e=>{unmark();if(e.target instanceof Element){marked=e.target;oldOutline=marked.style.outline;marked.style.outline='2px solid #3478c8';}};
 const unique=s=>{try{return document.querySelectorAll(s).length===1}catch{return false}};
 const selectorFor=e=>{
  if(e.id){const s='#'+CSS.escape(e.id);if(unique(s))return s;}
  const tag=e.localName;
  for(const attr of ['data-testid','name','aria-label']){
   const value=e.getAttribute(attr);if(value){const s=tag+'['+attr+'='+JSON.stringify(value)+']';if(unique(s))return s;}
  }
  const classes=Array.from(e.classList).slice(0,6);
  if(classes.length){const s=tag+classes.map(c=>'.'+CSS.escape(c)).join('');if(unique(s))return s;}
  const path=[];let node=e;
  while(node&&node.nodeType===1){
   let part=node.localName;
   if(node.id&&unique('#'+CSS.escape(node.id))){path.unshift('#'+CSS.escape(node.id));break;}
   const parent=node.parentElement;
   if(parent){const siblings=Array.from(parent.children).filter(n=>n.localName===node.localName);if(siblings.length>1)part+=':nth-of-type('+(siblings.indexOf(node)+1)+')';}
   path.unshift(part);if(unique(path.join(' > ')))break;node=parent;
  }
  return path.join(' > ');
 };
 const click=e=>{
  e.preventDefault();e.stopImmediatePropagation();
  const el=e.target;
  if(!(el instanceof Element)||el.getRootNode()!==document)return;
  if(el.localName==='iframe'){banner.textContent='iframe内の取得には対応していません。Escで中止してください。';return;}
  state.result={Kind:'Web',PageUrl:location.origin+location.pathname,Selector:selectorFor(el),Id:el.id||'',Classes:Array.from(el.classList),Tag:el.localName};
  state.cleanup();
 };
 const escape=e=>{if(e.key==='Escape'){state.cancelled=true;state.cleanup();}};
 state.cleanup=()=>{unmark();banner.remove();window.removeEventListener('mouseover',over,true);window.removeEventListener('click',click,true);window.removeEventListener('keydown',escape,true);};
 window.addEventListener('mouseover',over,true);window.addEventListener('click',click,true);window.addEventListener('keydown',escape,true);
 return true;
})()
'@
 return $script.Replace('__KEY__',($key|ConvertTo-Json -Compress))
}
function Capture-WebTarget($page) {
 $key='__psRpa_'+[guid]::NewGuid().ToString('N')
 $keyJson=$key|ConvertTo-Json -Compress
 try {
  [void](Invoke-WebExpression $page (Get-WebPickerScript $key))
  [void](Invoke-WebProtocol $page 'Page.bringToFront' @{})
  $clock=[Diagnostics.Stopwatch]::StartNew()
  do {
   Wait-Rpa 150
   $state=Invoke-WebExpression $page "(() => {const s=window[$keyJson];return s?{result:s.result,cancelled:s.cancelled}:null;})()"
   if($null -eq $state){throw 'ページが切り替わりました。取得をやり直してください。'}
   if($state.cancelled){throw [OperationCanceledException]::new('Web要素の取得を中止しました。')}
   if($state.result){$script:WebActivePageId=$page.id;return $state.result}
  } while($clock.ElapsedMilliseconds -lt 60000)
  throw 'Web要素取得がタイムアウトしました（60秒）。'
 } finally {
  # Cleanup must still run when Esc/Stop has set the cancellation flag.
  $cancelBefore=$script:cancel;$script:cancel=$false
  try{[void](Invoke-WebExpression $page "(() => {const s=window[$keyJson];if(s)s.cleanup();delete window[$keyJson];})()")}catch{}
  finally{$script:cancel=$cancelBefore}
 }
}
function Find-WebTargetPage($target) {
 $matches=@(Get-WebPages|Where-Object{
  try{$url=[uri]$_.url;$url.GetLeftPart([UriPartial]::Path) -ceq $target.PageUrl}catch{$false}
 })
 if($matches.Count -gt 1){throw '同じURLのページが複数あります。対象以外のタブを閉じてください。'}
 if($matches.Count -eq 1){return $matches[0]}
 return $null
}
function Invoke-WebStep($step,[string]$inputValue,[switch]$InputMode) {
 $target=$step.Target|ConvertFrom-Json -ErrorAction Stop
 $clock=[Diagnostics.Stopwatch]::StartNew()
 $selector=$target.Selector|ConvertTo-Json -Compress
 do {
  Wait-Rpa 0
  $page=Find-WebTargetPage $target
  if($page){
   $probe=@'
(() => {
 const nodes=document.querySelectorAll(__SELECTOR__);
 if(nodes.length!==1)return {count:nodes.length};
 const el=nodes[0];el.scrollIntoView({block:'center',inline:'center'});
 const r=el.getBoundingClientRect(),style=getComputedStyle(el);
 const x=r.left+r.width/2,y=r.top+r.height/2,hit=document.elementFromPoint(x,y);
 return {count:1,enabled:!el.disabled&&el.getAttribute('aria-disabled')!=='true',visible:r.width>0&&r.height>0&&style.visibility!=='hidden'&&style.display!=='none',clear:hit===el||el.contains(hit),x,y};
})()
'@
   $info=Invoke-WebExpression $page ($probe.Replace('__SELECTOR__',$selector))
   if($info.count -gt 1){throw 'CSSセレクターに一致する要素が複数あります。取得し直してください。'}
   if($info.count -eq 1 -and $info.enabled -and $info.visible -and $info.clear){break}
  }
  Wait-Rpa 200
 } while($clock.ElapsedMilliseconds -lt 10000)
 if(-not $page -or -not $info -or $info.count -ne 1 -or -not $info.enabled -or -not $info.visible -or -not $info.clear){throw 'Web要素が見つからないか、無効・非表示・他の要素に隠れています（10秒）。対象ページとセレクターを確認してください。'}
 if($InputMode){
  $value=$inputValue|ConvertTo-Json -Compress
  $inputScript=@'
(() => {
 const nodes=document.querySelectorAll(__SELECTOR__);if(nodes.length!==1)return false;
 const el=nodes[0];if(el.disabled||el.readOnly)return false;
 const proto=el instanceof HTMLInputElement?HTMLInputElement.prototype:el instanceof HTMLTextAreaElement?HTMLTextAreaElement.prototype:null;
 if(!proto)return false;
 const setter=Object.getOwnPropertyDescriptor(proto,'value').set;
 setter.call(el,__VALUE__);el.dispatchEvent(new Event('input',{bubbles:true}));el.dispatchEvent(new Event('change',{bubbles:true}));return true;
})()
'@
  $bindings=@{__SELECTOR__=$selector;__VALUE__=$value}
  $expression=[regex]::Replace($inputScript,'__(SELECTOR|VALUE)__',[Text.RegularExpressions.MatchEvaluator]{param($m) $bindings[$m.Value]})
  $ok=Invoke-WebExpression $page $expression
  if(-not $ok){throw 'このWeb要素には入力できません。inputまたはtextareaを取得してください。'}
 } else {
  [void](Invoke-WebProtocol $page 'Input.dispatchMouseEvent' @{type='mouseMoved';x=$info.x;y=$info.y})
  try{[void](Invoke-WebProtocol $page 'Input.dispatchMouseEvent' @{type='mousePressed';x=$info.x;y=$info.y;button='left';clickCount=1})}
  finally{
   # Release even if cancellation arrives after mousePressed.
   $cancelBefore=$script:cancel;$script:cancel=$false
   try{[void](Invoke-WebProtocol $page 'Input.dispatchMouseEvent' @{type='mouseReleased';x=$info.x;y=$info.y;button='left';clickCount=1})}
   finally{$script:cancel=$cancelBefore}
  }
 }
}
