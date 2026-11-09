function Read-MailTarget([string]$json) {
 $target=$json|ConvertFrom-Json -ErrorAction Stop
 if($target.Kind -ne 'Mail'){throw '「メール作成設定」で下書きの内容を指定してください。'}
 foreach($file in $target.Attachments){
  if(-not [IO.Path]::IsPathRooted([string]$file)){throw '添付ファイルには絶対パスを指定してください。'}
 }
 return $target
}
function Release-MailObject($object) {
 if($null -ne $object -and [Runtime.InteropServices.Marshal]::IsComObject($object)){
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($object)
 }
}
function Get-MailActivation($settings) {
 if($null -eq $settings.ActivateAfterCreate){return $true}
 if($settings.ActivateAfterCreate -isnot [bool]){throw 'メールのアクティブ化設定はオンまたはオフを指定してください。'}
 return $settings.ActivateAfterCreate
}
function Invoke-MailDraft($step) {
 Wait-Rpa 0
 $settings=Read-MailTarget $step.Target
 $activate=Get-MailActivation $settings
 $day=if($script:ScenarioDateBase){$script:ScenarioDateBase}else{Get-RpaToday}
 $files=@(foreach($file in $settings.Attachments){
  $path=Expand-RpaVariables (Expand-RpaDateTokens $file $day)
  if(-not (Test-Path -LiteralPath $path -PathType Leaf)){throw "添付ファイルが見つかりません: $path"}
  $path
 })
 $subject=Expand-RpaVariables (Expand-RpaDateTokens ([string]$settings.Subject) $day)
 $body=Expand-RpaVariables ([string]$step.Value)
 $app=$null;$item=$null;$attachments=$null;$saved=$false
 try {
  try{$app=New-Object -ComObject Outlook.Application -ErrorAction Stop}
  catch{throw 'クラシック Outlookを起動できません。インストールとメールプロファイルの設定を確認してください。'}
  $item=$app.CreateItem(0)
  $item.To=[string]$settings.To;$item.Subject=$subject
  $item.BodyFormat=1;$item.Body=$body
  $item.Save();$saved=$true
  if($files.Count){
   $attachments=$item.Attachments
   foreach($path in $files){$attachment=$attachments.Add($path,1);Release-MailObject $attachment}
   $item.Save()
  }
  Write-Log "Outlookの下書きに保存しました（添付 $($files.Count) 件）。"
  if($activate){
   $item.Display($false)
   $inspector=$null
   try{$inspector=$item.GetInspector;$inspector.Activate()}finally{Release-MailObject $inspector}
  }
 } catch {
  if($saved){throw "下書きの添付または表示処理に失敗しました。Outlookに保存された下書きを確認してください: $($_.Exception.Message)"}
  throw
 } finally {Release-MailObject $attachments;Release-MailObject $item;Release-MailObject $app}
}
function Invoke-OutlookActivation {
 Wait-Rpa 0
 $app=$null;$window=$null;$session=$null;$folder=$null
 try{
  $app=New-Object -ComObject Outlook.Application -ErrorAction Stop
  $window=$app.ActiveInspector()
  if(-not $window){$window=$app.ActiveExplorer()}
  if(-not $window){
   $session=$app.GetNamespace('MAPI');$folder=$session.GetDefaultFolder(6)
   $window=$folder.GetExplorer();$window.Display()
  }
  $window.Activate()
  Write-Log 'Outlookのウィンドウをアクティブ化しました。'
 }finally{Release-MailObject $window;Release-MailObject $folder;Release-MailObject $session;Release-MailObject $app}
}
