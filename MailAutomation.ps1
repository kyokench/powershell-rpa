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
function Invoke-MailDraft($step) {
 Wait-Rpa 0
 $settings=Read-MailTarget $step.Target
 $day=if($script:ScenarioDateBase){$script:ScenarioDateBase}else{Get-RpaToday}
 $files=@(foreach($file in $settings.Attachments){
  $path=Expand-RpaDateTokens $file $day
  if(-not (Test-Path -LiteralPath $path -PathType Leaf)){throw "添付ファイルが見つかりません: $path"}
  $path
 })
 $app=$null;$item=$null;$attachments=$null;$saved=$false
 try {
  try{$app=New-Object -ComObject Outlook.Application -ErrorAction Stop}
  catch{throw 'クラシック Outlookを起動できません。インストールとメールプロファイルの設定を確認してください。'}
  $item=$app.CreateItem(0)
  $item.To=[string]$settings.To;$item.Subject=[string]$settings.Subject
  $item.BodyFormat=1;$item.Body=[string]$step.Value
  $item.Save();$saved=$true
  if($files.Count){
   $attachments=$item.Attachments
   foreach($path in $files){$attachment=$attachments.Add($path,1);Release-MailObject $attachment}
   $item.Save()
  }
  Write-Log "Outlookの下書きに保存しました（添付 $($files.Count) 件）。"
 } catch {
  if($saved){throw "添付付き下書きを完成できませんでした。Outlookに途中の下書きが残っている可能性があります: $($_.Exception.Message)"}
  throw
 } finally {Release-MailObject $attachments;Release-MailObject $item;Release-MailObject $app}
}
