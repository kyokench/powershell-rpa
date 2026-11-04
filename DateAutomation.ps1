# Date text generation is independent of the PC's locale and time zone.
function Get-RpaToday {
 [TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([datetime]::UtcNow,'Tokyo Standard Time').Date
}
function Expand-RpaDateTokens([string]$text, [datetime]$day) {
 $culture=[Globalization.CultureInfo]::InvariantCulture
 $text.Replace('{年月}',$day.ToString("yyyy'年'M'月'",$culture)).Replace('{日付}',$day.ToString('yyyy.MM.dd',$culture))
}
function Get-DateInputText($step, [datetime]$Today=(Get-RpaToday)) {
 if([string]::IsNullOrWhiteSpace($step.Value) -or $step.Value.Length -gt 100){throw '日付の書式を指定してください（100文字以内）。'}
 $options=if($step.DateOptions){$step.DateOptions|ConvertFrom-Json -ErrorAction Stop}else{[pscustomobject]@{Basis='Today';OffsetDays=0}}
 $offset=0
 if(-not [int]::TryParse([string]$options.OffsetDays,[ref]$offset) -or $offset -lt -36500 -or $offset -gt 36500){throw '日付の加減算は-36500～36500日の整数で指定してください。'}
 switch($options.Basis){
  'Today' {$date=$Today.Date}
  'Fixed' {
   $date=[datetime]::MinValue
   if(-not [datetime]::TryParseExact([string]$options.FixedDate,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$date)){throw '指定日が正しくありません。日付入力設定で設定し直してください。'}
  }
  default {throw '日付の基準は実行日または指定日を選択してください。'}
 }
 $format=[string]$step.Value
 # M and d alone are .NET standard formats; force their custom-format meaning.
 if($format.Length -eq 1 -and $format -cin @('y','M','d')){$format='%'+$format}
 try{$date.AddDays($offset).ToString($format,[Globalization.CultureInfo]::GetCultureInfo('ja-JP'))}
 catch{throw '日付の書式または加減算後の日付が正しくありません。日付入力設定のプレビューを確認してください。'}
}
function Invoke-DateInput($step) {
 $text=if($script:ScenarioDateBase){Get-DateInputText $step -Today $script:ScenarioDateBase}else{Get-DateInputText $step}
 if($step.Target){
  $target=$step.Target|ConvertFrom-Json -ErrorAction Stop
  switch($target.Kind){
   'Web' {Invoke-WebStep $step -InputMode -inputValue $text}
   'Excel' {Invoke-ExcelStep ([pscustomobject]@{Action='Excelセル入力';Value=$text;Target=$step.Target})}
   default {Invoke-ElementStep ([pscustomobject]@{Action='要素入力';Value=$text;Target=$step.Target})}
  }
 }else{
  [Windows.Forms.Clipboard]::SetText($text)
  [Windows.Forms.SendKeys]::SendWait('^v')
 }
}
