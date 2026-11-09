# Date text generation is independent of the PC's locale and time zone.
function Get-RpaToday {
 [TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([datetime]::UtcNow,'Tokyo Standard Time').Date
}
function Expand-RpaDateTokens([string]$text, [datetime]$day) {
 $culture=[Globalization.CultureInfo]::InvariantCulture
 [regex]::Replace($text,'\{(年月日|日付|年月|年|月|日)([+-][0-9]{1,5})?\}',[Text.RegularExpressions.MatchEvaluator]{
  param($match)
  $part=$match.Groups[1].Value;$offset=0
  if($match.Groups[2].Success){$offset=[int]$match.Groups[2].Value}
  if([Math]::Abs($offset) -gt 36500){throw '日付の加減算が大きすぎます。'}
  $date=switch($part){
   '年月' {$day.AddMonths($offset)}
   '年' {$day.AddYears($offset)}
   '月' {$day.AddMonths($offset)}
   default {$day.AddDays($offset)}
  }
  switch($part){
   '年月' {$date.ToString("yyyy'年'M'月'",$culture)}
   '年月日' {$date.ToString("yyyy'年'M'月'd'日'",$culture)}
   '日付' {$date.ToString('yyyy.MM.dd',$culture)}
   '年' {$date.ToString('yyyy',$culture)}
   '月' {$date.Month.ToString($culture)}
   '日' {$date.Day.ToString($culture)}
  }
 })
}
function Assert-RpaVariableName([string]$name) {
 if($name -notmatch '^[\p{L}_][\p{L}\p{Nd}_]{0,63}\z'){throw '変数名は文字または _ で始め、文字・数字・_ で64文字以内にしてください。'}
}
function Expand-RpaVariables([string]$text,[switch]$ValidateTemplate) {
 [regex]::Replace($text,'\{変数:([^{}]*)\}',[Text.RegularExpressions.MatchEvaluator]{
  param($match)
  $name=$match.Groups[1].Value
  Assert-RpaVariableName $name
  if($ValidateTemplate){return '変数値'}
  if($null -eq $script:RpaVariables -or -not $script:RpaVariables.ContainsKey($name)){throw "変数「$name」が未設定です。先にExcelセル値を変数に保存してください。"}
  return [string]$script:RpaVariables[$name]
 })
}
function Get-DateInputText($step, [datetime]$Today=(Get-RpaToday), [switch]$ValidateTemplate) {
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
  'MonthStart' {$date=[datetime]::new($Today.Year,$Today.Month,1)}
  'MonthEnd' {$date=[datetime]::new($Today.Year,$Today.Month,[datetime]::DaysInMonth($Today.Year,$Today.Month))}
  'MonthDay' {
   $monthDay=0
   if(-not [int]::TryParse([string]$options.MonthDay,[ref]$monthDay) -or $monthDay -notin (1..31)){throw '月内の日は1～31で指定してください。'}
   if($monthDay -gt [datetime]::DaysInMonth($Today.Year,$Today.Month)){
    if(-not $ValidateTemplate){throw "実行月に $monthDay 日はありません。"}
    $monthDay=1
   }
   $date=[datetime]::new($Today.Year,$Today.Month,$monthDay)
  }
  default {throw '日付入力設定で基準日を選択してください。'}
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
