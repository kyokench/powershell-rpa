function Read-DateCondition([string]$expression) {
 $match=[regex]::Match($expression,'^\s*(日|月|年)\s*(>=|<=|!=|=|>|<)\s*([0-9]{1,4})\s*$')
 if(-not $match.Success){throw '日付条件は「日=1」「日!=1」「月=12」などで指定してください。'}
 $part=$match.Groups[1].Value;$value=[int]$match.Groups[3].Value
 $maximum=switch($part){'日'{31};'月'{12};'年'{9999}}
 if($value -lt 1 -or $value -gt $maximum){throw '日付条件の数値が範囲外です。'}
 [pscustomobject]@{Part=$part;Operator=$match.Groups[2].Value;Value=$value}
}
function Test-DateCondition([string]$expression,[datetime]$day) {
 $condition=Read-DateCondition $expression
 $actual=switch($condition.Part){'日'{$day.Day};'月'{$day.Month};'年'{$day.Year}}
 switch($condition.Operator){
  '=' {return $actual -eq $condition.Value}
  '!=' {return $actual -ne $condition.Value}
  '>' {return $actual -gt $condition.Value}
  '<' {return $actual -lt $condition.Value}
  '>=' {return $actual -ge $condition.Value}
  '<=' {return $actual -le $condition.Value}
 }
}
function Assert-ScenarioFlow($steps) {
 $stack=New-Object 'System.Collections.Generic.List[object]'
 foreach($step in $steps){
  switch($step.Action){
   '日付条件If' {$null=Read-DateCondition $step.Value;$stack.Add([pscustomobject]@{ElseSeen=$false})}
   'Else' {
    if($stack.Count -eq 0 -or $stack[$stack.Count-1].ElseSeen){throw 'Elseの位置が正しくありません。対応するIfを含めてください。'}
    $stack[$stack.Count-1].ElseSeen=$true
   }
   'EndIf' {
    if($stack.Count -eq 0){throw 'EndIfに対応するIfがありません。部分実行ではIf～EndIf全体を選択してください。'}
    $stack.RemoveAt($stack.Count-1)
   }
  }
 }
 if($stack.Count){throw 'Ifに対応するEndIfがありません。部分実行ではIf～EndIf全体を選択してください。'}
}
function Invoke-DateFlow($step,$stack,[bool]$active,[datetime]$day) {
 switch($step.Action){
  '日付条件If' {
   $result=Test-DateCondition $step.Value $day
   $stack.Add([pscustomobject]@{ParentActive=$active;ConditionResult=$result;ElseSeen=$false})
   Write-Log "日付条件 $($step.Value): $result"
   return ($active -and $result)
  }
  'Else' {
   $frame=$stack[$stack.Count-1];$frame.ElseSeen=$true
   return ($frame.ParentActive -and -not $frame.ConditionResult)
  }
  'EndIf' {
   $frame=$stack[$stack.Count-1];$stack.RemoveAt($stack.Count-1)
   return $frame.ParentActive
  }
 }
}
