# Desktop Excel COM operations. Only the instance created here is controlled.
$script:ExcelApp=$null
$script:ExcelBook=$null
$script:ExcelSheetName=$null
$script:ExcelVisible=$true
function Release-ExcelObject($object){
 if($null -ne $object -and [Runtime.InteropServices.Marshal]::IsComObject($object)){
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($object)
 }
}
function Assert-ExcelBook {
 if($null -eq $script:ExcelBook){throw '先に「Excelブックを開く」を実行してください。'}
 try{$null=$script:ExcelBook.Name}catch{throw 'Excelブックが閉じられています。「Excelブックを開く」から再実行してください。'}
}
function Save-ExcelBeforeClose {
 if(-not $script:ExcelBook.Saved){
  if($script:ExcelBook.ReadOnly){throw '未保存のブックは読み取り専用のため自動保存できません。'}
  $script:ExcelBook.Save()
  if(-not $script:ExcelBook.Saved){throw 'Excelの自動保存が完了しませんでした。保存先やExcelのダイアログを確認してください。'}
  Write-Log 'Excelブックを自動保存しました。'
 }
}
function Resolve-ExcelPdfName($target,[datetime]$day) {
 $template=if([string]::IsNullOrWhiteSpace([string]$target.FileName)){'{日付}.pdf'}else{[string]$target.FileName}
 $name=Expand-RpaDateTokens $template $day
 if($name.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0 -or $name.EndsWith('.') -or $name.EndsWith(' ')){throw 'PDFファイル名に使用できない文字が含まれています。保存先フォルダーは別の欄で指定してください。'}
 if(-not $name.EndsWith('.pdf',[StringComparison]::OrdinalIgnoreCase)){$name+='.pdf'}
 return $name
}
function Resolve-ExcelPdfFolder([string]$path,[datetime]$day) {
 $dynamic=$path.Contains('{年月}')
 $folder=[IO.Path]::GetFullPath((Expand-RpaDateTokens $path $day))
 if(Test-Path -LiteralPath $folder -PathType Container){return $folder}
 $parent=[IO.Path]::GetDirectoryName($folder)
 if($dynamic -and $parent -and (Test-Path -LiteralPath $parent -PathType Container)){
  [void][IO.Directory]::CreateDirectory($folder)
  return $folder
 }
 throw "PDFの保存先フォルダーが見つかりません: $folder"
}
function Resolve-ExcelCellAddress([string]$address,[datetime]$day) {
 if($address -notmatch '^\$?[A-Za-z]{1,3}\$?(?:[1-9][0-9]{0,6}|\{日\})$'){throw 'セル番地は A1 または A{日} のように指定してください。'}
 $address.Replace('{日}',$day.Day.ToString([Globalization.CultureInfo]::InvariantCulture))
}
function Set-ExcelCellClipboard([string]$value) {
 if($value.Length){[Windows.Forms.Clipboard]::SetText($value)}
 else{[Windows.Forms.Clipboard]::Clear()}
}
function Invoke-ExcelStep($step){
 Wait-Rpa 0
 $day=if($script:ScenarioDateBase){$script:ScenarioDateBase}else{Get-RpaToday}
 switch($step.Action){
  'Excelブックを開く' {
   $path=[IO.Path]::GetFullPath([string]$step.Value)
   if(-not (Test-Path -LiteralPath $path -PathType Leaf)){throw '指定したExcelファイルが見つかりません。'}
   if([IO.Path]::GetExtension($path) -notin @('.xlsx','.xlsm','.xlsb','.xls')){throw 'Excelファイル（xlsx/xlsm/xlsb/xls）を指定してください。'}
   if($script:ExcelBook){
    Assert-ExcelBook
    if($script:ExcelBook.FullName -ieq $path){return}
    Save-ExcelBeforeClose
    $script:ExcelBook.Close($false);Release-ExcelObject $script:ExcelBook;$script:ExcelBook=$null
   }
   $books=$null
   try{
    if(-not $script:ExcelApp){
     $script:ExcelApp=New-Object -ComObject Excel.Application -ErrorAction Stop
     $script:ExcelApp.Visible=$script:ExcelVisible
     $script:ExcelApp.AutomationSecurity=3
    }
    $books=$script:ExcelApp.Workbooks
    $script:ExcelBook=$books.Open($path,0,$false)
    $sheet=$script:ExcelBook.ActiveSheet
    try{$script:ExcelSheetName=[string]$sheet.Name}finally{Release-ExcelObject $sheet}
   }catch{
    if(-not $script:ExcelBook -and $script:ExcelApp){
     try{$script:ExcelApp.Quit()}catch{}
     Release-ExcelObject $script:ExcelApp;$script:ExcelApp=$null
    }
    throw 'Excelファイルを開けませんでした。Excelのインストール、ファイル形式、パスワード保護、使用中の状態を確認してください。'
   }finally{Release-ExcelObject $books}
  }
  'Excelシート選択' {
   Assert-ExcelBook
   $sheets=$null;$sheet=$null
   try{
    $sheets=$script:ExcelBook.Worksheets;$sheet=$sheets.Item((Expand-RpaDateTokens $step.Value $day))
    $sheet.Activate();$script:ExcelSheetName=[string]$sheet.Name
   }finally{Release-ExcelObject $sheet;Release-ExcelObject $sheets}
  }
  'Excelセル入力' {
   Assert-ExcelBook
   if($script:ExcelBook.ReadOnly){throw '読み取り専用のブックには入力できません。'}
   $target=$step.Target|ConvertFrom-Json -ErrorAction Stop
   $sheetName=if($target.Sheet){Expand-RpaDateTokens $target.Sheet $day}else{$script:ExcelSheetName}
   $sheets=$null;$sheet=$null;$range=$null
   try{
    $sheets=$script:ExcelBook.Worksheets;$sheet=$sheets.Item($sheetName)
    $range=$sheet.Range((Resolve-ExcelCellAddress $target.Cell $day))
    if($range.CountLarge -ne 1){throw '入力先は1セルを指定してください。'}
    $range.NumberFormat='@'
    $range.Value2=[string]$step.Value
   }finally{Release-ExcelObject $range;Release-ExcelObject $sheet;Release-ExcelObject $sheets}
  }
  {$_ -in @('Excelセル選択','Excelセルコピー','Excelシート印刷','ExcelシートPDF保存')} {
   Assert-ExcelBook
   $target=$step.Target|ConvertFrom-Json -ErrorAction Stop
   $sheetName=if($target.Sheet){Expand-RpaDateTokens $target.Sheet $day}else{$script:ExcelSheetName}
   $sheets=$null;$sheet=$null;$range=$null
   try{
    $sheets=$script:ExcelBook.Worksheets;$sheet=$sheets.Item($sheetName)
    switch($step.Action){
     'Excelセル選択' {
      $range=$sheet.Range((Resolve-ExcelCellAddress $target.Cell $day))
      $sheet.Activate();$range.Select()
      $script:ExcelSheetName=[string]$sheet.Name
     }
     'Excelセルコピー' {
      $range=$sheet.Range((Resolve-ExcelCellAddress $target.Cell $day))
      $script:ExcelApp.Calculate()
      Set-ExcelCellClipboard ([string]$range.Value2)
      Write-Log "Excelセルをクリップボードにコピーしました: $sheetName!$($range.Address())"
     }
     'Excelシート印刷' { $sheet.PrintOut() }
     'ExcelシートPDF保存' {
      $folder=Resolve-ExcelPdfFolder $step.Value $day
      $destination=Join-Path $folder (Resolve-ExcelPdfName $target $day)
      $alerts=$script:ExcelApp.DisplayAlerts
      try{
       $script:ExcelApp.DisplayAlerts=$false
       $sheet.ExportAsFixedFormat(0,$destination,0,$true,$false)
      }finally{$script:ExcelApp.DisplayAlerts=$alerts}
      Write-Log "PDF保存: $destination"
     }
    }
   }finally{Release-ExcelObject $range;Release-ExcelObject $sheet;Release-ExcelObject $sheets}
  }
  'Excelブック保存' {
   Assert-ExcelBook
   if($script:ExcelBook.ReadOnly){throw '読み取り専用のブックは上書き保存できません。'}
   $script:ExcelBook.Save()
  }
  'Excelブックを閉じる' {
   Assert-ExcelBook
   Save-ExcelBeforeClose
   $script:ExcelBook.Close($false)
   Release-ExcelObject $script:ExcelBook;$script:ExcelBook=$null;$script:ExcelSheetName=$null
   # Do not close additional workbooks the user may have opened in this instance.
   $books=$null
   try{
    $books=$script:ExcelApp.Workbooks
    if($books.Count -eq 0){$script:ExcelApp.Quit();Release-ExcelObject $script:ExcelApp;$script:ExcelApp=$null}
   }finally{Release-ExcelObject $books}
  }
 }
}
