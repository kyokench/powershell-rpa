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
    if(-not $script:ExcelBook.Saved){throw '現在のExcelブックに未保存の変更があります。先に「Excelブック保存」を実行してください。'}
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
    $range=$sheet.Range([string]$target.Cell)
    if($range.CountLarge -ne 1){throw '入力先は1セルを指定してください。'}
    $range.NumberFormat='@'
    $range.Value2=[string]$step.Value
   }finally{Release-ExcelObject $range;Release-ExcelObject $sheet;Release-ExcelObject $sheets}
  }
  {$_ -in @('Excelセル選択','Excelシート印刷','ExcelシートPDF保存')} {
   Assert-ExcelBook
   $target=$step.Target|ConvertFrom-Json -ErrorAction Stop
   $sheetName=if($target.Sheet){Expand-RpaDateTokens $target.Sheet $day}else{$script:ExcelSheetName}
   $sheets=$null;$sheet=$null;$range=$null
   try{
    $sheets=$script:ExcelBook.Worksheets;$sheet=$sheets.Item($sheetName)
    switch($step.Action){
     'Excelセル選択' {
      $range=$sheet.Range([string]$target.Cell)
      $sheet.Activate();$range.Select()
      $script:ExcelSheetName=[string]$sheet.Name
     }
     'Excelシート印刷' { $sheet.PrintOut() }
     'ExcelシートPDF保存' {
      $folder=Resolve-ExcelPdfFolder $step.Value $day
      $destination=Join-Path $folder ($day.ToString('yyyy.MM.dd',[Globalization.CultureInfo]::InvariantCulture)+'.pdf')
      if(Test-Path -LiteralPath $destination){throw "同名のPDFが既にあります: $destination"}
      $sheet.ExportAsFixedFormat(0,$destination,0,$true,$false)
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
   if(-not $script:ExcelBook.Saved){throw '未保存の変更があります。先に「Excelブック保存」を実行してください。'}
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
