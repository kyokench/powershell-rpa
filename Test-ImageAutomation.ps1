#requires -Version 5.1
$ErrorActionPreference='Stop'
$test=@'
$ErrorActionPreference='Stop'
$pattern=[Drawing.Bitmap]::new(48,32,[Drawing.Imaging.PixelFormat]::Format24bppRgb)
for($y=0;$y -lt 32;$y++){for($x=0;$x -lt 48;$x++){$pattern.SetPixel($x,$y,[Drawing.Color]::FromArgb(($x*17+$y*11)%256,($x*5+$y*23)%256,($x*29+$y*7)%256))}}
$canvas=[Drawing.Bitmap]::new(240,180,[Drawing.Imaging.PixelFormat]::Format24bppRgb)
try{
 $g=[Drawing.Graphics]::FromImage($canvas);$g.Clear([Drawing.Color]::White);$g.DrawImageUnscaled($pattern,75,65);$g.Dispose()
 $points=@(Find-RpaImage $canvas $pattern 0)
 if($points.Count -ne 1 -or $points[0].X -ne 75 -or $points[0].Y -ne 65){throw 'Match location failed'}
 $g=[Drawing.Graphics]::FromImage($canvas);$g.DrawImageUnscaled($pattern,140,110);$g.Dispose()
 if(@(Find-RpaImage $canvas $pattern 0).Count -ne 2){throw 'Duplicate detection failed'}
 Write-Output 'PASS: image match location and duplicate detection'
}finally{$canvas.Dispose()}
Add-Type 'using System; using System.Runtime.InteropServices; public static class ImageTestWindow { [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr w,int n); }'
$form.Show()
$fixture=New-Object Windows.Forms.Form;$fixture.Text='RPA Image Test';$fixture.ClientSize=[Drawing.Size]::new(240,160);$fixture.TopMost=$true
$picture=New-Object Windows.Forms.PictureBox;$picture.Size=$pattern.Size;$picture.Location=[Drawing.Point]::new(70,60);$picture.Image=$pattern
$fixture.Controls.Add($picture)
$script:clicked=$false;$picture.Add_MouseClick({$script:clicked=$true})
$offsetButton=New-Object Windows.Forms.Button;$offsetButton.Text='Offset';$offsetButton.SetBounds(145,60,70,35)
$script:offsetClicked=$false;$offsetButton.Add_Click({$script:offsetClicked=$true});$fixture.Controls.Add($offsetButton)
$fixture.Show();[void][ImageTestWindow]::ShowWindow($fixture.Handle,5);$fixture.Activate();Wait-Rpa 500
$crop=$null;$export=Join-Path $PSScriptRoot '_image-export.vbs'
try{
 $screen=Get-RpaScreenshot
 try{
  $bounds=[RpaImageMatcher]::Bounds()
  $area=[RpaImageMatcher]::WindowBounds($picture.Handle)
  $crop=$screen.Clone([Drawing.Rectangle]::new($area.X-$bounds.X,$area.Y-$bounds.Y,$area.Width,$area.Height),[Drawing.Imaging.PixelFormat]::Format24bppRgb)
 }finally{$screen.Dispose()}
 $stream=New-Object IO.MemoryStream
 try{$crop.Save($stream,[Drawing.Imaging.ImageFormat]::Png);$data=[Convert]::ToBase64String($stream.ToArray())}finally{$stream.Dispose()}
 $step=[pscustomobject]@{Action='画像クリック';Value='';Delay='0';X='0';Y='0';Target=(@{Kind='Image';ImageData=$data;Width=48;Height=32;Tolerance=0}|ConvertTo-Json -Compress)}
 Check-Steps @($step)
 Invoke-ImageStep $step;Wait-Rpa 150
 if(-not $script:clicked){throw ('Actual image click failed: '+$log.Text+'; cursor='+[Windows.Forms.Cursor]::Position+'; client='+$picture.PointToClient([Windows.Forms.Cursor]::Position))}
 $selectionTimer=New-Object Windows.Forms.Timer;$selectionTimer.Interval=150
 $selectionTimer.Add_Tick({
  $overlayHere=@([Windows.Forms.Application]::OpenForms|Where-Object{$_.FormBorderStyle -eq 'None' -and $_.TopMost})[0]
  if(-not $overlayHere){return}
  $selectionTimer.Stop()
  $flags=[Reflection.BindingFlags]'Instance,NonPublic'
  $px=$area.X-$bounds.X;$py=$area.Y-$bounds.Y
  $overlayHere.GetType().GetMethod('OnMouseDown',$flags).Invoke($overlayHere,@([Windows.Forms.MouseEventArgs]::new([Windows.Forms.MouseButtons]::Left,1,$px,$py,0)))
  $overlayHere.GetType().GetMethod('OnMouseMove',$flags).Invoke($overlayHere,@([Windows.Forms.MouseEventArgs]::new([Windows.Forms.MouseButtons]::Left,0,$px+$area.Width,$py+$area.Height,0)))
  $overlayHere.GetType().GetMethod('OnMouseUp',$flags).Invoke($overlayHere,@([Windows.Forms.MouseEventArgs]::new([Windows.Forms.MouseButtons]::Left,1,$px+$area.Width,$py+$area.Height,0)))
 })
 try{
  $selectionTimer.Start();$captured=Capture-ImageTarget
  if(-not $captured -or $captured.Width -ne $area.Width -or $captured.Height -ne $area.Height){throw 'Image region capture failed'}
 }finally{$selectionTimer.Dispose()}
 Write-Output 'PASS: drag-region capture dialog and cropped image metadata'
 Add-StoredStep $step
 $legacy=Read-ImageTarget $step.Target
 try{if($legacy.OffsetX -ne 0 -or $legacy.OffsetY -ne 0){throw 'Legacy image offset changed'}}finally{$legacy.Bitmap.Dispose()}
 $buttonArea=[RpaImageMatcher]::WindowBounds($offsetButton.Handle)
 $dx=[int]([Math]::Floor($buttonArea.X+$buttonArea.Width/2)-[Math]::Floor($area.X+$area.Width/2))
 $dy=[int]([Math]::Floor($buttonArea.Y+$buttonArea.Height/2)-[Math]::Floor($area.Y+$area.Height/2))
 $settingsTimer=New-Object Windows.Forms.Timer;$settingsTimer.Interval=100
 $settingsTimer.Add_Tick({
  $settings=@([Windows.Forms.Application]::OpenForms|Where-Object{$_.Text -eq '画像設定'})[0]
  if(-not $settings){return};$settingsTimer.Stop()
  $settings.Controls['OffsetX'].Value=$dx;$settings.Controls['OffsetY'].Value=$dy
  (@($settings.Controls|Where-Object{$_ -is [Windows.Forms.Button] -and $_.Text -eq '設定'})[0]).PerformClick()
 })
 try{$settingsTimer.Start();$imageSetup.PerformClick()}finally{$settingsTimer.Dispose()}
 $updated=@(Get-Steps)[0]
 $offsetData=$updated.Target|ConvertFrom-Json
 if($offsetData.OffsetX -ne $dx -or $offsetData.OffsetY -ne $dy){throw 'Offset settings dialog failed'}
 $fixture.Activate();Wait-Rpa 100;Invoke-ImageStep $updated;Wait-Rpa 150
 if(-not $script:offsetClicked){throw 'Image-relative offset click failed'}
 $decoded=Read-ImageTarget $updated.Target
 try{
  $match=[Drawing.Point]::new($area.X-$bounds.X,$area.Y-$bounds.Y)
  $decoded.OffsetX=-10;$decoded.OffsetY=-5
  $negative=Get-ImageClickPoint $bounds $match $decoded
  if($negative.X -ne $area.X+[int][Math]::Floor($area.Width/2)-10 -or $negative.Y -ne $area.Y+[int][Math]::Floor($area.Height/2)-5){throw 'Negative offset calculation failed'}
  $decoded.OffsetX=30000;$refused=$false
  try{$null=Get-ImageClickPoint $bounds $match $decoded}catch{$refused=$true}
  if(-not $refused){throw 'Off-screen click accepted'}
 }finally{$decoded.Bitmap.Dispose()}
 Write-Output 'PASS: legacy zero offset, settings dialog, actual offset click, negative offset, off-screen rejection'
 Export-ScenarioFile $export @(Get-Steps)
 $vbs=[IO.File]::ReadAllText($export)
 $encoded=([regex]::Matches($vbs,'data = data & "([A-Za-z0-9+/=]+)"')|ForEach-Object{$_.Groups[1].Value}) -join ''
 $source=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($encoded))
 $errors=$null;[void][Management.Automation.Language.Parser]::ParseInput($source,[ref]$null,[ref]$errors)
 if($errors.Count){throw ($errors|Out-String)}
 if($source -match '(?m)^\. \(Join-Path \$PSScriptRoot'){throw 'Image module not embedded'}
 $line=@($source -split "`n"|Where-Object{$_.StartsWith('$script:AutoRunScenario = ')})[0]
 $doc=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($line.Split([char]39)[1]))|ConvertFrom-Json
 if(($doc.Steps[0].Target|ConvertFrom-Json).ImageData -ne $data){throw 'PNG not embedded'}
 if(($doc.Steps[0].Target|ConvertFrom-Json).OffsetX -ne $dx){throw 'Click offset not exported'}
 Write-Output 'PASS: screen search and actual click, image validation, embedded PNG and standalone export'
}finally{
 if($crop){$crop.Dispose()};$picture.Image=$null;$fixture.Dispose();$pattern.Dispose()
 Remove-Item -LiteralPath $export -Force -ErrorAction SilentlyContinue
}
$form.Hide()
'@
$path=Join-Path $PSScriptRoot '_image-verify.ps1'
try{
 $source=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'RpaStudio.ps1'))
 [IO.File]::WriteAllText($path,$source.Replace('[void]$form.ShowDialog()',$test),[Text.UTF8Encoding]::new($true))
 & powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File $path
 if($LASTEXITCODE -ne 0){throw 'Image verification failed'}
}finally{Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue}
