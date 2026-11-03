Add-Type -AssemblyName System.Drawing
Add-Type -ReferencedAssemblies System.Drawing @'
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
public static class RpaImageMatcher {
 [DllImport("user32.dll")] public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
 [DllImport("user32.dll")] static extern int GetSystemMetrics(int index);
 [DllImport("user32.dll")] static extern bool SetCursorPos(int x,int y);
 [DllImport("user32.dll")] static extern void mouse_event(uint flags,uint x,uint y,uint data,UIntPtr extra);
 public static void Click(int x,int y) {
  var old=SetThreadDpiAwarenessContext(new IntPtr(-4));
  try{if(!SetCursorPos(x,y))throw new InvalidOperationException("マウスを移動できません。");mouse_event(2,0,0,0,UIntPtr.Zero);mouse_event(4,0,0,0,UIntPtr.Zero);}
  finally{if(old!=IntPtr.Zero)SetThreadDpiAwarenessContext(old);}
 }
 [StructLayout(LayoutKind.Sequential)] struct Rect { public int left,top,right,bottom; }
 [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr window,out Rect rect);
 [DllImport("user32.dll")] static extern IntPtr MonitorFromPoint(Point point,uint flags);
 public static bool IsOnScreen(Point point) {
  var old=SetThreadDpiAwarenessContext(new IntPtr(-4));
  try{return MonitorFromPoint(point,0)!=IntPtr.Zero;}finally{if(old!=IntPtr.Zero)SetThreadDpiAwarenessContext(old);}
 }
 public static Rectangle WindowBounds(IntPtr window) {
  var old=SetThreadDpiAwarenessContext(new IntPtr(-4));
  try{Rect r;if(!GetWindowRect(window,out r))throw new InvalidOperationException("Window not found");return Rectangle.FromLTRB(r.left,r.top,r.right,r.bottom);}
  finally{if(old!=IntPtr.Zero)SetThreadDpiAwarenessContext(old);}
 }
 public static Rectangle Bounds() {
  var old=SetThreadDpiAwarenessContext(new IntPtr(-4));
  try{return new Rectangle(GetSystemMetrics(76),GetSystemMetrics(77),GetSystemMetrics(78),GetSystemMetrics(79));}
  finally{if(old!=IntPtr.Zero)SetThreadDpiAwarenessContext(old);}
 }
 public static Bitmap Capture() {
  var old=SetThreadDpiAwarenessContext(new IntPtr(-4));Bitmap bitmap=null;
  try{var r=Bounds();bitmap=new Bitmap(r.Width,r.Height,PixelFormat.Format24bppRgb);
   using(var g=Graphics.FromImage(bitmap))g.CopyFromScreen(r.Location,Point.Empty,r.Size);
   return bitmap;
  }catch{if(bitmap!=null)bitmap.Dispose();throw;}
  finally{if(old!=IntPtr.Zero)SetThreadDpiAwarenessContext(old);}
 }
 static byte[] Pixels(Bitmap bitmap,out int stride) {
  var data=bitmap.LockBits(new Rectangle(0,0,bitmap.Width,bitmap.Height),ImageLockMode.ReadOnly,PixelFormat.Format24bppRgb);
  try { stride=Math.Abs(data.Stride);var bytes=new byte[stride*bitmap.Height];
   for(int y=0;y<bitmap.Height;y++)Marshal.Copy(IntPtr.Add(data.Scan0,y*data.Stride),bytes,y*stride,stride);
   return bytes;
  } finally { bitmap.UnlockBits(data); }
 }
 public static Bitmap Normalize(Image image) {
  var result=new Bitmap(image.Width,image.Height,PixelFormat.Format24bppRgb);
  using(var g=Graphics.FromImage(result))g.DrawImage(image,new Rectangle(0,0,image.Width,image.Height),0,0,image.Width,image.Height,GraphicsUnit.Pixel);
  return result;
 }
 public static bool HasDetail(Bitmap bitmap) {
  int stride;var pixels=Pixels(bitmap,out stride);int min=765,max=0;
  for(int y=0;y<bitmap.Height;y++)for(int x=0;x<bitmap.Width;x++){
   int p=y*stride+x*3,v=pixels[p]+pixels[p+1]+pixels[p+2];min=Math.Min(min,v);max=Math.Max(max,v);
  }return max-min>=60;
 }
 public static Task<Point[]> FindAsync(Bitmap screen,Bitmap target,int tolerance,CancellationToken token) {
  return Task.Run(()=>Find(screen,target,tolerance,token),token);
 }
 static Point[] Find(Bitmap screen,Bitmap target,int tolerance,CancellationToken token) {
  int sw=screen.Width,sh=screen.Height,tw=target.Width,th=target.Height,ss,ts;
  var source=Pixels(screen,out ss);var template=Pixels(target,out ts);
  var points=new List<Point>();
  int[] sampleX={0,tw/2,tw-1,tw/4,tw*3/4};int[] sampleY={0,th/2,th-1,th/4,th*3/4};
  for(int y=0;y<=sh-th;y++){
   token.ThrowIfCancellationRequested();
   for(int x=0;x<=sw-tw;x++){
    bool match=true;
    for(int sy=0;sy<sampleY.Length&&match;sy++)for(int sx=0;sx<sampleX.Length&&match;sx++){
     int a=(y+sampleY[sy])*ss+(x+sampleX[sx])*3,b=sampleY[sy]*ts+sampleX[sx]*3;
     for(int c=0;c<3;c++)if(Math.Abs(source[a+c]-template[b+c])>tolerance){match=false;break;}
    }
    if(!match)continue;
    for(int ty=0;ty<th&&match;ty++){
     token.ThrowIfCancellationRequested();
     int a=(y+ty)*ss+x*3,b=ty*ts;
     for(int i=0;i<tw*3;i++)if(Math.Abs(source[a+i]-template[b+i])>tolerance){match=false;break;}
    }
    if(!match)continue;
    // Merge only the tiny positional variations caused by the color tolerance.
    if(points.Count==0)points.Add(new Point(x,y));
    else if(Math.Abs(points[0].X-x)>2||Math.Abs(points[0].Y-y)>2){points.Add(new Point(x,y));return points.ToArray();}
   }
  }return points.ToArray();
 }
}
'@
function Get-RpaScreenshot {
 [RpaImageMatcher]::Capture()
}
function Read-ImageTarget($targetJson) {
 $target=$targetJson|ConvertFrom-Json -ErrorAction Stop
 if($target.Kind -ne 'Image' -or -not $target.ImageData -or $target.ImageData.Length -gt 1400000){throw '「画像取得」で見本画像を設定してください。'}
 $tolerance=0
 if(-not [int]::TryParse([string]$target.Tolerance,[ref]$tolerance) -or $tolerance -lt 0 -or $tolerance -gt 100){throw '画像の許容差は0～100で指定してください。'}
 $offsetX=Get-ImageOffset $target 'OffsetX'
 $offsetY=Get-ImageOffset $target 'OffsetY'
 $stream=$null;$image=$null;$bitmap=$null
 try{
  $stream=[IO.MemoryStream]::new([Convert]::FromBase64String($target.ImageData))
  $image=[Drawing.Image]::FromStream($stream)
  if($image.Width -lt 8 -or $image.Height -lt 8 -or $image.Width -gt 600 -or $image.Height -gt 400){throw '見本画像は幅8～600px、高さ8～400pxで指定してください。'}
  $bitmap=[RpaImageMatcher]::Normalize($image)
  if(-not [RpaImageMatcher]::HasDetail($bitmap)){throw '見本画像が単色に近いため識別できません。文字やマークを含めて取得してください。'}
  $result=[pscustomobject]@{Bitmap=$bitmap;Tolerance=$tolerance;OffsetX=$offsetX;OffsetY=$offsetY}
  $bitmap=$null
  return $result
 }finally{
  if($bitmap){$bitmap.Dispose()};if($image){$image.Dispose()};if($stream){$stream.Dispose()}
 }
}
function Get-ImageOffset($target,[string]$name) {
 if($null -eq $target.$name){return 0}
 $value=0
 if(-not [int]::TryParse([string]$target.$name,[ref]$value) -or $value -lt -30000 -or $value -gt 30000){throw '画像のクリック位置のずらし量は-30000～30000pxの整数で指定してください。'}
 return $value
}
function Get-ImageClickPoint($bounds,$match,$target) {
 $x=$bounds.X+$match.X+[int][Math]::Floor($target.Bitmap.Width/2)+$target.OffsetX
 $y=$bounds.Y+$match.Y+[int][Math]::Floor($target.Bitmap.Height/2)+$target.OffsetY
 $point=[Drawing.Point]::new($x,$y)
 if(-not $bounds.Contains($point) -or -not [RpaImageMatcher]::IsOnScreen($point)){throw '画像からずらしたクリック位置が画面の外にあります。画像設定の横・縦のずらし量を確認してください。'}
 return $point
}
function Find-RpaImage($screen,$target,[int]$tolerance) {
 $cancel=New-Object Threading.CancellationTokenSource
 $task=$null
 try{
  $task=[RpaImageMatcher]::FindAsync($screen,$target,$tolerance,$cancel.Token)
  while(-not $task.IsCompleted){Wait-Rpa 20}
  $task.GetAwaiter().GetResult()
 }catch{
  $cancel.Cancel()
  if($task){while(-not $task.IsCompleted){Start-Sleep -Milliseconds 5}}
  throw
 }finally{$cancel.Dispose()}
}
function Invoke-ImageStep($step) {
 $target=Read-ImageTarget $step.Target
 try{
  $clock=[Diagnostics.Stopwatch]::StartNew()
  do{
   Wait-Rpa 0
   $bounds=[RpaImageMatcher]::Bounds()
   $screen=Get-RpaScreenshot
   try{$matches=@(Find-RpaImage $screen $target.Bitmap $target.Tolerance)}finally{$screen.Dispose()}
   if($matches.Count -gt 1){throw '同じ画像が複数見つかりました。周囲の文字などを含めて画像を取得し直してください。'}
   if($matches.Count -eq 1){
    $x=$bounds.X+$matches[0].X+[int][Math]::Floor($target.Bitmap.Width/2)
    $y=$bounds.Y+$matches[0].Y+[int][Math]::Floor($target.Bitmap.Height/2)
    if($step.Action -ne '画像待機'){
     $point=Get-ImageClickPoint $bounds $matches[0] $target
     Wait-Rpa 0
     [RpaImageMatcher]::Click($point.X,$point.Y)
     if($step.Action -eq '画像ダブルクリック'){Wait-Rpa 80;[RpaImageMatcher]::Click($point.X,$point.Y)}
     Write-Log "画像クリック: X=$($point.X) / Y=$($point.Y)（中心から横 $($target.OffsetX)px / 縦 $($target.OffsetY)px）"
    }
    Write-Log "画像検出: X=$x / Y=$y"
    return
   }
   Wait-Rpa 200
  }while($clock.ElapsedMilliseconds -lt 10000)
  throw '見本画像が見つかりません（10秒）。対象画面、表示倍率、解像度、色やテーマを確認してください。'
 }finally{$target.Bitmap.Dispose()}
}
function Capture-ImageTarget {
 $old=[RpaImageMatcher]::SetThreadDpiAwarenessContext([IntPtr]::new(-4))
 try{Capture-ImageTargetCore}finally{if($old -ne [IntPtr]::Zero){[void][RpaImageMatcher]::SetThreadDpiAwarenessContext($old)}}
}
function Capture-ImageTargetCore {
 $screen=Get-RpaScreenshot
 $bounds=[RpaImageMatcher]::Bounds()
 $overlay=New-Object Windows.Forms.Form
 $overlay.FormBorderStyle='None';$overlay.StartPosition='Manual';$overlay.Bounds=$bounds
 $overlay.TopMost=$true;$overlay.Cursor=[Windows.Forms.Cursors]::Cross;$overlay.KeyPreview=$true
 $overlay.BackgroundImage=$screen;$overlay.BackgroundImageLayout='None'
 $state=[pscustomobject]@{Start=[Drawing.Point]::Empty;Rect=[Drawing.Rectangle]::Empty;Dragging=$false}
 $overlay.Add_MouseDown({param($sender,$e)
  if($e.Button -eq 'Left'){$state.Start=$e.Location;$state.Dragging=$true}
 })
 $overlay.Add_MouseMove({param($sender,$e)
  if($state.Dragging){
   $state.Rect=[Drawing.Rectangle]::FromLTRB([Math]::Min($state.Start.X,$e.X),[Math]::Min($state.Start.Y,$e.Y),[Math]::Max($state.Start.X,$e.X),[Math]::Max($state.Start.Y,$e.Y))
   $overlay.Invalidate()
  }
 })
 $overlay.Add_MouseUp({param($sender,$e)
  if($e.Button -eq 'Left' -and $state.Dragging){$state.Dragging=$false;$overlay.DialogResult='OK'}
 })
 $overlay.Add_KeyDown({param($sender,$e) if($e.KeyCode -eq 'Escape'){$overlay.DialogResult='Cancel'}})
 $overlay.Add_Paint({param($sender,$e)
  $e.Graphics.FillRectangle([Drawing.Brushes]::White,10,10,600,32)
  $e.Graphics.DrawString('ボタン等をドラッグで囲んでください。Escで中止（最大600×400px）',$form.Font,[Drawing.Brushes]::Black,15,15)
  if($state.Rect.Width -gt 0 -and $state.Rect.Height -gt 0){$e.Graphics.DrawRectangle([Drawing.Pens]::Red,$state.Rect)}
 })
 $crop=$null;$stream=$null
 try{
  if($overlay.ShowDialog() -ne 'OK'){return $null}
  $rect=$state.Rect
  if($rect.Width -lt 8 -or $rect.Height -lt 8 -or $rect.Width -gt 600 -or $rect.Height -gt 400){throw '幅8～600px、高さ8～400pxで、文字やマークを含む小さな範囲を選択してください。'}
  $crop=$screen.Clone($rect,[Drawing.Imaging.PixelFormat]::Format24bppRgb)
  if(-not [RpaImageMatcher]::HasDetail($crop)){throw '単色に近い画像です。文字やマークを含めて取得してください。'}
  $stream=New-Object IO.MemoryStream
  $crop.Save($stream,[Drawing.Imaging.ImageFormat]::Png)
  $data=[Convert]::ToBase64String($stream.ToArray())
  if($data.Length -gt 1400000){throw '見本画像が大きすぎます。小さな範囲を選択してください。'}
  return [pscustomobject]@{Kind='Image';ImageData=$data;Width=$crop.Width;Height=$crop.Height;Tolerance=20;OffsetX=0;OffsetY=0}
 }finally{
  $overlay.BackgroundImage=$null;$overlay.Dispose()
  if($stream){$stream.Dispose()};if($crop){$crop.Dispose()};$screen.Dispose()
 }
}
