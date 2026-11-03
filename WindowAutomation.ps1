Add-Type @'
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
public static class RpaWindows {
 public sealed class WindowInfo { public IntPtr Handle; public string Title, ProcessName, ClassName; public uint ProcessId; }
 delegate bool EnumCallback(IntPtr window, IntPtr param);
 [DllImport("user32.dll")] static extern bool EnumWindows(EnumCallback callback, IntPtr param);
 [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr window);
 [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr window, StringBuilder text, int length);
 [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr window, StringBuilder text, int length);
 [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
 [DllImport("user32.dll")] static extern bool IsIconic(IntPtr window);
 [DllImport("user32.dll")] static extern bool ShowWindowAsync(IntPtr window, int command);
 [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr window);
 [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
 public static WindowInfo[] List() {
  var windows=new List<WindowInfo>();
  EnumWindows(delegate(IntPtr window, IntPtr unused) {
   if(!IsWindowVisible(window))return true;
   var title=new StringBuilder(4096);GetWindowText(window,title,title.Capacity);if(title.Length==0)return true;
   uint id;GetWindowThreadProcessId(window,out id);
   var type=new StringBuilder(512);GetClassName(window,type,type.Capacity);
   try{using(var process=Process.GetProcessById((int)id)){windows.Add(new WindowInfo{Handle=window,Title=title.ToString(),ClassName=type.ToString(),ProcessName=process.ProcessName,ProcessId=id});}}catch{}
   return true;
  },IntPtr.Zero);
  return windows.ToArray();
 }
 public static void Activate(IntPtr window) { if(IsIconic(window))ShowWindowAsync(window,9);SetForegroundWindow(window); }
}
'@
function Resolve-RpaWindow($step) {
 $windows=@([RpaWindows]::List())
 $target=if($step.Target){$step.Target|ConvertFrom-Json -ErrorAction Stop}else{$null}
 if($target -and $target.Process){$windows=@($windows|Where-Object{$_.ProcessName -ieq $target.Process})}
 $found=@($windows|Where-Object{$_.Title -ceq $step.Value})
 if($found.Count -eq 0 -and $target -and $target.Process -and $target.Class){
  $found=@($windows|Where-Object{$_.ClassName -ceq $target.Class})
 }
 if($found.Count -gt 1){throw '対象ウィンドウが複数あります。「ウィンドウ指定」で指定し直すか、対象以外を閉じてください。'}
 if($found.Count -eq 1){return $found[0]}
 return $null
}
function Invoke-WindowActivation($step) {
 $clock=[Diagnostics.Stopwatch]::StartNew()
 do {
  Wait-Rpa 0
  $window=Resolve-RpaWindow $step
  if($window){
   [RpaWindows]::Activate($window.Handle)
   Wait-Rpa 150
   if([RpaWindows]::GetForegroundWindow() -eq $window.Handle){return}
  }
  Wait-Rpa 100
 }while($clock.ElapsedMilliseconds -lt 10000)
 throw 'ウィンドウをアクティブ化できません（10秒）。対象が開いていること、名前・権限を確認してください。Windowsの前面化制限で失敗する場合もあります。'
}
