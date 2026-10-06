[CmdletBinding()]
param([Parameter(Mandatory)][string]$PlanPath,[Parameter(Mandatory)][ValidateSet('T1','T2','T3','T4')][string]$Style,[Parameter(Mandatory)][string]$OutputPath,[string]$Primary='073D86',[string]$Pale='DFEBF5')
. "$PSScriptRoot/common.ps1"
if (Test-Path -LiteralPath $OutputPath) {throw 'Use a new plan path'}
$plan=Read-Json $PlanPath;$sx=$plan.canvas.width/960;$sy=$plan.canvas.height/540
$plan.title_style.x=32*$sx;$plan.title_style.y=20*$sy;$plan.title_style.w=896*$sx;$plan.title_style.h=48*$sy
$plan.title_style | Add-Member -NotePropertyName color -NotePropertyValue $Primary -Force
foreach ($s in $plan.slides) {
 if ($s.kind -ne 'body') {continue}
 $s.elements=@($s.elements | Where-Object { (Value-Or $_ 'role' '') -ne 'titlebar' -and -not ($_.type -eq 'line' -and $_.y -lt 90*$sy) })
 $e=switch ($Style) {
  'T1' {@{type='line';x=32*$sx;y=77*$sy;w=896*$sx;h=0;color=$Primary;width=1.5;role='titlebar'}}
  'T2' {@{type='shape';shape='rect';x=0;y=0;w=$plan.canvas.width;h=80*$sy;fill=$Primary;line_color='none';role='titlebar'}}
  'T3' {@{type='shape';shape='rect';x=15*$sx;y=19*$sy;w=6*$sx;h=49*$sy;fill=$Primary;line_color='none';role='titlebar'}}
  'T4' {@{type='shape';shape='roundrect';x=20*$sx;y=10*$sy;w=920*$sx;h=68*$sy;fill=$Pale;line_color='none';role='titlebar'}}
 }
 $s.elements=@($e)+@($s.elements)
 $s | Add-Member -NotePropertyName title_color -NotePropertyValue $(if ($Style -eq 'T2') {'FFFFFF'} else {$Primary}) -Force
}
$plan | Add-Member -NotePropertyName titlebar_style -NotePropertyValue $Style -Force
$plan | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
"Titlebar applied: $Style"
