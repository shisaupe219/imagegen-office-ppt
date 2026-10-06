[CmdletBinding()]
param([Parameter(Mandatory)][string]$OutputDirectory,[string]$Primary='073D86',[string]$Pale='DFEBF5')
. "$PSScriptRoot/common.ps1"
if (Test-Path -LiteralPath $OutputDirectory) {throw 'Use a new output directory'}
[void][IO.Directory]::CreateDirectory($OutputDirectory)
$app=$null;$deck=$null
try {
 $app=New-Object -ComObject PowerPoint.Application;$deck=$app.Presentations.Add(0)
 $deck.PageSetup.SlideWidth=960;$deck.PageSetup.SlideHeight=540
 $slide=$deck.Slides.Add(1,12);$slide.DisplayMasterShapes=0;$slide.FollowMasterBackground=0
 $slide.Background.Fill.Solid();$slide.Background.Fill.ForeColor.RGB=Color-Value 'FFFFFF'
 $names=@('T1 简洁横线型','T2 深色通栏型','T3 左侧标记型','T4 浅色圆角型')
 for ($i=0;$i -lt 4;$i++) {
  $x=32+($i%2)*464;$y=40+[Math]::Floor($i/2)*245
  $panel=[pscustomobject]@{x=$x;y=$y;w=432;h=206;fill='FFFFFF';shape='rect';line_color='DDE4EA';line_width=1}
  [void](Add-NativeShape $slide $panel "panel-$i")
  if ($i -eq 1) {[void](Add-NativeShape $slide ([pscustomobject]@{x=$x;y=$y;w=432;h=67;fill=$Primary;shape='rect';line_color='none'}) 'bar')}
  if ($i -eq 2) {[void](Add-NativeShape $slide ([pscustomobject]@{x=$x+12;y=$y+16;w=5;h=40;fill=$Primary;shape='rect';line_color='none'}) 'bar')}
  if ($i -eq 3) {[void](Add-NativeShape $slide ([pscustomobject]@{x=$x+9;y=$y+9;w=414;h=55;fill=$Pale;shape='roundrect';line_color='none'}) 'bar')}
  [void](Add-Text $slide '一、章节标题示例' ([pscustomobject]@{x=$x+25;y=$y+22;w=382;h=38;size=24;bold=$true;color=$(if($i -eq 1){'FFFFFF'}else{$Primary})}) 'Microsoft YaHei' "title-$i")
  if ($i -eq 0) {[void](Add-NativeLine $slide ([pscustomobject]@{x=$x+25;y=$y+65;w=382;h=0;color=$Primary;width=1.5}) 'rule')}
  [void](Add-Text $slide '本页小标题：统一层级与阅读顺序' ([pscustomobject]@{x=$x+25;y=$y+87;w=382;h=40;size=18;color='24364C'}) 'Microsoft YaHei' "sub-$i")
  [void](Add-Text $slide $names[$i] ([pscustomobject]@{x=$x+25;y=$y+149;w=382;h=33;size=21;bold=$true;color=$Primary}) 'Microsoft YaHei' "label-$i")
 }
 $deck.SaveAs((Join-Path $OutputDirectory 'titlebar-board.pptx'),24)
 $slide.Export((Join-Path $OutputDirectory 'titlebar-board.png'),'PNG',1920,1080)
 'Four titlebar samples exported'
} finally {if ($null -ne $deck) {$deck.Close();Release-Com $deck};Release-Com $app}
