[CmdletBinding()]
param([Parameter(Mandatory)][string]$OutputDirectory)
. "$PSScriptRoot/common.ps1"
$dest=[IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $dest) { throw 'Use a new palette preview directory' }
$data=Read-Json (Join-Path (Split-Path -Parent $PSScriptRoot) 'assets/palettes.json')
[void][IO.Directory]::CreateDirectory($dest)
$app=$null;$deck=$null
try {
    $app=New-Object -ComObject PowerPoint.Application
    $deck=$app.Presentations.Add(0);$deck.PageSetup.SlideWidth=960;$deck.PageSetup.SlideHeight=660
    $slide=$deck.Slides.Add(1,12);$slide.DisplayMasterShapes=0;$slide.FollowMasterBackground=0
    $slide.Background.Fill.Solid();$slide.Background.Fill.ForeColor.RGB=Color-Value 'FFFFFF'
    [void](Add-Text $slide '选择主色调｜保留参考结构，配色由你决定' ([pscustomobject]@{x=32;y=20;w=896;h=42;size=27;color='24364C';bold=$true}) 'Microsoft YaHei' 'board-title')
    $i=0
    foreach ($palette in $data.palettes) {
        $x=32+($i%2)*464;$y=82+[Math]::Floor($i/2)*186
        [void](Add-NativeShape $slide ([pscustomobject]@{x=$x;y=$y;w=432;h=166;fill='FFFFFF';line_color='D8DFE6'}) "card-$i")
        [void](Add-Text $slide "$($palette.id)  $($palette.name)" ([pscustomobject]@{x=$x+16;y=$y+10;w=400;h=31;size=23;color=$palette.primary;bold=$true}) 'Microsoft YaHei' "name-$i")
        $j=0
        foreach ($entry in @(@('primary','主色'),@('accent','强调'),@('pale','浅色'),@('body','正文'))) {
            [void](Add-NativeShape $slide ([pscustomobject]@{x=$x+16+$j*101;y=$y+53;w=91;h=42;fill=$palette.($entry[0]);line_color='none'}) "chip-$i-$j")
            [void](Add-Text $slide $entry[1] ([pscustomobject]@{x=$x+16+$j*101;y=$y+100;w=91;h=25;size=15;color='24364C';align='center'}) 'Microsoft YaHei' "role-$i-$j")
            $j++
        }
        [void](Add-Text $slide "主色 #$($palette.primary)   标题 · 证据 · 结论" ([pscustomobject]@{x=$x+16;y=$y+132;w=400;h=25;size=15;color=$palette.body}) 'Microsoft YaHei' "sample-$i")
        $i++
    }
    [void](Add-Text $slide '可点选A—F，也可输入自定义HEX；实际选择通过会话选项提交。' ([pscustomobject]@{x=32;y=636;w=896;h=23;size=14;color='24364C'}) 'Microsoft YaHei' 'board-hint')
    $slide.Export((Join-Path $dest 'palette-board.png'),'PNG',2400,1650)
    @{palette_source_sha256=(Get-FileHash -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) 'assets/palettes.json')).Hash;ids=$data.palettes.id} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $dest 'palette-board.json') -Encoding UTF8
    "Palette board exported: $dest"
} finally {
    if ($null -ne $deck) { $deck.Close();Release-Com $deck };Release-Com $app
}
