[CmdletBinding()]
param([Parameter(Mandatory)][string]$OutputDirectory)
. "$PSScriptRoot/common.ps1"
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a new test directory' }
[void][IO.Directory]::CreateDirectory($OutputDirectory)
$plan = [pscustomobject]@{
 canvas=@{width=960;height=540;background='FFFFFF'}
 title_style=@{x=32;y=20;w=890;h=45;size=28;color='51416E'}
 page_style=@{x=880;y=512;w=40;h=20;size=12;align='right'}
 slides=@(@{kind='cover';title='Layout test';core_message='Verify actual text layout';evidence=@();sources=@();elements=@(
 @{type='shape';x=80;y=130;w=350;h=170;fill='EBE7F2';line_color='none'},
 @{type='text';text="Centered label`nSecond line";x=80;y=130;w=350;h=170;size=24;align='center';vertical_align='middle';padding_left=20;padding_right=20;padding_top=15;padding_bottom=15;line_spacing=1.25;paragraph_after=4},
 @{type='text';text="Left aligned paragraph`nReadable spacing";x=490;y=130;w=370;h=170;size=24;align='left';vertical_align='top';padding_left=18;padding_right=18;padding_top=15;padding_bottom=15;line_spacing=1.2;paragraph_before=3;paragraph_after=6}
 )})
}
$planPath=Join-Path $OutputDirectory 'plan.json';$manifestPath=Join-Path $OutputDirectory 'manifest.json'
$plan | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $planPath -Encoding UTF8
@{assets=@()} | ConvertTo-Json | Set-Content -LiteralPath $manifestPath -Encoding UTF8
$ppt=Join-Path $OutputDirectory 'layout.pptx'
& "$PSScriptRoot/assemble_powerpoint.ps1" -PlanPath $planPath -ManifestPath $manifestPath -OutputPath $ppt
$app=$null;$deck=$null
try {
 $app=New-Object -ComObject PowerPoint.Application
 $deck=$app.Presentations.Open($ppt,-1,0,0)
 $t=$deck.Slides.Item(1).Shapes.Item('text-2')
 if ($t.TextFrame.VerticalAnchor -ne 3 -or $t.TextFrame.MarginLeft -ne 20 -or $t.TextFrame.TextRange.ParagraphFormat.SpaceWithin -ne 30 -or $t.TextFrame.TextRange.ParagraphFormat.SpaceAfter -ne 4) {throw 'Text parameters not persisted'}
 $deck.Slides.Item(1).Export((Join-Path $OutputDirectory 'layout.png'),'PNG',1920,1080)
} finally {if ($null -ne $deck) {$deck.Close();Release-Com $deck};Release-Com $app}
# The validator must reject cross-page reuse even when assets have different IDs.
$plan | Add-Member -NotePropertyName presentation_spec -NotePropertyValue @{speaker='Speaker';presentation_date='2026-10-06'}
$plan.slides=@(@{kind='cover';title='Title';elements=@(@{type='text';text='Speaker 2026-10-06'},@{type='image';asset_id='a'})},@{kind='ending';title='End';elements=@(@{type='image';asset_id='b'})})
$plan | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $planPath -Encoding UTF8
@{assets=@(@{id='a';asset_kind='illustration';sha256='same'},@{id='b';asset_kind='illustration';sha256='same'})} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
$rejected=$false
try {& "$PSScriptRoot/validate_presentation_plan.ps1" -PlanPath $planPath -ManifestPath $manifestPath} catch {if ($_.Exception.Message -like '*reused across pages*') {$rejected=$true} else {throw}}
if (-not $rejected) {throw 'Duplicate illustration was accepted'}
'Persisted layout properties and duplicate illustration rejection: PASS'
