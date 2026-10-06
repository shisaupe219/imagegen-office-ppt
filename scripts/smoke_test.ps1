[CmdletBinding()]
param([Parameter(Mandatory)][string]$OutputDirectory)
. "$PSScriptRoot/common.ps1"
$dest = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $dest) { throw 'Use a new test directory' }
[void][IO.Directory]::CreateDirectory($dest)
$root = Split-Path -Parent $PSScriptRoot
& "$PSScriptRoot/validate_package.ps1"
$testPlan = Join-Path $root 'examples/slide-plan.json'
$plan = Read-Json $testPlan
foreach ($slide in $plan.slides) {
    foreach ($element in $slide.elements) {
        Check-Box $element $plan.canvas
        if ($element.type -eq 'chart') { Check-Chart $element }
    }
}
$bad = [pscustomobject]@{chart_type='column';x=0;y=0;w=400;h=200;categories=@('A');series=@([pscustomobject]@{name='invalid';values=@(-1)})}
$rejected = $false
try { Check-Chart $bad } catch { $rejected = $true }
if (-not $rejected) { throw 'Negative values were not rejected' }
$ppt = Join-Path $dest 'smoke.pptx'
& "$PSScriptRoot/assemble_powerpoint.ps1" -PlanPath $testPlan -ManifestPath (Join-Path $root 'examples/asset-manifest.json') -OutputPath $ppt
$audit = Read-Json "$ppt.audit.json"
if (@($audit.possible_text_overflow).Count) { throw 'Text overflow in smoke example' }
& "$PSScriptRoot/export_slides.ps1" -PresentationPath $ppt -OutputDirectory (Join-Path $dest 'pages')
& "$PSScriptRoot/export_contact_sheet.ps1" -PreviewDirectory (Join-Path $dest 'pages') -OutputDirectory (Join-Path $dest 'overview') -Limit 20 -Columns 2
'Smoke test passed; inspect exported pages before design acceptance.'
