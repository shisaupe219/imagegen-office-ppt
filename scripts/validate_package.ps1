[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$required = @('SKILL.md','agents/openai.yaml','references/workflow.md','references/layout-library.md','references/design-contract.md','references/acceptance.md','examples/layout-library.json','examples/slide-plan.json','examples/asset-manifest.json','examples/state.json','scripts/common.ps1','scripts/check_environment.ps1','scripts/assemble_powerpoint.ps1','scripts/export_slides.ps1','scripts/smoke_test.ps1','scripts/validate_presentation_plan.ps1','references/layout-semantics.md','references/output-and-preview.md','scripts/resolve_output_path.ps1','scripts/assemble_image_powerpoint.ps1','scripts/export_titlebar_board.ps1','scripts/apply_titlebar_style.ps1')
foreach ($p in $required) { if (-not (Test-Path -LiteralPath (Join-Path $root $p))) { throw "Missing file: $p" } }
if (-not (Test-Path -LiteralPath (Join-Path $root 'references/content-modes.md'))) { throw 'Missing content mode rules' }
if (-not (Test-Path -LiteralPath (Join-Path $root 'references/visual-components.md'))) { throw 'Missing visual component rules' }
foreach ($p in @('references/visual-design.md','references/color-selection.md','assets/palettes.json','assets/palette-board.png','assets/palette-board.json','scripts/create_theme.ps1','scripts/export_palette_board.ps1')) {
    if (-not (Test-Path -LiteralPath (Join-Path $root $p))) { throw "Missing v2.2 file: $p" }
}
$skill = Get-Content -LiteralPath (Join-Path $root 'SKILL.md') -Raw -Encoding UTF8
if ($skill -notmatch '(?s)^---\r?\nname: imagegen-office-ppt\r?\ndescription: [^\r\n]+\r?\n---') { throw 'Invalid required SKILL frontmatter' }
foreach ($match in [regex]::Matches($skill, '\]\((references/[^)]+)\)')) {
    if (-not (Test-Path -LiteralPath (Join-Path $root $match.Groups[1].Value))) { throw 'Broken reference link' }
}
foreach ($p in Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.ps1') {
    $tokens = $null; $errors = $null
    [void][Management.Automation.Language.Parser]::ParseFile($p.FullName, [ref]$tokens, [ref]$errors)
    if ($errors.Count) { throw "$($p.Name): $($errors.Message -join '; ')" }
}
foreach ($p in Get-ChildItem -LiteralPath (Join-Path $root 'examples') -Filter '*.json') { [void](Get-Content -LiteralPath $p.FullName -Raw -Encoding UTF8 | ConvertFrom-Json) }
$state = Get-Content -LiteralPath (Join-Path $root 'examples/state.json') -Raw -Encoding UTF8 | ConvertFrom-Json
foreach ($field in @('content_mode','content_mode_confirmation','protected_information','page_content_mapping','approved_omissions','bottom_annotations','integrated_image_annotations','palette_id','primary_color','palette_confirmation_message','theme_path','visual_review','output_mode','preview_styles','preview_page_ids','selected_style','titlebar_style','desktop_path','final_output_path')) {
    if ($state.PSObject.Properties.Name -notcontains $field) { throw "Missing state field: $field" }
}
$palettes = Get-Content -LiteralPath (Join-Path $root 'assets/palettes.json') -Raw -Encoding UTF8 | ConvertFrom-Json
if (@($palettes.palettes).Count -ne 6 -or @($palettes.palettes.id | Select-Object -Unique).Count -ne 6) { throw 'Invalid palette count or IDs' }
foreach ($palette in $palettes.palettes) {
    foreach ($role in @('primary','accent','pale','body','background')) { if ($palette.$role -notmatch '^[0-9A-Fa-f]{6}$') { throw 'Invalid palette HEX' } }
}
$board=Get-Content -LiteralPath (Join-Path $root 'assets/palette-board.json') -Raw -Encoding UTF8 | ConvertFrom-Json
if ($board.palette_source_sha256 -ne (Get-FileHash -LiteralPath (Join-Path $root 'assets/palettes.json')).Hash) { throw 'Palette board configuration stale; re-export it' }
'Package files, PowerShell syntax and example JSON: PASS'
