[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$required = @('SKILL.md','agents/openai.yaml','references/workflow.md','references/layout-library.md','references/design-contract.md','references/acceptance.md','examples/layout-library.json','examples/slide-plan.json','examples/asset-manifest.json','examples/state.json','scripts/common.ps1','scripts/check_environment.ps1','scripts/assemble_powerpoint.ps1','scripts/export_slides.ps1','scripts/smoke_test.ps1')
foreach ($p in $required) { if (-not (Test-Path -LiteralPath (Join-Path $root $p))) { throw "Missing file: $p" } }
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
'Package files, PowerShell syntax and example JSON: PASS'
