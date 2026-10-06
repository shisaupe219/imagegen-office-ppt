[CmdletBinding(DefaultParameterSetName='Preset')]
param(
    [Parameter(Mandatory,ParameterSetName='Preset')][ValidateSet('A','B','C','D','E','F')][string]$PaletteId,
    [Parameter(Mandatory,ParameterSetName='Custom')][ValidatePattern('^[0-9A-Fa-f]{6}$')][string]$CustomPrimary,
    [Parameter(Mandatory)][string]$OutputPath
)
. "$PSScriptRoot/common.ps1"
$target = [IO.Path]::GetFullPath($OutputPath)
if (Test-Path -LiteralPath $target) { throw 'Theme exists; use a versioned path' }
if ($PSCmdlet.ParameterSetName -eq 'Preset') {
    $data=Read-Json (Join-Path (Split-Path -Parent $PSScriptRoot) 'assets/palettes.json')
    $palette=@($data.palettes | Where-Object id -eq $PaletteId)[0]
} else {
    $primary=$CustomPrimary.ToUpperInvariant(); $channels=@()
    foreach ($offset in @(0,2,4)) { $n=[Convert]::ToInt32($primary.Substring($offset,2),16); $channels += ([int][Math]::Round($n*0.12+255*0.88)).ToString('X2') }
    $palette=[pscustomobject]@{id='custom';name='用户自定义';primary=$primary;accent='A45B2B';pale=($channels -join '');body='24364C';background='FFFFFF'}
}
$theme=[pscustomobject]@{
    palette_id=$palette.id; palette_name=$palette.name
    colors=@{primary=$palette.primary;accent=$palette.accent;pale=$palette.pale;body=$palette.body;background=$palette.background}
    title_style=@{x=32;y=18;w=896;h=45;font='Microsoft YaHei';size=31;color=$palette.primary;bold=$true}
    page_style=@{x=884;y=510;w=40;h=20;size=11;color=$palette.body;align='right'}
    body_defaults=@{font='Microsoft YaHei';size=20;color=$palette.body;bold=$false}
    requires_visual_review=$true
}
[void][IO.Directory]::CreateDirectory((Split-Path -Parent $target))
$theme | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $target -Encoding UTF8
"Theme created: $target; apply roles to new plan and verify actual slides"
