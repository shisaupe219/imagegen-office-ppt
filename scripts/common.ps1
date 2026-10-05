Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Read-Json([string]$Path) {
    Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
}
function Value-Or($Object, [string]$Key, $Default) {
    if ($null -ne $Object -and $Object.PSObject.Properties.Name -contains $Key) { return $Object.$Key }
    return $Default
}
function Color-Value([string]$Hex) {
    $Hex = $Hex.TrimStart('#')
    if ($Hex -notmatch '^[0-9A-Fa-f]{6}$') { throw "Invalid RGB color: $Hex" }
    [Convert]::ToInt32($Hex.Substring(0,2),16) + 256 * [Convert]::ToInt32($Hex.Substring(2,2),16) + 65536 * [Convert]::ToInt32($Hex.Substring(4,2),16)
}
function Release-Com($Object) {
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
    }
}
function Check-Box($Box, $Canvas) {
    foreach ($key in @('x','y','w','h')) {
        if ($Box.PSObject.Properties.Name -notcontains $key) { throw "Box missing $key" }
        $n = [double]$Box.$key
        if ([double]::IsNaN($n) -or [double]::IsInfinity($n)) { throw "Invalid box $key" }
    }
    if ($Box.x -lt 0 -or $Box.y -lt 0 -or $Box.w -le 0 -or $Box.h -le 0 -or ($Box.x + $Box.w) -gt ($Canvas.width + 0.01) -or ($Box.y + $Box.h) -gt ($Canvas.height + 0.01)) { throw 'Box lies outside canvas' }
}
function Add-Text($Slide, [string]$Text, $Box, [string]$DefaultFont, [string]$Name) {
    $shape = $Slide.Shapes.AddTextbox(1, [single]$Box.x, [single]$Box.y, [single]$Box.w, [single]$Box.h)
    $shape.Name = $Name
    $shape.Fill.Visible = 0
    $shape.Line.Visible = 0
    $frame = $shape.TextFrame
    $frame.MarginLeft = 0; $frame.MarginRight = 0; $frame.MarginTop = 0; $frame.MarginBottom = 0
    $frame.WordWrap = -1; $frame.AutoSize = 0
    $range = $frame.TextRange
    $range.Text = $Text
    $range.Font.Name = [string](Value-Or $Box 'font' $DefaultFont)
    $range.Font.NameFarEast = [string](Value-Or $Box 'font' $DefaultFont)
    $range.Font.Size = [single](Value-Or $Box 'size' 22)
    $range.Font.Bold = if (Value-Or $Box 'bold' $false) { -1 } else { 0 }
    $range.Font.Color.RGB = Color-Value (Value-Or $Box 'color' '222222')
    $align = Value-Or $Box 'align' 'left'
    $range.ParagraphFormat.Alignment = switch ($align) { 'left' {1} 'center' {2} 'right' {3} default {throw "Unknown alignment: $align"} }
    # PowerPoint can inherit AutoSize from its default text-box style.
    $shape.TextFrame2.AutoSize = 0
    $frame.AutoSize = 0
    $shape.Width = [single]$Box.w; $shape.Height = [single]$Box.h
    $shape.Left = [single]$Box.x; $shape.Top = [single]$Box.y
    return $shape
}
