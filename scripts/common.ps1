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
    $isLine = (Value-Or $Box 'type' '') -eq 'line'
    if ($Box.x -lt 0 -or $Box.y -lt 0 -or $Box.w -lt 0 -or $Box.h -lt 0 -or ((-not $isLine) -and ($Box.w -eq 0 -or $Box.h -eq 0)) -or ($isLine -and $Box.w -eq 0 -and $Box.h -eq 0) -or ($Box.x + $Box.w) -gt ($Canvas.width + 0.01) -or ($Box.y + $Box.h) -gt ($Canvas.height + 0.01)) { throw 'Box lies outside canvas' }
}
function Add-Text($Slide, [string]$Text, $Box, [string]$DefaultFont, [string]$Name) {
    $shape = $Slide.Shapes.AddTextbox(1, [single]$Box.x, [single]$Box.y, [single]$Box.w, [single]$Box.h)
    $shape.Name = $Name
    $shape.Fill.Visible = 0
    $shape.Line.Visible = 0
    $frame = $shape.TextFrame
    $frame.MarginLeft = 0; $frame.MarginRight = 0; $frame.MarginTop = 0; $frame.MarginBottom = 0
    foreach ($entry in @(@('padding_left','MarginLeft'),@('padding_right','MarginRight'),@('padding_top','MarginTop'),@('padding_bottom','MarginBottom'))) {
        $padding = [double](Value-Or $Box $entry[0] 0)
        if ($padding -lt 0) { throw 'Text padding cannot be negative' }
        $frame.($entry[1]) = [single]$padding
    }
    if (($frame.MarginLeft+$frame.MarginRight) -ge $Box.w -or ($frame.MarginTop+$frame.MarginBottom) -ge $Box.h) { throw 'Padding consumes text box' }
    $frame.VerticalAnchor = switch (Value-Or $Box 'vertical_align' 'top') { 'top' {1} 'middle' {3} 'bottom' {4} default {throw 'Unknown vertical alignment'} }
    $frame.WordWrap = -1; $frame.AutoSize = 0
    $range = $frame.TextRange
    $range.Text = $Text
    $range.Font.Name = [string](Value-Or $Box 'font' $DefaultFont)
    $range.Font.NameFarEast = [string](Value-Or $Box 'font' $DefaultFont)
    $range.Font.Size = [single](Value-Or $Box 'size' 22)
    $spacing = [double](Value-Or $Box 'line_spacing' 1.2)
    if ($spacing -le 0) { throw 'Line spacing must be positive' }
    $range.ParagraphFormat.LineRuleWithin = 0
    $range.ParagraphFormat.SpaceWithin = [single]($range.Font.Size * $spacing)
    $range.ParagraphFormat.LineRuleBefore = 0
    $range.ParagraphFormat.LineRuleAfter = 0
    $before = [double](Value-Or $Box 'paragraph_before' 0)
    $after = [double](Value-Or $Box 'paragraph_after' 0)
    if ($before -lt 0 -or $after -lt 0) { throw 'Paragraph spacing cannot be negative' }
    $range.ParagraphFormat.SpaceBefore = [single]$before
    $range.ParagraphFormat.SpaceAfter = [single]$after
    $range.Font.Bold = if (Value-Or $Box 'bold' $false) { -1 } else { 0 }
    $range.Font.Color.RGB = Color-Value (Value-Or $Box 'color' '222222')
    $align = Value-Or $Box 'align' 'left'
    $range.ParagraphFormat.Alignment = switch ($align) { 'left' {1} 'center' {2} 'right' {3} default {throw "Unknown alignment: $align"} }
    foreach ($highlight in @(Value-Or $Box 'highlights' @())) {
        $word = [string]$highlight.text
        if ([string]::IsNullOrEmpty($word)) { throw 'Empty highlight' }
        $offset = 0
        while (($offset = $Text.IndexOf($word, $offset, [StringComparison]::Ordinal)) -ge 0) {
            $part = $range.Characters($offset+1, $word.Length)
            $part.Font.Color.RGB = Color-Value (Value-Or $highlight 'color' 'BE282B')
            $part.Font.Bold = if (Value-Or $highlight 'bold' $true) { -1 } else { 0 }
            $offset += $word.Length
        }
    }
    # PowerPoint can inherit AutoSize from its default text-box style.
    $shape.TextFrame2.AutoSize = 0
    $frame.AutoSize = 0
    $shape.Width = [single]$Box.w; $shape.Height = [single]$Box.h
    $shape.Left = [single]$Box.x; $shape.Top = [single]$Box.y
    return $shape
}

function Add-NativeShape($Slide, $Box, [string]$Name) {
    $code = switch (Value-Or $Box 'shape' 'rect') { 'rect' {1} 'roundrect' {5} 'ellipse' {9} 'rightArrow' {33} 'downArrow' {36} default {throw 'Unsupported shape'} }
    $shape = $Slide.Shapes.AddShape($code,[single]$Box.x,[single]$Box.y,[single]$Box.w,[single]$Box.h)
    $shape.Name = $Name
    $fill = Value-Or $Box 'fill' 'DFEBF5'
    if ($fill -eq 'none') { $shape.Fill.Visible = 0 } else { $shape.Fill.Solid(); $shape.Fill.ForeColor.RGB = Color-Value $fill }
    $gradient = Value-Or $Box 'fill_gradient' $null
    if ($null -ne $gradient) {
        $style = switch (Value-Or $gradient 'direction' 'horizontal') { 'horizontal' {2} 'vertical' {1} default {throw 'Gradient supports horizontal/vertical only'} }
        $shape.Fill.Visible = -1
        $shape.Fill.TwoColorGradient($style,1)
        $shape.Fill.ForeColor.RGB = Color-Value $gradient.from
        $shape.Fill.BackColor.RGB = Color-Value $gradient.to
    }
    $lineColor = Value-Or $Box 'line_color' '073D86'
    if ($lineColor -eq 'none') { $shape.Line.Visible = 0 } else { $shape.Line.ForeColor.RGB = Color-Value $lineColor; $shape.Line.Weight = [single](Value-Or $Box 'line_width' 1) }
    return $shape
}
function Add-NativeLine($Slide, $Box, [string]$Name) {
    $shape = $Slide.Shapes.AddLine([single]$Box.x,[single]$Box.y,[single]($Box.x+$Box.w),[single]($Box.y+$Box.h))
    $shape.Name = $Name
    $shape.Line.ForeColor.RGB = Color-Value (Value-Or $Box 'color' '073D86')
    $shape.Line.Weight = [single](Value-Or $Box 'width' 1)
    if (Value-Or $Box 'end_arrow' $false) { $shape.Line.EndArrowheadStyle = 3 }
    return $shape
}

function Check-Chart($Box) {
    if ($Box.chart_type -notin @('bar','column','line')) { throw 'Unsupported chart type' }
    if (@($Box.categories).Count -lt 1 -or @($Box.series).Count -lt 1) { throw 'Empty chart' }
    if ($Box.chart_type -eq 'line' -and @($Box.categories).Count -lt 2) { throw 'Line chart needs two categories' }
    $minimum = [double](Value-Or $Box 'min' 0)
    if ([double]::IsNaN($minimum) -or [double]::IsInfinity($minimum) -or $minimum -lt 0 -or ($Box.chart_type -ne 'line' -and $minimum -ne 0)) { throw 'Bars require zero baseline; negative range unsupported' }
    $largest = 0.0
    foreach ($series in $Box.series) {
        if (@($series.values).Count -ne @($Box.categories).Count) { throw 'Chart category/value count mismatch' }
        foreach ($value in $series.values) {
            $number = [double]$value
            if ($null -eq $value -or [double]::IsNaN($number) -or [double]::IsInfinity($number) -or $number -lt $minimum) { throw 'Invalid chart value' }
            $largest = [Math]::Max($largest,$number)
        }
    }
    $maximum = [double](Value-Or $Box 'max' ([Math]::Max(1,$largest*1.15)))
    if ([double]::IsNaN($maximum) -or [double]::IsInfinity($maximum) -or $maximum -le $minimum -or $maximum -lt $largest) { throw 'Chart range excludes data' }
    if ($Box.w -lt 240 -or $Box.h -lt 180) { throw 'Chart too small' }
}

function Add-DataChart($Slide, $Box, [string]$Font, [string]$Name) {
    Check-Chart $Box
    $minimum = [double](Value-Or $Box 'min' 0)
    $all = @($Box.series | ForEach-Object { $_.values })
    $maximum = [double](Value-Or $Box 'max' ([Math]::Max(1,($all | Measure-Object -Maximum).Maximum*1.15)))
    $size = [single](Value-Or $Box 'size' 14)
    $bar = $Box.chart_type -eq 'bar'
    $left = $Box.x + $(if ($bar) {120} else {65})
    $top = $Box.y+35; $pw = $Box.w-($left-$Box.x)-55; $ph = $Box.h-100
    $count = @($Box.categories).Count; $seriesCount = @($Box.series).Count
    $colors = @('073D86','BE282B','5295B6','666666')
    for ($tick=0; $tick -le 4; $tick++) {
        $value = $minimum + ($maximum-$minimum)*$tick/4
        if ($bar) { $tx=$left+$pw*$tick/4; $ty=$top+$ph+4; $lw=0; $lh=$ph; $lx=$tx; $ly=$top; $labelX=$tx-22; $labelY=$ty }
        else { $tx=$left; $ty=$top+$ph-$ph*$tick/4; $lw=$pw; $lh=0; $lx=$left; $ly=$ty; $labelX=$left-62; $labelY=$ty-8 }
        [void](Add-NativeLine $Slide ([pscustomobject]@{x=$lx;y=$ly;w=$lw;h=$lh;color='D8E2EB';width=0.5}) "$Name-grid-$tick")
        [void](Add-Text $Slide ($value.ToString('0.##',[Globalization.CultureInfo]::InvariantCulture)) ([pscustomobject]@{x=$labelX;y=$labelY;w=55;h=22;size=$size-2;color='586B81';align='center'}) $Font "$Name-tick-$tick")
    }
    for ($j=0; $j -lt $seriesCount; $j++) {
        $series=$Box.series[$j]; $color=Value-Or $series 'color' $colors[$j % $colors.Count]
        $legendW = ($Box.w-10)/$seriesCount
        [void](Add-Text $Slide $series.name ([pscustomobject]@{x=$Box.x+$j*$legendW;y=$Box.y;w=$legendW;h=25;size=$size;color=$color;bold=$true}) $Font "$Name-legend-$j")
        $prevX=0; $prevY=0
        for ($i=0; $i -lt $count; $i++) {
            $value=[double]$series.values[$i]; $fraction=($value-$minimum)/($maximum-$minimum)
            if ($bar) {
                $band=$ph/$count; $thickness=$band*0.7/$seriesCount
                $bx=$left; $by=$top+$i*$band+$band*0.15+$j*$thickness; $bw=$pw*$fraction; $bh=$thickness*0.9
                $vx=$bx+$bw+4; $vy=$by; $vw=50
            } elseif ($Box.chart_type -eq 'column') {
                $band=$pw/$count; $thickness=$band*0.7/$seriesCount
                $bx=$left+$i*$band+$band*0.15+$j*$thickness; $by=$top+$ph*(1-$fraction); $bw=$thickness*0.9; $bh=$ph*$fraction
                $vx=$bx-10; $vy=$by-23; $vw=$bw+20
            } else {
                $bx=$left+$i*$pw/($count-1); $by=$top+$ph*(1-$fraction); $bw=6; $bh=6
                if ($i -gt 0) {
                    $segment=$Slide.Shapes.AddLine([single]$prevX,[single]$prevY,[single]$bx,[single]$by)
                    $segment.Name="$Name-series-$j-segment-$i"; $segment.Line.ForeColor.RGB=Color-Value $color; $segment.Line.Weight=2
                }
                $prevX=$bx; $prevY=$by; $bx-=3; $by-=3; $vx=$bx-22; $vy=$by-22; $vw=50
            }
            if ($Box.chart_type -eq 'line' -or $value -gt 0) {
                [void](Add-NativeShape $Slide ([pscustomobject]@{x=$bx;y=$by;w=$bw;h=$bh;shape=$(if ($Box.chart_type -eq 'line') {'ellipse'} else {'rect'});fill=$color;line_color='none'}) "$Name-series-$j-value-$i")
            }
            if (Value-Or $Box 'show_values' $true) {
                [void](Add-Text $Slide ($value.ToString('0.##',[Globalization.CultureInfo]::InvariantCulture)) ([pscustomobject]@{x=$vx;y=$vy;w=$vw;h=22;size=$size-2;color=$color;align='center'}) $Font "$Name-value-label-$j-$i")
            }
        }
    }
    for ($i=0; $i -lt $count; $i++) {
        if ($bar) { $cx=$Box.x; $cy=$top+$ph*($i+0.5)/$count-12; $cw=110 }
        else { $band=$pw/$count; $center=$(if ($Box.chart_type -eq 'line') {$left+$i*$pw/($count-1)} else {$left+($i+0.5)*$band}); $cw=[Math]::Min($band,140); $cx=$center-$cw/2; $cy=$top+$ph+4 }
        [void](Add-Text $Slide $Box.categories[$i] ([pscustomobject]@{x=$cx;y=$cy;w=$cw;h=35;size=$size;color='23364D';align='center'}) $Font "$Name-category-$i")
    }
    [void](Add-Text $Slide ([string](Value-Or $Box 'unit' '')) ([pscustomobject]@{x=$Box.x;y=$Box.y+$Box.h-23;w=$Box.w;h=23;size=$size-2;color='586B81';align='right'}) $Font "$Name-unit")
}
