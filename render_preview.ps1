Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$clrBg      = [System.Drawing.ColorTranslator]::FromHtml("#1E1E2E")
$clrCard    = [System.Drawing.ColorTranslator]::FromHtml("#2A2A3D")
$clrInput   = [System.Drawing.ColorTranslator]::FromHtml("#313145")
$clrBorder  = [System.Drawing.ColorTranslator]::FromHtml("#3D3D5C")
$clrText    = [System.Drawing.ColorTranslator]::FromHtml("#CDD6F4")
$clrMuted   = [System.Drawing.ColorTranslator]::FromHtml("#6C7086")
$clrAccent  = [System.Drawing.ColorTranslator]::FromHtml("#0078D4")
$clrSuccess = [System.Drawing.ColorTranslator]::FromHtml("#A6E3A1")
$clrSig     = [System.Drawing.ColorTranslator]::FromHtml("#A6E3A1")
$clrFolder  = [System.Drawing.ColorTranslator]::FromHtml("#89B4FA")

$W = 1100; $H = 860
$bmp = New-Object System.Drawing.Bitmap($W, $H)
$g   = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit

$fontMain  = New-Object System.Drawing.Font("Segoe UI",  9)
$fontBold  = New-Object System.Drawing.Font("Segoe UI",  9,  [System.Drawing.FontStyle]::Bold)
$fontMono  = New-Object System.Drawing.Font("Consolas",  8)
$fontTitle = New-Object System.Drawing.Font("Segoe UI",  8)
$fontSmall = New-Object System.Drawing.Font("Consolas",  7)

function Draw-ClippedText($gr, $txt, $fnt, $br, $x, $y, $maxW) {
    $fmt = New-Object System.Drawing.StringFormat
    $fmt.Trimming    = [System.Drawing.StringTrimming]::EllipsisCharacter
    $fmt.FormatFlags = [System.Drawing.StringFormatFlags]::NoWrap
    $rect = New-Object System.Drawing.RectangleF($x, $y, $maxW, 20)
    $gr.DrawString($txt, $fnt, $br, $rect, $fmt)
    $fmt.Dispose()
}

# 1 Background
$g.Clear($clrBg)

# 2 Title bar
$titleH = 32
$g.FillRectangle((New-Object System.Drawing.SolidBrush($clrCard)), 0, 0, $W, $titleH)
$txB = New-Object System.Drawing.SolidBrush($clrText)
$g.DrawString("Managed Bookmarks Creator  |  v2.11.3.0  |  25/03/2026", $fontTitle, $txB, 8, 9)
# close buttons
$g.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(196,43,28))), ($W-46), 0, 46, $titleH)
$g.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(60,60,80))), ($W-92), 0, 46, $titleH)
$g.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(60,60,80))), ($W-138), 0, 46, $titleH)
$g.DrawString("X", $fontTitle, $txB, ($W-34), 9)
$g.DrawString("O", $fontTitle, $txB, ($W-79), 9)
$g.DrawString("_", $fontTitle, $txB, ($W-126), 9)
$txB.Dispose()

# 3 Toolbar
$toolY = $titleH
$toolH = 34
$g.FillRectangle((New-Object System.Drawing.SolidBrush($clrCard)), 0, $toolY, $W, $toolH)
$sepPen = New-Object System.Drawing.Pen($clrBorder)
$g.DrawLine($sepPen, 0, ($toolY+$toolH), $W, ($toolY+$toolH))

$toolBtns = @("[+] Folder","[+] Sub","[+] Link","[+] Root","|","[e] Edit","[x] Del","|","[^] Up","[v] Down","[<] Undo","[>] Redo","|","[o] Load","[s] Save","[c] Copy","|","[v] Validate")
$txB2 = New-Object System.Drawing.SolidBrush($clrText)
$cx = 6
foreach ($b in $toolBtns) {
    if ($b -eq "|") {
        $g.DrawLine($sepPen, ($cx+4), ($toolY+6), ($cx+4), ($toolY+$toolH-6)); $cx += 14
    } else {
        $sz = $g.MeasureString($b, $fontMain)
        $g.DrawString($b, $fontMain, $txB2, $cx, ($toolY+9))
        $cx += [int]$sz.Width + 8
    }
}
$txB2.Dispose(); $sepPen.Dispose()

# 4 Content layout
$margin   = 12
$sigH     = 108
$statusH  = 26
$bottomY  = $H - $sigH - $statusH - 4
$topY     = $toolY + $toolH + 4
$contH    = $bottomY - $topY
$splitX   = 380
$rightX   = $splitX + 4

# Left panel
$g.FillRectangle((New-Object System.Drawing.SolidBrush($clrInput)), $margin, $topY, ($splitX-$margin), $contH)

# Bar label row
$barH = 34
$g.FillRectangle((New-Object System.Drawing.SolidBrush($clrCard)), $margin, $topY, ($splitX-$margin), $barH)
$txB3 = New-Object System.Drawing.SolidBrush($clrText)
$g.DrawString("Bookmarks bar label:", $fontMain, $txB3, ($margin+6), ($topY+9))
$g.FillRectangle((New-Object System.Drawing.SolidBrush($clrInput)), ($margin+152), ($topY+5), ($splitX-$margin-160), 24)
$g.DrawString("Company Bookmarks", $fontMain, $txB3, ($margin+155), ($topY+9))
$txB3.Dispose()

# Tree items
$treeY  = $topY + $barH + 4
$rowH   = 22
$indent = 18

$items = @(
    @{lbl="[+] Intranet";         dep=0; fld=$true;  sel=$false},
    @{lbl="    HR Portal";         dep=1; fld=$false; sel=$false},
    @{lbl="    IT Helpdesk";       dep=1; fld=$false; sel=$false},
    @{lbl="    Confluence";        dep=1; fld=$false; sel=$false},
    @{lbl="[+] Tools";             dep=0; fld=$true;  sel=$true},
    @{lbl="    Azure Portal";      dep=1; fld=$false; sel=$false},
    @{lbl="    GitHub";            dep=1; fld=$false; sel=$false},
    @{lbl="  [+] Monitoring";      dep=1; fld=$true;  sel=$false},
    @{lbl="      Grafana";         dep=2; fld=$false; sel=$false},
    @{lbl="      Kibana";          dep=2; fld=$false; sel=$false},
    @{lbl="Microsoft 365";         dep=0; fld=$false; sel=$false},
    @{lbl="SharePoint";            dep=0; fld=$false; sel=$false},
    @{lbl="[+] Finance";           dep=0; fld=$true;  sel=$false},
    @{lbl="    SAP Portal";        dep=1; fld=$false; sel=$false},
    @{lbl="    Budget Dashboard";  dep=1; fld=$false; sel=$false}
)

$folderBrush = New-Object System.Drawing.SolidBrush($clrFolder)
$urlBrush    = New-Object System.Drawing.SolidBrush($clrText)
$mutBrush    = New-Object System.Drawing.SolidBrush($clrMuted)

for ($i = 0; $i -lt $items.Count; $i++) {
    $item = $items[$i]
    $iy = $treeY + $i * $rowH
    if (($iy + $rowH) -gt ($bottomY - 6)) { break }
    $ix = $margin + 6 + $item.dep * $indent

    if ($item.sel) {
        $selFill = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(50, 0, 120, 212))
        $g.FillRectangle($selFill, $margin, $iy, ($splitX-$margin), $rowH)
        $selFill.Dispose()
        $selPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(180, 0, 120, 212))
        $g.DrawRectangle($selPen, $margin, $iy, ($splitX-$margin-1), ($rowH-1))
        $selPen.Dispose()
    }

    $color = if ($item.fld) { $folderBrush } else { $urlBrush }
    $font  = if ($item.sel) { $fontBold }    else { $fontMain }
    Draw-ClippedText $g $item.lbl $font $color ($ix+2) ($iy+3) ($splitX-$margin-$ix-10)
}
$folderBrush.Dispose(); $urlBrush.Dispose(); $mutBrush.Dispose()

# Splitter
$g.FillRectangle((New-Object System.Drawing.SolidBrush($clrBorder)), $splitX, $topY, 4, $contH)

# 5 Right panel — Tabs
$tabX    = $rightX + 6
$tabW    = $W - $tabX - $margin
$tabBarH = 28
$tabCY   = $topY + $tabBarH

$g.FillRectangle((New-Object System.Drawing.SolidBrush($clrCard)), $tabX, $topY, $tabW, $tabBarH)

$tabNames = @("JSON Preview","JSON Import")
$tbx = $tabX
foreach ($ti in 0..($tabNames.Count-1)) {
    $tn = $tabNames[$ti]
    $sz = $g.MeasureString($tn, $fontMain)
    $tw = [int]$sz.Width + 24
    if ($ti -eq 0) {
        $g.FillRectangle((New-Object System.Drawing.SolidBrush($clrInput)), $tbx, ($topY+2), $tw, ($tabBarH-2))
        $ap = New-Object System.Drawing.Pen($clrAccent); $ap.Width = 2
        $g.DrawLine($ap, $tbx, ($topY+$tabBarH-1), ($tbx+$tw), ($topY+$tabBarH-1)); $ap.Dispose()
        $g.DrawString($tn, $fontMain, (New-Object System.Drawing.SolidBrush($clrText)), ($tbx+12), ($topY+7))
    } else {
        $g.DrawString($tn, $fontMain, (New-Object System.Drawing.SolidBrush($clrMuted)), ($tbx+12), ($topY+7))
    }
    $tbx += $tw + 2
}

$jsonAreaH = $contH - $tabBarH
$g.FillRectangle((New-Object System.Drawing.SolidBrush($clrInput)), $tabX, $tabCY, $tabW, $jsonAreaH)

$jsonLines = @(
    "[",
    "  { `"toplevel_name`": `"Company Bookmarks`" },",
    "  {",
    "    `"name`": `"Intranet`",",
    "    `"children`": [",
    "      { `"name`": `"HR Portal`",   `"url`": `"https://hr.company.local`" },",
    "      { `"name`": `"IT Helpdesk`", `"url`": `"https://helpdesk.company.local`" },",
    "      { `"name`": `"Confluence`",  `"url`": `"https://wiki.company.local`" }",
    "    ]",
    "  },",
    "  {",
    "    `"name`": `"Tools`",",
    "    `"children`": [",
    "      { `"name`": `"Azure Portal`", `"url`": `"https://portal.azure.com`" },",
    "      { `"name`": `"GitHub`",       `"url`": `"https://github.com`" },",
    "      {",
    "        `"name`": `"Monitoring`",",
    "        `"children`": [",
    "          { `"name`": `"Grafana`", `"url`": `"https://grafana.company.local`" },",
    "          { `"name`": `"Kibana`",  `"url`": `"https://kibana.company.local`" }",
    "        ]",
    "      }",
    "    ]",
    "  },",
    "  { `"name`": `"Microsoft 365`", `"url`": `"https://office.com`" },",
    "  { `"name`": `"SharePoint`",    `"url`": `"https://sharepoint.company.local`" },",
    "  {",
    "    `"name`": `"Finance`",",
    "    `"children`": [",
    "      { `"name`": `"SAP Portal`",       `"url`": `"https://sap.company.local`" },",
    "      { `"name`": `"Budget Dashboard`", `"url`": `"https://budget.company.local`" }",
    "    ]",
    "  }",
    "]"
)

$clrKey  = [System.Drawing.ColorTranslator]::FromHtml("#89DCEB")
$clrStr  = [System.Drawing.ColorTranslator]::FromHtml("#A6E3A1")
$clrPunct= [System.Drawing.ColorTranslator]::FromHtml("#CDD6F4")
$clrUrl2 = [System.Drawing.ColorTranslator]::FromHtml("#FAB387")
$bKey  = New-Object System.Drawing.SolidBrush($clrKey)
$bStr  = New-Object System.Drawing.SolidBrush($clrStr)
$bPunct= New-Object System.Drawing.SolidBrush($clrPunct)
$bUrl2 = New-Object System.Drawing.SolidBrush($clrUrl2)

$lineH = 14
$jx = $tabX + 8; $jy = $tabCY + 5
foreach ($line in $jsonLines) {
    if (($jy + $lineH) -gt ($tabCY + $jsonAreaH - 4)) { break }
    $tokens = [System.Text.RegularExpressions.Regex]::Matches($line, '"[^"]*"|[{}\[\],:]+|\s+|[^\s{}\[\],:"]+')
    $cx2 = $jx
    foreach ($tok in $tokens) {
        $t = $tok.Value
        $b = $bPunct
        if ($t -match '^"toplevel_name"|^"name"|^"children"|^"url"') { $b = $bKey }
        elseif ($t -match '^"https?://') { $b = $bUrl2 }
        elseif ($t -match '^"') { $b = $bStr }
        $sz = $g.MeasureString($t, $fontMono)
        if ($cx2 + $sz.Width -lt ($tabX + $tabW - 8)) {
            $g.DrawString($t, $fontMono, $b, $cx2, $jy)
        }
        $cx2 += [float]$sz.Width - 1.5
    }
    $jy += $lineH
}
$bKey.Dispose(); $bStr.Dispose(); $bPunct.Dispose(); $bUrl2.Dispose()

# 6 Status bar
$statusY = $bottomY + 2
$g.FillRectangle((New-Object System.Drawing.SolidBrush($clrCard)), 0, $statusY, $W, $statusH)
$sp = New-Object System.Drawing.Pen($clrBorder)
$g.DrawLine($sp, 0, $statusY, $W, $statusY); $sp.Dispose()
$g.DrawString("15 items loaded  |  5 folders  |  10 links  |  Ready", $fontMain, (New-Object System.Drawing.SolidBrush($clrMuted)), $margin, ($statusY+5))

# 7 Signature
$sigY = $statusY + $statusH
$g.FillRectangle((New-Object System.Drawing.SolidBrush($clrBg)), 0, $sigY, $W, $sigH)
$sp2 = New-Object System.Drawing.Pen($clrBorder)
$g.DrawLine($sp2, 0, $sigY, $W, $sigY); $sp2.Dispose()
$sigLines = @(
    "+===========================================+",
    "|  ## ##  WRITTEN & DESIGNED BY DECOSTER.TECH |",
    "|  #####                                      |",
    "+===========================================+",
    "   ## ##ans            scripting@decoster.tech"
)
$sB = New-Object System.Drawing.SolidBrush($clrSig)
$sy = $sigY + 8
foreach ($line in $sigLines) { $g.DrawString($line, $fontMono, $sB, $margin, $sy); $sy += 18 }
$sB.Dispose()

# 8 Save
$outPath = Join-Path $PSScriptRoot "managed_bookmarks_preview.png"
$bmp.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
Write-Host "Saved: $outPath"
