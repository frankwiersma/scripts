# PowerShell PPTX to Markdown Converter
# Requires Office COM automation

param(
    [Parameter(Mandatory=$true)]
    [string]$PptxPath,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "out.md",
    
    [Parameter(Mandatory=$false)]
    [string]$ImageDir = "img",
    
    [Parameter(Mandatory=$false)]
    [int]$ImageWidth = $null,
    
    [Parameter(Mandatory=$false)]
    [switch]$DisableImage,
    
    [Parameter(Mandatory=$false)] 
    [switch]$DisableColor,
    
    [Parameter(Mandatory=$false)]
    [switch]$DisableNotes,
    
    [Parameter(Mandatory=$false)]
    [switch]$EnableSlides,
    
    [Parameter(Mandatory=$false)]
    [int]$MinBlockSize = 15
)

# Initialize PowerPoint
$powerPoint = $null
$presentation = $null

try {
    Write-Host "Initializing PowerPoint..."
    $powerPoint = New-Object -ComObject PowerPoint.Application
    
    # Don't try to set visibility - let PowerPoint manage its own window state
    
} catch {
    Write-Error "Failed to initialize PowerPoint COM object: $($_.Exception.Message)"
    exit 1
}

# Helper functions
function Convert-ColorToHex {
    param([int]$color)
    $r = ($color -band 0xFF0000) -shr 16
    $g = ($color -band 0x00FF00) -shr 8
    $b = $color -band 0x0000FF
    return "{0:X2}{1:X2}{2:X2}" -f $r,$g,$b
}

function Format-Text {
    param(
        [string]$text,
        $color = $null,
        [bool]$bold = $false,
        [bool]$italic = $false
    )
    
    if ([string]::IsNullOrEmpty($text)) { return "" }
    
    $text = $text.Trim()
    
    # Escape special characters
    $text = $text -replace '([\\`*_{}[\]()#+-.!])', '\$1'
    
    if ($bold) { $text = "**$text**" }
    if ($italic) { $text = "*$text*" }
    if ($color -and !$DisableColor) {
        $hexColor = Convert-ColorToHex $color
        $text = "<span style='color:#$hexColor'>$text</span>"
    }
    
    return $text
}

function Save-SlideImage {
    param(
        $shape,
        [int]$slideNumber,
        [int]$shapeNumber
    )
    
    if ($DisableImage) { return $null }
    
    try {
        # Create image directory if it doesn't exist
        if (!(Test-Path $ImageDir)) {
            New-Item -ItemType Directory -Path $ImageDir -Force | Out-Null
        }
        
        $imagePath = Join-Path $ImageDir "slide${slideNumber}_shape${shapeNumber}.png"
        $shape.Export($imagePath, 2) # 2 = PNG format
        
        # Format image markdown
        $relPath = (Resolve-Path $imagePath -Relative) -replace '\\', '/'
        if ($ImageWidth) {
            return "<img src='$relPath' style='max-width:${ImageWidth}px' />`n"
        } else {
            return "![]($relPath)`n"
        }
    } catch {
        Write-Warning ("Failed to save image from slide {0}, shape {1}: {2}" -f $slideNumber, $shapeNumber, $_.Exception.Message)
        return $null
    }
}

function Process-Shape {
    param(
        $shape,
        [int]$slideNumber,
        [int]$shapeNumber
    )
    
    $output = ""
    
    try {
        # Process shape based on type
        switch ($shape.Type) {
            # Title
            {$_ -eq 14 -or $_ -eq 15} {
                if ($shape.HasTextFrame) {
                    $text = Format-Text $shape.TextFrame.TextRange.Text
                    $output += "# $text`n`n"
                }
            }
            
            # Text Box
            {$_ -eq 17} {
                if ($shape.HasTextFrame) {
                    $textRange = $shape.TextFrame.TextRange
                    
                    # Process each paragraph
                    for ($i = 1; $i -le $textRange.Paragraphs().Count; $i++) {
                        $para = $textRange.Paragraphs($i)
                        if ($para.Text.Length -lt $MinBlockSize) { continue }
                        
                        $text = Format-Text -text $para.Text -bold:$para.Font.Bold -italic:$para.Font.Italic -color:$para.Font.Color.RGB
                        
                        # Handle bullet points - check both Bullet.Type and Bullet.Visible
                        if (($para.ParagraphFormat.Bullet.Type -ne 0) -or ($para.ParagraphFormat.Bullet.Visible -eq -1)) {
                            $indent = "  " * ([Math]::Max(0, $para.IndentLevel - 1))
                            $output += "$indent* $text`n"
                        } else {
                            $output += "$text`n`n"
                        }
                    }
                }
            }
            
            # Picture
            {$_ -eq 13} {
                $imageMd = Save-SlideImage -shape $shape -slideNumber $slideNumber -shapeNumber $shapeNumber
                if ($imageMd) {
                    $output += $imageMd + "`n"
                }
            }
            
            # Table
            {$_ -eq 19} {
                $table = $shape.Table
                # Use actual cell contents for header if possible, otherwise use generic Column N
                $headers = @()
                for ($col = 1; $col -le $table.Columns.Count; $col++) {
                    $cellText = $table.Cell(1, $col).Shape.TextFrame.TextRange.Text.Trim()
                    if ([string]::IsNullOrWhiteSpace($cellText)) {
                        $headers += "Column $col"
                    } else {
                        $headers += $cellText
                    }
                }
                
                $output += "| " + ($headers -join " | ") + " |`n"
                $output += "|" + ("---|" * $table.Columns.Count) + "`n"
                
                # Start from row 2 if we used row 1 for headers
                $startRow = if ($headers[0] -notmatch '^Column \d+$') { 2 } else { 1 }
                
                for ($row = $startRow; $row -le $table.Rows.Count; $row++) {
                    $output += "| "
                    for ($col = 1; $col -le $table.Columns.Count; $col++) {
                        $cell = $table.Cell($row, $col)
                        $text = Format-Text $cell.Shape.TextFrame.TextRange.Text
                        $output += "$text | "
                    }
                    $output += "`n"
                }
                $output += "`n"
            }
        }
    } catch {
        Write-Warning ("Error processing shape {0} on slide {1}: {2}" -f $shapeNumber, $slideNumber, $_.Exception.Message)
    }
    
    return $output
}

# Main conversion process
try {
    Write-Host "Opening PowerPoint file: $PptxPath"
    $presentation = $powerPoint.Presentations.Open((Resolve-Path $PptxPath).Path)
    $markdown = ""
    
    Write-Host "Processing $($presentation.Slides.Count) slides..."
    
    # Process each slide
    for ($slideNumber = 1; $slideNumber -le $presentation.Slides.Count; $slideNumber++) {
        Write-Progress -Activity "Converting PowerPoint to Markdown" -Status "Processing slide $slideNumber of $($presentation.Slides.Count)" -PercentComplete (($slideNumber / $presentation.Slides.Count) * 100)
        
        $slide = $presentation.Slides($slideNumber)
        
        # Process shapes on the slide
        for ($shapeNumber = 1; $shapeNumber -le $slide.Shapes.Count; $shapeNumber++) {
            $shape = $slide.Shapes($shapeNumber)
            $markdown += Process-Shape -shape $shape -slideNumber $slideNumber -shapeNumber $shapeNumber
        }
        
        # Add slide notes if enabled
        if (!$DisableNotes -and $slide.HasNotesPage) {
            try {
                $notes = $slide.NotesPage.Shapes | Where-Object {$_.PlaceholderFormat.Type -eq 2} | Select-Object -First 1
                if ($notes -and $notes.TextFrame.TextRange.Text.Trim()) {
                    $markdown += "---`n"
                    $markdown += Format-Text $notes.TextFrame.TextRange.Text
                    $markdown += "`n---`n`n"
                }
            } catch {
                Write-Warning "Could not process notes for slide $slideNumber"
            }
        }
        
        # Add slide delimiter if enabled
        if ($EnableSlides -and $slideNumber -lt $presentation.Slides.Count) {
            $markdown += "`n---`n`n"
        }
    }
    
    Write-Host "Saving markdown to: $OutputPath"
    $markdown | Out-File -FilePath $OutputPath -Encoding utf8 -Force
    
    Write-Host "Conversion completed successfully"
    
} catch {
    Write-Error "Conversion failed: $($_.Exception.Message)"
    exit 1
} finally {
    Write-Host "Cleaning up..."
    if ($presentation) {
        $presentation.Close()
    }
    if ($powerPoint) {
        $powerPoint.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerPoint) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}