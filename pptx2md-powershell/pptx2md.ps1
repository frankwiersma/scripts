#Requires -Version 3.0
<#
.SYNOPSIS
    Converts PowerPoint presentations to Markdown format with images.
.DESCRIPTION
    This script converts PowerPoint (.pptx) files to Markdown format, extracting images and formatting text appropriately.
.PARAMETER PptxPath
    The path to the PowerPoint file to convert.
.PARAMETER OutputPath
    Optional. The path where the markdown file should be saved. Defaults to '.\output\out.md'.
.PARAMETER ImageDir
    Optional. The directory where extracted images should be saved. Defaults to '.\output\img'.
.PARAMETER ImageWidth
    Optional. Maximum width for exported images in pixels.
.PARAMETER DisableImage
    Optional. Switch to disable image extraction.
.PARAMETER DisableColor
    Optional. Switch to disable color formatting in output.
.PARAMETER DisableNotes
    Optional. Switch to disable extraction of slide notes.
.PARAMETER EnableSlides
    Optional. Switch to add slide separators in the output.
.PARAMETER MinBlockSize
    Optional. Minimum size for text blocks to be included. Defaults to 15 characters.
.EXAMPLE
    .\pptx2md.ps1 -PptxPath "presentation.pptx"
.EXAMPLE
    .\pptx2md.ps1 -PptxPath "presentation.pptx" -OutputPath "custom\output.md" -ImageWidth 800
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, Position = 0, HelpMessage = "Path to the PowerPoint file")]
    [ValidateScript({
        if (-not (Test-Path $_)) {
            throw "PowerPoint file not found at path: $_"
        }
        if (-not ($_ -match "\.pptx?$")) {
            throw "File must be a PowerPoint file (.ppt or .pptx)"
        }
        $true
    })]
    [string]$PptxPath,

    [Parameter(HelpMessage = "Path for the output markdown file")]
    [string]$OutputPath = (Join-Path (Get-Location) "output\out.md"),

    [Parameter(HelpMessage = "Directory for extracted images")]
    [string]$ImageDir = (Join-Path (Get-Location) "output\img"),

    [Parameter(HelpMessage = "Maximum width for images in pixels")]
    [int]$ImageWidth,

    [Parameter(HelpMessage = "Disable image extraction")]
    [switch]$DisableImage,

    [Parameter(HelpMessage = "Disable color formatting")]
    [switch]$DisableColor,

    [Parameter(HelpMessage = "Disable slide notes extraction")]
    [switch]$DisableNotes,

    [Parameter(HelpMessage = "Enable slide separators")]
    [switch]$EnableSlides,

    [Parameter(HelpMessage = "Minimum text block size")]
    [int]$MinBlockSize = 15
)

# Error handling preferences
$ErrorActionPreference = 'Stop'
$VerbosePreference = 'Continue'

# Required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Script-level variables
$script:powerPoint = $null
$script:presentation = $null

# Helper Functions
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    switch ($Level) {
        'Info'    { Write-Host "[$timestamp] $Message" }
        'Warning' { Write-Warning "[$timestamp] $Message" }
        'Error'   { Write-Error "[$timestamp] $Message" }
    }
}

function Initialize-Directories {
    try {
        # Create output directory for markdown
        $outputDir = Split-Path -Parent $OutputPath
        if (!(Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
            Write-Log "Created output directory: $outputDir"
        }

        # Create image directory if needed
        if (!$DisableImage -and !(Test-Path $ImageDir)) {
            New-Item -ItemType Directory -Path $ImageDir -Force | Out-Null
            Write-Log "Created image directory: $ImageDir"
        }
    }
    catch {
        Write-Log "Failed to create directories: $_" -Level Error
        throw
    }
}

function Initialize-PowerPoint {
    try {
        Write-Log "Initializing PowerPoint..."
        $script:powerPoint = New-Object -ComObject PowerPoint.Application
        if (-not $script:powerPoint) {
            throw "Failed to create PowerPoint COM object"
        }
    }
    catch {
        Write-Log "Failed to initialize PowerPoint: $_" -Level Error
        throw
    }
}

function Convert-ColorToHex {
    param([int]$color)
    try {
        $r = ($color -band 0xFF0000) -shr 16
        $g = ($color -band 0x00FF00) -shr 8
        $b = $color -band 0x0000FF
        return "{0:X2}{1:X2}{2:X2}" -f $r, $g, $b
    }
    catch {
        return "000000"
    }
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

function Save-ShapeAsImage {
    param(
        $slide,
        $shape,
        [int]$slideNumber,
        [int]$shapeNumber
    )
    
    if ($DisableImage) { return $null }
    
    try {
        $imageFile = Join-Path $ImageDir "slide${slideNumber}_shape${shapeNumber}.png"
        
        # Make PowerPoint window visible temporarily
        $wasVisible = $powerPoint.Visible
        $powerPoint.Visible = $true
        
        # Switch to slide view and select shape
        $window = $powerPoint.ActiveWindow
        $window.ViewType = 1  # ppViewSlide
        $slide.Select()
        Start-Sleep -Milliseconds 100
        
        # Try to export image
        $success = $false
        
        # Method 1: Direct export
        try {
            $shape.Export($imageFile, 2)  # 2 = PNG format
            if (Test-Path $imageFile) { $success = $true }
        }
        catch {
            Write-Log "Direct export failed: $_" -Level Warning
        }
        
        # Method 2: Copy-paste method
        if (-not $success) {
            try {
                $shape.Copy()
                $image = [System.Windows.Forms.Clipboard]::GetImage()
                if ($image) {
                    $image.Save($imageFile, [System.Drawing.Imaging.ImageFormat]::Png)
                    $image.Dispose()
                    if (Test-Path $imageFile) { $success = $true }
                }
            }
            catch {
                Write-Log "Clipboard method failed: $_" -Level Warning
            }
        }
        
        # Restore PowerPoint visibility
        $powerPoint.Visible = $wasVisible
        
        if ($success) {
            # Return markdown image link
            $relPath = (Resolve-Path $imageFile -Relative) -replace '\\', '/'
            if ($ImageWidth) {
                return "<img src='$relPath' style='max-width:${ImageWidth}px' />`n"
            }
            return "![]($relPath)`n"
        }
        else {
            Write-Log "Failed to export image from slide ${slideNumber}, shape ${shapeNumber}" -Level Warning
            return $null
        }
    }
    catch {
        Write-Log "Error saving image: $_" -Level Warning
        return $null
    }
}

function Process-Table {
    param($table)
    
    $output = ""
    
    try {
        $rows = $table.Rows.Count
        $cols = $table.Columns.Count
        
        if ($rows -eq 0 -or $cols -eq 0) { return "" }
        
        # Header row
        $output += "| " + (1..$cols | ForEach-Object {
            $cell = $table.Cell(1, $_)
            $text = $cell.Shape.TextFrame.TextRange.Text.Trim()
            $text = $text -replace '\|', '\|'
            $text
        } | Join-String -Separator " | ") + " |`n"
        
        # Separator row
        $output += "|" + (" --- |" * $cols) + "`n"
        
        # Data rows
        for ($row = 2; $row -le $rows; $row++) {
            $output += "| " + (1..$cols | ForEach-Object {
                $cell = $table.Cell($row, $_)
                $text = $cell.Shape.TextFrame.TextRange.Text.Trim()
                $text = $text -replace '\|', '\|'
                $text
            } | Join-String -Separator " | ") + " |`n"
        }
        
        $output += "`n"
    }
    catch {
        Write-Log "Error processing table: $_" -Level Warning
    }
    
    return $output
}

function Process-Shape {
    param(
        $shape,
        $slide,
        [int]$slideNumber,
        [int]$shapeNumber
    )
    
    $output = ""
    
    try {
        switch ($shape.Type) {
            { $_ -in 14, 15 } {  # Title shapes
                if ($shape.HasTextFrame) {
                    $text = Format-Text $shape.TextFrame.TextRange.Text
                    $output += "# $text`n`n"
                }
            }
            
            17 {  # Text box
                if ($shape.HasTextFrame) {
                    $textRange = $shape.TextFrame.TextRange
                    
                    for ($i = 1; $i -le $textRange.Paragraphs().Count; $i++) {
                        $para = $textRange.Paragraphs($i)
                        if ($para.Text.Length -lt $MinBlockSize) { continue }
                        
                        $text = Format-Text -text $para.Text -bold:$para.Font.Bold -italic:$para.Font.Italic -color:$para.Font.Color.RGB
                        
                        if (($para.ParagraphFormat.Bullet.Type -ne 0) -or ($para.ParagraphFormat.Bullet.Visible -eq -1)) {
                            $indent = "  " * ([Math]::Max(0, $para.IndentLevel - 1))
                            $output += "$indent* $text`n"
                        }
                        else {
                            $output += "$text`n`n"
                        }
                    }
                }
            }
            
            19 {  # Table
                $output += Process-Table $shape.Table
            }
            
            13 {  # Picture
                $imageMd = Save-ShapeAsImage -slide $slide -shape $shape -slideNumber $slideNumber -shapeNumber $shapeNumber
                if ($imageMd) {
                    $output += $imageMd + "`n"
                }
            }
        }
    }
    catch {
        Write-Log "Error processing shape ${shapeNumber} on slide ${slideNumber}: $_" -Level Warning
    }
    
    return $output
}

function Convert-Presentation {
    try {
        Write-Log "Opening PowerPoint file: $PptxPath"
        $script:presentation = $powerPoint.Presentations.Open($PptxPath)
        $markdown = ""
        
        $totalSlides = $presentation.Slides.Count
        Write-Log "Processing $totalSlides slides..."
        
        for ($slideNumber = 1; $slideNumber -le $totalSlides; $slideNumber++) {
            Write-Progress -Activity "Converting PowerPoint to Markdown" -Status "Slide $slideNumber of $totalSlides" -PercentComplete (($slideNumber / $totalSlides) * 100)
            
            $slide = $presentation.Slides($slideNumber)
            
            foreach ($shape in $slide.Shapes) {
                $markdown += Process-Shape -shape $shape -slide $slide -slideNumber $slideNumber -shapeNumber $shape.Id
            }
            
            # Process slide notes
            if (!$DisableNotes -and $slide.HasNotesPage) {
                try {
                    $notes = $slide.NotesPage.Shapes | Where-Object { $_.PlaceholderFormat.Type -eq 2 } | Select-Object -First 1
                    if ($notes -and $notes.TextFrame.TextRange.Text.Trim()) {
                        $markdown += "`n---`n"
                        $markdown += Format-Text $notes.TextFrame.TextRange.Text
                        $markdown += "`n---`n`n"
                    }
                }
                catch {
                    Write-Log "Error processing notes on slide ${slideNumber}: $_" -Level Warning
                }
            }
            
            # Add slide separator if enabled
            if ($EnableSlides -and $slideNumber -lt $totalSlides) {
                $markdown += "`n---`n`n"
            }
        }
        
        Write-Progress -Activity "Converting PowerPoint to Markdown" -Completed
        
        Write-Log "Saving markdown to: $OutputPath"
        [System.IO.File]::WriteAllText($OutputPath, $markdown, [System.Text.Encoding]::UTF8)
        
        Write-Log "Conversion completed successfully"
    }
    catch {
        Write-Log "Conversion failed: $_" -Level Error
        throw
    }
}

function Cleanup-Resources {
    Write-Log "Cleaning up resources..."
    if ($script:presentation) {
        $script:presentation.Close()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:presentation) | Out-Null
    }
    if ($script:powerPoint) {
        $script:powerPoint.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:powerPoint) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# Main execution
try {
    Initialize-Directories
    Initialize-PowerPoint
    Convert-Presentation
}
catch {
    Write-Log "Script execution failed: $_" -Level Error
    exit 1
}
finally {
    Cleanup-Resources
}
