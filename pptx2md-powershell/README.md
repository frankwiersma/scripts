# PowerShell PPTX to Markdown Converter

This PowerShell script (`pptx2md.ps1`) converts PowerPoint slides (.pptx) into Markdown format, with options for embedding images and notes.

## Requirements
- PowerShell
- Microsoft PowerPoint (for COM automation)

## Usage

```powershell
.\pptx2md.ps1 -PptxPath "path\to\file.pptx" [-OutputPath "out.md"] [-ImageDir "img"] [-ImageWidth 800] [-DisableImage] [-DisableColor] [-DisableNotes] [-EnableSlides] [-MinBlockSize 15]
```

## Parameters

- `-PptxPath` (required): Path to the PowerPoint file.
- `-OutputPath`: Markdown output file (default: `out.md`).
- `-ImageDir`: Directory for slide images (default: `img`).
- `-ImageWidth`: Set max width for images in Markdown.
- `-DisableImage`: Exclude images from output.
- `-DisableColor`: Disable colored text formatting.
- `-DisableNotes`: Exclude slide notes.
- `-EnableSlides`: Separate slides with Markdown delimiters.
- `-MinBlockSize`: Minimum text block size to include.