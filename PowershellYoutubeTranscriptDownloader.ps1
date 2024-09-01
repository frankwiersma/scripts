# PowerShell script to download and process YouTube subtitles and copy them to a clipboard

$videoUrl = "https://www.youtube.com/watch?v=xxxxxxxxxxx"
yt-dlp --skip-download --write-subs --write-auto-subs --sub-lang en --sub-format ttml --convert-subs srt --output "transcript.%(ext)s" $videoUrl

if (Test-Path "transcript.en.srt") {
    # Remove timestamps, line numbers, and HTML tags
    $content = Get-Content "transcript.en.srt" | 
        Where-Object { $_ -notmatch '^\d+$' -and $_ -notmatch '^\d{2}:\d{2}:\d{2},\d{3} --> \d{2}:\d{2}:\d{2},\d{3}$' } | 
        ForEach-Object { $_ -replace '<[^>]+>', '' } | 
        Where-Object { $_.Trim() -ne '' }

    # Join sentences and format paragraphs
    $paragraph = ""
    $output = @()
    foreach ($line in $content) {
        $paragraph += $line.Trim() + " "
        if ($line -match '[.!?]$') {
            $output += $paragraph.Trim()
            $paragraph = ""
        }
    }
    if ($paragraph -ne "") {
        $output += $paragraph.Trim()
    }

    # Join paragraphs with newlines and copy to clipboard
    $formattedText = $output -join "`n`n"
    $formattedText | Set-Clipboard

    Remove-Item "transcript.en.srt"
    Write-Host "Sentences joined and copied to clipboard successfully."
} else {
    Write-Host "transcript.en.srt not found, skipping processing."
}